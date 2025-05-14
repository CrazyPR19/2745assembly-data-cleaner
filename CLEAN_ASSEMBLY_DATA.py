import streamlit as st
import pandas as pd
import io
import os
import datetime

st.title("Clean & Enrich Assembly Sheet Excel")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Strip spaces from column names and values
    df.columns = [col.strip() for col in df.columns]
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # Fill down ASSEMBLY MARK
    if 'ASSEMBLY MARK' in df.columns:
        df['ASSEMBLY MARK'] = df['ASSEMBLY MARK'].replace('', pd.NA)
        df['ASSEMBLY MARK'] = df['ASSEMBLY MARK'].fillna(method='ffill')

    # Create ASSEMBLY DESC by extracting text before "-" in ASSEMBLY MARK
    if 'ASSEMBLY MARK' in df.columns:
        df['ASSEMBLY DESC'] = df['ASSEMBLY MARK'].astype(str).apply(lambda x: x.split('-')[0] if '-' in x else x)
        cols = df.columns.tolist()
        asm_index = cols.index('ASSEMBLY MARK')
        cols.insert(asm_index + 1, cols.pop(cols.index('ASSEMBLY DESC')))
        df = df[cols]

    # Fill up NMDC DWG NO from bottom
    if 'NMDC DWG NO' in df.columns:
        for i in range(len(df) - 1, 0, -1):
            if pd.isna(df.at[i - 1, 'NMDC DWG NO']) or str(df.at[i - 1, 'NMDC DWG NO']).strip() == '':
                if pd.notna(df.at[i, 'NMDC DWG NO']):
                    df.at[i - 1, 'NMDC DWG NO'] = df.at[i, 'NMDC DWG NO']

    # Drop the last row if NMDC DWG NO or STRUCTURE NAME is blank
    if not df.empty and 'NMDC DWG NO' in df.columns and 'STRUCTURE NAME' in df.columns:
        last_row = df.iloc[-1]
        if pd.isna(last_row['NMDC DWG NO']) or str(last_row['NMDC DWG NO']).strip() == '' \
           or pd.isna(last_row['STRUCTURE NAME']) or str(last_row['STRUCTURE NAME']).strip() == '':
            df = df.iloc[:-1]

    # Check required columns
    required_cols = ['NMDC DWG NO', 'PART MARK', 'QTY / PCS']
    for col in required_cols:
        if col not in df.columns:
            st.error(f"Required column '{col}' is missing.")
            st.stop()

    # Convert QTY / PCS to numeric
    df['QTY / PCS'] = pd.to_numeric(df['QTY / PCS'], errors='coerce')

    # Build Assembly Qty map before grouping
    headers = df[df['PART MARK'].isna() & df['QTY / PCS'].notna()]
    parent_qty_map = (
        headers.drop_duplicates(subset=['NMDC DWG NO', 'ASSEMBLY MARK'])
               .set_index(['NMDC DWG NO', 'ASSEMBLY MARK'])['QTY / PCS']
               .astype(float).to_dict()
    )

    # Group and sum QTY / PCS
    group_cols = [col for col in df.columns if col not in ['QTY / PCS']]
    df = df.groupby(['NMDC DWG NO', 'ASSEMBLY MARK', 'PART MARK'], dropna=False).agg({
        **{col: 'first' for col in group_cols if col not in ['NMDC DWG NO', 'ASSEMBLY MARK', 'PART MARK']},
        'QTY / PCS': 'sum'
    }).reset_index()

    # Create MEMBERPROFILE where DESCRIPTION / NAME is not blank
    df["MEMBERPROFILE"] = df.apply(
        lambda row: f"{row['DESCRIPTION / NAME'].strip()}-{row['PROFILE'].strip()}"
        if pd.notna(row["DESCRIPTION / NAME"]) and str(row["DESCRIPTION / NAME"]).strip() != ""
        else "", axis=1
    )
    if "PROFILE" in df.columns and "MEMBERPROFILE" in df.columns:
        cols = df.columns.tolist()
        profile_index = cols.index("PROFILE")
        cols.insert(profile_index + 1, cols.pop(cols.index("MEMBERPROFILE")))
        df = df[cols]

    # Create PieceMarkNo = 3M-NMDC_DWG_NO(without -M)-0PART_MARK
    def create_piece_mark(row):
        try:
            if pd.notna(row['NMDC DWG NO']) and pd.notna(row['PART MARK']):
                clean_dwg = str(row['NMDC DWG NO']).replace("-M", "")
                part_mark = str(int(float(row['PART MARK']))) if str(row['PART MARK']).replace('.0','').isdigit() else str(row['PART MARK'])
                return f"3M-{clean_dwg}-0{part_mark}"
            else:
                return ""
        except:
            return ""

    part_mark_index = df.columns.get_loc('PART MARK')
    df.insert(part_mark_index + 1, 'PieceMarkNo', df.apply(create_piece_mark, axis=1))

    # Insert Assembly Qty
    def compute_assembly_qty(row):
        try:
            key = (row['NMDC DWG NO'], row['ASSEMBLY MARK'])
            return parent_qty_map.get(key, 1.0)
        except:
            return 1.0

    insert_after = df.columns.get_loc('QTY / PCS') + 1
    df.insert(insert_after, 'Assembly Qty', df.apply(compute_assembly_qty, axis=1))

    # Insert Total Qty
    def get_total_qty(row):
        try:
            return float(row['QTY / PCS']) * float(row['Assembly Qty'])
        except:
            return 0.0

    df.insert(insert_after + 1, 'Total Qty', df.apply(get_total_qty, axis=1))

    # Insert Total Weight (KG)
    def get_total_weight(row):
        try:
            return round(float(row['Total Qty']) * float(row['UNIT WEIGHT (KG)']), 2)
        except:
            return 0.0

    if 'WEIGHT (KG)' in df.columns:
        weight_col_index = df.columns.get_loc('WEIGHT (KG)') + 1
        df.insert(weight_col_index, 'Total Weight (KG)', df.apply(get_total_weight, axis=1))

    # Create Assembly Master sheet ONLY from distinct headers
    assembly_master = df[df['STRUCTURE NAME'].isna() | (df['STRUCTURE NAME'].astype(str).str.strip() == '')][['NMDC DWG NO', 'ASSEMBLY MARK', 'ASSEMBLY DESC']].drop_duplicates()
    assembly_master['Assembly Qty'] = assembly_master.apply(
        lambda row: parent_qty_map.get((row['NMDC DWG NO'], row['ASSEMBLY MARK']), 1.0), axis=1
    )
    assembly_master['Ass_Location'] = 'YC'
    assembly_master['YIC_Code'] = 'FR'

    # Create Assembly Control Sheet
    control_sheet = df[(df['PART MARK'].notna()) & (df['STRUCTURE NAME'].notna()) & (df['STRUCTURE NAME'].astype(str).str.strip() != '')].copy()
    control_sheet['PROJECTID'] = 2745
    control_sheet['STRUCTURE_NO'] = control_sheet['STRUCTURE NAME']
    control_sheet['Length_mm'] = ''
    control_sheet['Width_mm'] = ''
    control_sheet['Height_mm'] = ''
    control_sheet['Grade'] = control_sheet['GRADE'] if 'GRADE' in control_sheet.columns else ''
    control_sheet['UNIT AREA (SQM) /LENGTH (L)'] = control_sheet['UNIT AREA (SQM) /LENGTH (L)'] if 'UNIT AREA (SQM) /LENGTH (L)' in control_sheet.columns else ''

    assembly_control = control_sheet[[
        'PROJECTID', 'STRUCTURE_NO', 'ASSEMBLY MARK', 'NMDC DWG NO', 'PieceMarkNo', 'MEMBERPROFILE',
        'Total Qty', 'Length_mm', 'Width_mm', 'Height_mm', 'Grade',
        'UNIT WEIGHT (KG)', 'Total Weight (KG)', 'UNIT AREA (SQM) /LENGTH (L)'
    ]]

    # Save to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='MainSheet')
        assembly_master.to_excel(writer, index=False, sheet_name='AssemblyMaster')
        assembly_control.to_excel(writer, index=False, sheet_name='AssemblyControlSheet')
    output.seek(0)

    # Dynamic filename
    original_name = os.path.splitext(uploaded_file.name)[0]
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    new_filename = f"{original_name}_AssemblySheets_{timestamp}.xlsx"

    st.download_button(
        label="Download Assembly Master & Control Sheets",
        data=output,
        file_name=new_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
