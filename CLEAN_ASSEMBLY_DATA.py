import streamlit as st
import pandas as pd
import io

st.title("Clean & Enrich Assembly Sheet Excel")

uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["xlsx", "csv"])

if uploaded_file is not None:
    # Read the file based on extension
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
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

        # Reorder columns: place ASSEMBLY DESC right after ASSEMBLY MARK
        cols = df.columns.tolist()
        asm_index = cols.index('ASSEMBLY MARK')
        cols.insert(asm_index + 1, cols.pop(cols.index('ASSEMBLY DESC')))
        df = df[cols]

    # Fill up NMDC DWG NO
    if 'NMDC DWG NO' in df.columns:
        for i in range(len(df) - 1, 0, -1):
            if pd.isna(df.at[i - 1, 'NMDC DWG NO']) or str(df.at[i - 1, 'NMDC DWG NO']).strip() == '':
                if pd.notna(df.at[i, 'NMDC DWG NO']):
                    df.at[i - 1, 'NMDC DWG NO'] = df.at[i, 'NMDC DWG NO']

    # Ensure NMDC DWG NO and PART MARK exist
    if 'NMDC DWG NO' not in df.columns or 'PART MARK' not in df.columns:
        st.error("Required columns 'NMDC DWG NO' or 'PART MARK' are missing.")
        st.stop()

    # Create MEMBERPROFILE column only where DESCRIPTION / NAME is not blank
    df["MEMBERPROFILE"] = df.apply(
        lambda row: f"{row['DESCRIPTION / NAME'].strip()}-{row['PROFILE'].strip()}"
        if pd.notna(row["DESCRIPTION / NAME"]) and str(row["DESCRIPTION / NAME"]).strip() != ""
        else "", axis=1
    )
    # Reorder columns to place MEMBERPROFILE after PROFILE
    cols = df.columns.tolist()
    if "PROFILE" in cols and "MEMBERPROFILE" in cols:
        profile_index = cols.index("PROFILE")
        # Move MEMBERPROFILE next to PROFILE
        cols.insert(profile_index + 1, cols.pop(cols.index("MEMBERPROFILE")))
        df = df[cols]

    # Create PieceMarkNo as concatenation of NMDC DWG NO + "-0" + PART MARK
    def create_piece_mark(row):
        try:
            if pd.notna(row['NMDC DWG NO']) and pd.notna(row['PART MARK']):
                return f"{row['NMDC DWG NO']}-0{row['PART MARK']}"
            else:
                return ""
        except:
            return ""

    # Insert 'PieceMarkNo' after 'PART MARK'
    part_mark_index = df.columns.get_loc('PART MARK')
    df.insert(part_mark_index + 1, 'PieceMarkNo', df.apply(create_piece_mark, axis=1))

    # Clean key fields
    key_cols = ['NMDC DWG NO', 'ASSEMBLY MARK', 'STRUCTURE NAME']
    for col in key_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Clean and convert QTY / PCS
    if 'QTY / PCS' not in df.columns:
        st.error("'QTY / PCS' column is missing.")
        st.stop()
    df['QTY / PCS'] = pd.to_numeric(df['QTY / PCS'], errors='coerce')

    # STEP 1: Build a mapping from (NMDC DWG NO, ASSEMBLY MARK) → QTY / PCS
    parent_qty_map = (
        df.dropna(subset=['NMDC DWG NO', 'ASSEMBLY MARK', 'QTY / PCS'])
          .drop_duplicates(subset=['NMDC DWG NO', 'ASSEMBLY MARK'])
          .set_index(['NMDC DWG NO', 'ASSEMBLY MARK'])['QTY / PCS']
          .astype(float)
          .to_dict()
    )

    # STEP 2: Assign "Assembly Qty"
    def compute_assembly_qty(row):
        try:
            if pd.isna(row['PART MARK']):
                return float(row['QTY / PCS'])  # Use its own value
            else:
                key = (row['NMDC DWG NO'], row['ASSEMBLY MARK'])
                return parent_qty_map.get(key, 1.0)
        except:
            return 1.0

    insert_after = df.columns.get_loc('QTY / PCS') + 1
    df.insert(insert_after, 'Assembly Qty', df.apply(compute_assembly_qty, axis=1))

    # STEP 3: Create "Total Qty"
    def get_total_qty(row):
        try:
            return float(row['QTY / PCS']) * float(row['Assembly Qty'])
        except:
            return 0.0

    df.insert(insert_after + 1, 'Total Qty', df.apply(get_total_qty, axis=1))

    # STEP 4: Create "Total Weight (KG)"
    def get_total_weight(row):
        try:
            return round(float(row['Total Qty']) * float(row['UNIT WEIGHT (KG)']), 2)
        except:
            return 0.0

    weight_col_index = df.columns.get_loc('WEIGHT (KG)') + 1
    df.insert(weight_col_index, 'Total Weight (KG)', df.apply(get_total_weight, axis=1))

    # Convert to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    # Streamlit download button
    st.download_button(
        label="Download Cleaned Excel File",
        data=output,
        file_name="cleaned_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
