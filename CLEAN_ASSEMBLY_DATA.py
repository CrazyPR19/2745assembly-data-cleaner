import streamlit as st
import pandas as pd
import io

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
    # Insert columns
    insert_after = df.columns.get_loc('QTY / PCS') + 1
    df.insert(insert_after, 'Assembly Qty', df.apply(compute_assembly_qty, axis=1))

    # ============== STEP 2: Create "Total Qty" =====================
    def get_total_qty(row):
        try:
            return float(row['QTY / PCS']) * float(row['Assembly Qty'])
        except:
            return 0.0

    df.insert(insert_after + 1, 'Total Qty', df.apply(get_total_qty, axis=1))

    # ============== STEP 3: Create "Total Weight (KG)" =====================
    def get_total_weight(row):
        try:
            return round(float(row['Total Qty']) * float(row['UNIT WEIGHT (KG)']), 2)
        except:
            return 0.0

    weight_col_index = df.columns.get_loc('WEIGHT (KG)') + 1
    df.insert(weight_col_index, 'Total Weight (KG)', df.apply(get_total_weight, axis=1))
    
    

    # Save and export
#     output = io.BytesIO()
#     df.to_excel(output, index=False)
# 
#     st.success("File cleaned and enriched successfully!")
#     st.download_button("Download Cleaned File", data=output.getvalue(), file_name="Cleaned_Assembly_Sheet_Final.xlsx")

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




