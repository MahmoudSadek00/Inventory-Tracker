import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Domanza Inventory Application", layout="wide")
st.title("📦 Domanza Inventory Application")

# Uploading files
products_file = st.file_uploader("Upload Products File (CSV or Excel)", type=['csv', 'xlsx', 'xls'])
schedule_file = st.file_uploader("Upload Schedule Sheet (CSV or Excel)", type=['csv', 'xlsx', 'xls'])

# Helper function to read file
def read_file(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)

if products_file and schedule_file:
    try:
        # Load data
        df = read_file(products_file)
        schedule_df = read_file(schedule_file)

        # Clean schedule sheet
        schedule_df = schedule_df.iloc[:, :3]
        schedule_df.columns = ['Branch', 'Date', 'Brand']
        schedule_df['Date'] = pd.to_datetime(schedule_df['Date'], errors='coerce')
        schedule_df = schedule_df.dropna(subset=['Date'])

        today = pd.to_datetime(datetime.today().date())
        today_schedule = schedule_df[schedule_df['Date'] == today]

        today_brands = today_schedule['Brand'].dropna().unique().tolist()
        today_branches = today_schedule['Branch'].dropna().unique().tolist()

        # Extract brand from name_ar
        df['brand'] = df['name_ar'].apply(lambda x: x.split('-')[0].strip() if pd.notnull(x) else "")

        # Extract category from name_ar (word after 3rd dash)
        df['Category'] = df['name_ar'].apply(
            lambda x: x.split('-')[3].strip() if pd.notnull(x) and len(x.split('-')) > 3 else ""
        )

        # Columns check
        columns_needed = ['brand', 'name_ar', 'barcodes', 'available_quantity', 'branch_name', 'Category']
        missing_cols = [col for col in columns_needed if col not in df.columns]

        if missing_cols:
            st.error(f"❌ Missing required columns: {missing_cols}")
        else:
            result_df = df[columns_needed].copy()

            result_df = result_df.rename(columns={
                'brand': 'Brand',
                'name_ar': 'Product Name',
                'barcodes': 'Barcodes',
                'available_quantity': 'Available Quantity',
                'branch_name': 'Branch'
            })

            # Filter by today’s brand and branch
            result_df = result_df[
                result_df['Brand'].isin(today_brands) &
                result_df['Branch'].isin(today_branches)
            ]

            result_df = result_df.sort_values(by='Product Name')
            result_df['Actual Quantity'] = ''

            st.success("✅ Data ready!")
            st.dataframe(result_df)

            if st.button("⬇️ Download Inventory Excel"):
                from io import BytesIO
                output = BytesIO()

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    header_format = workbook.add_format({'bold': True})

                    summary_data = []  # لتجميع الـ Summary

                    for brand in result_df['Brand'].unique():
                        brand_df = result_df[result_df['Brand'] == brand].copy()
                        brand_df = brand_df[['Branch', 'Brand', 'Product Name', 'Category', 'Barcodes', 'Available Quantity', 'Actual Quantity']]

                        sheet_name = brand[:31]
                        brand_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)

                        worksheet = writer.sheets[sheet_name]

                        for col_num, col_name in enumerate(brand_df.columns):
                            worksheet.write(0, col_num, col_name, header_format)
                            max_len = max(brand_df[col_name].astype(str).map(len).max(), len(col_name))
                            worksheet.set_column(col_num, col_num, max_len + 2)

                        row_count = len(brand_df)
                        worksheet.write(0, 7, 'Difference', header_format)
                        for row in range(1, row_count + 1):
                            formula = f"=G{row+1}-F{row+1}"
                            worksheet.write_formula(row, 7, formula)

                            # Add to summary
                            product_name = brand_df.iloc[row-1]['Product Name']
                            summary_data.append((product_name, formula))

                    # ===== Create Summary Sheet at the beginning =====
                    summary_sheet = workbook.add_worksheet('Summary')
                    summary_sheet.write(0, 0, 'Product Name', header_format)
                    summary_sheet.write(0, 1, 'Difference', header_format)

                    for idx, (product, formula) in enumerate(summary_data, start=1):
                        summary_sheet.write(idx, 0, product)
                        summary_sheet.write_formula(idx, 1, formula)

                    # Auto width for both columns
                    max_product_len = max([len(str(p)) for p, _ in summary_data] + [12])
                    summary_sheet.set_column(0, 0, max_product_len + 2)
                    summary_sheet.set_column(1, 1, 12)

                    # تأكد إن شيت Summary هو أول واحد
                    workbook.worksheets_objs.insert(0, workbook.worksheets_objs.pop())

                # Save file with branch name + date
                file_name = f"{today_branches[0]}_{today.strftime('%Y-%m-%d')}.xlsx"

                st.download_button(
                    label="📥 Download Final Excel File",
                    data=output.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
