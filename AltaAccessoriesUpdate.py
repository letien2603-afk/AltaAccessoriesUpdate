import streamlit as st
import pandas as pd
import re
import io

# =====================================================================
# HÀM XỬ LÝ CHUỖI VÀ LATEST INVOICES
# =====================================================================
def clean_invoice_series(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(r'\.0$', '', regex=True)
    base = s.str.replace(r'[\s\-]*(COR|REV)\d*$', '', regex=True, flags=re.IGNORECASE)
    cleaned = base.str.lstrip('0')
    return cleaned.where(cleaned != '', base).replace('nan', pd.NA)

def get_invoice_rank(inv):
    if pd.isna(inv): return 0
    inv_str = str(inv).upper().strip()
    match = re.search(r'([\s\-]*)(COR|REV)(\d*)$', inv_str)
    if not match: return 0
    
    separator = match.group(1)
    type_str = match.group(2)
    num_str = match.group(3)
    
    type_val = 2 if type_str == 'COR' else 1
    num_val = int(num_str) if num_str else 1
    dash_count = separator.count('-')
    return (num_val * 100) + (type_val * 10) + dash_count

def increment_or_append_suffix(val, suffix_type):
    if pd.isnull(val): return val
    val_str = str(val).strip()
    match = re.search(r'(?i)(.*?)(?:[\s\-]*)(COR|REV)(\d*)$', val_str)
    if match:
        prefix = match.group(1).rstrip('- ')
        num_str = match.group(3)
        current_num = int(num_str) if num_str else 1
        next_num = current_num + 1
        return f"{prefix}-{suffix_type}{next_num}"
    else:
        return f"{val_str.rstrip('- ')}-{suffix_type}"

def replace_cor_with_rev(val):
    if pd.isnull(val): return val
    val = str(val)
    return re.sub(r'COR(\d*)$', r'REV\1', val)

# =====================================================================
# CHƯƠNG TRÌNH CHÍNH (STREAMLIT)
# =====================================================================
def main():
    st.set_page_config(page_title="ATF & Finance Processor", layout="wide")
    st.title("ATF & Finance Data Processing Tool")

    if 'processed' not in st.session_state:
        st.session_state.processed = False
        st.session_state.ytd_excel = None
        st.session_state.output_excel = None
        st.session_state.output_csv = None

    # --- BƯỚC 1: NHẬP THÔNG TIN VÀ UPLOAD FILES ---
    st.subheader("1. Required Information & Uploads")
    
    # [YÊU CẦU 1]: Thêm ô nhập Comment
    user_comment = st.text_input("Enter Comment for COR/REV updates:", placeholder="Nhập comment tại đây...")
    
    col1, col2 = st.columns(2)
    with col1:
        atf_file = st.file_uploader("Upload ATF File", type=['xlsx', 'xls', 'xlsb', 'csv'])
    with col2:
        finance_file = st.file_uploader("Upload Finance File", type=['xlsx', 'xls', 'xlsb'])

    if st.button("Start Processing", type="primary"):
        if not atf_file or not finance_file:
            st.error("Vui lòng upload đầy đủ ATF File và Finance File!")
            return

        progress_bar = st.progress(5, text="Đang đọc file ATF...")
        
        try:
            if atf_file.name.endswith('.csv'):
                df_atf = pd.read_csv(atf_file)
            else:
                df_atf = pd.read_excel(atf_file)
            
            # --- BƯỚC 2: QUÉT SIÊU TỐC FINANCE FILE ---
            progress_bar.progress(15, text="Đang quét nhanh Finance file...")
            fin_xls = pd.ExcelFile(finance_file)
            
            has_part_no = False
            has_sales_order = False
            part_col_name = None
            sales_col_name = None
            df_fin = None

            for sheet in fin_xls.sheet_names:
                temp_header = pd.read_excel(fin_xls, sheet_name=sheet, nrows=0)
                temp_header.columns = temp_header.columns.astype(str).str.replace("'", "", regex=False).str.strip()
                cols_lower = temp_header.columns.str.lower()
                
                part_matches = cols_lower.str.contains('part no|part number', regex=True, na=False)
                sales_matches = cols_lower.str.contains('sales order', regex=True, na=False)
                
                if part_matches.any() and sales_matches.any():
                    has_part_no = True
                    has_sales_order = True
                    part_col_name = temp_header.columns[part_matches].tolist().pop(0)
                    sales_col_name = temp_header.columns[sales_matches].tolist().pop(0)
                    
                    df_fin = pd.read_excel(fin_xls, sheet_name=sheet)
                    df_fin.columns = df_fin.columns.astype(str).str.replace("'", "", regex=False).str.strip()
                    break

            if not (has_part_no and has_sales_order):
                st.error("Either Part Number or Sales Order column is not available in Finance file")
                progress_bar.empty()
                st.stop()

            fin_part_list = df_fin[part_col_name].dropna().astype(str).str.strip().tolist()
            fin_sales_list = df_fin[sales_col_name].dropna().astype(str).str.strip().str.lstrip('0').tolist()

            # --- BƯỚC 3 & 4: MAPPING VÀ FILTER ATF ---
            progress_bar.progress(30, text="Đang matching Product ID và Sales Order...")
            if 'Product ID' in df_atf.columns:
                df_atf['Product ID Match'] = df_atf['Product ID'].astype(str).str.strip().isin(fin_part_list)
            else:
                df_atf['Product ID Match'] = False
                
            if 'Order Number' in df_atf.columns:
                atf_orders = df_atf['Order Number'].astype(str).str.strip().str.lstrip('0')
                df_atf['Sales Order Match'] = atf_orders.isin(fin_sales_list)
            else:
                df_atf['Sales Order Match'] = False

            progress_bar.progress(40, text="Đang lọc dữ liệu ATF...")
            filtered_atf = df_atf.copy()
            if 'Process Code' in filtered_atf.columns:
                cond_process = filtered_atf['Process Code'].astype(str).str.strip().str.upper() == 'CCREC'
            else:
                cond_process = False

            filtered_atf = filtered_atf[
                (cond_process) & 
                (filtered_atf['Product ID Match'] == True) & 
                (filtered_atf['Sales Order Match'] == True)
            ].copy()

            # --- BƯỚC 8: LẤY LATEST INVOICES ---
            progress_bar.progress(55, text="Đang lấy Latest Invoices...")
            if 'Invoice Number' in filtered_atf.columns:
                filtered_atf['Original Invoice'] = clean_invoice_series(filtered_atf['Invoice Number'])
                filtered_atf['SortKey'] = filtered_atf['Invoice Number'].apply(get_invoice_rank)
                
                if 'Transaction Amount' in filtered_atf.columns:
                    filtered_atf['Temp_Amount'] = pd.to_numeric(filtered_atf['Transaction Amount'], errors='coerce').round(2).abs()
                    max_keys = filtered_atf.groupby(['Original Invoice', 'Temp_Amount'], dropna=False)['SortKey'].transform('max')
                    filtered_atf.drop(columns=['Temp_Amount'], inplace=True)
                else:
                    max_keys = filtered_atf.groupby('Original Invoice')['SortKey'].transform('max')
                
                latest_invoices = filtered_atf[filtered_atf['SortKey'] == max_keys].copy()
                latest_invoices.drop(columns=['Original Invoice', 'SortKey'], inplace=True, errors='ignore')
            else:
                latest_invoices = filtered_atf.copy()

            # --- BƯỚC 6: TẠO PIVOT TABLE TỪ LATEST INVOICES ---
            progress_bar.progress(65, text="Đang tạo Pivot Table...")
            if not latest_invoices.empty and 'Order Number' in latest_invoices.columns and 'Product Description' in latest_invoices.columns:
                pivot_val = 'Transaction Amount' if 'Transaction Amount' in latest_invoices.columns else 'Order Number'
                agg_function = 'sum' if 'Transaction Amount' in latest_invoices.columns else 'size'
                
                summary_pt = pd.pivot_table(
                    latest_invoices, 
                    values=pivot_val, 
                    index=['Order Number'], 
                    columns=['Product Description'], 
                    aggfunc=agg_function,
                    fill_value=0,
                    margins=True,                 
                    margins_name='Grand Total'    
                ).reset_index()
            else:
                summary_pt = pd.DataFrame({"Message": ["No data available for Pivot Table"]})

            # --- BƯỚC 5 & 7 & [YÊU CẦU 2]: LƯU YTD ALTA ACCESSORIES REVIEW ---
            progress_bar.progress(70, text="Đang lưu YTD Alta Accessories Review...")
            ytd_buffer = io.BytesIO()
            with pd.ExcelWriter(ytd_buffer, engine='openpyxl') as writer:
                filtered_atf.drop(columns=['Original Invoice', 'SortKey'], inplace=True, errors='ignore')
                
                # Sheet 1: Toàn bộ dữ liệu lọc
                filtered_atf.to_excel(writer, sheet_name='YTD ATF', index=False)
                # Sheet 2: Pivot Table
                summary_pt.to_excel(writer, sheet_name='Summary', index=False)
                # Sheet 3: Danh sách Latest Invoices (Yêu cầu 2)
                latest_invoices.to_excel(writer, sheet_name='YTD ATF-Latest Inv', index=False)
            st.session_state.ytd_excel = ytd_buffer.getvalue()

            # --- BƯỚC 10 & 11: LỌC "CAMERAS" VÀ TẠO COR/REV ---
            progress_bar.progress(75, text="Đang xử lý dữ liệu COR và REV...")
            if 'Product Description' in latest_invoices.columns:
                latest_invoices = latest_invoices[latest_invoices['Product Description'].astype(str).str.strip().str.lower() == 'cameras'].copy()

            df_cor = latest_invoices.copy()
            df_rev = latest_invoices.copy()

            # --- BƯỚC 12: XỬ LÝ COR ---
            if 'Transaction Number' in df_cor.columns:
                df_cor['Transaction Number'] = df_cor['Transaction Number'].apply(lambda x: increment_or_append_suffix(x, 'COR'))
            if 'Invoice Number' in df_cor.columns:
                df_cor['Invoice Number'] = df_cor['Invoice Number'].apply(lambda x: increment_or_append_suffix(x, 'COR'))
            df_cor['Transaction Type'] = "MANUAL_ADJ"
            if 'Product Description' in df_cor.columns:
                df_cor['Product Description'] = "AVA-Cameras"
            if 'Product Line' in df_cor.columns:
                df_cor['Product Line'] = "Alta Video"
                
            # [YÊU CẦU 3C]: Cập nhật Comment cho sheet COR
            if user_comment:
                if 'Comments' in df_cor.columns:
                    df_cor['Comments'] = user_comment
                elif 'Comment' in df_cor.columns:
                    df_cor['Comment'] = user_comment

            # --- BƯỚC 13: XỬ LÝ REV ---
            if 'Transaction Number' in df_rev.columns:
                df_rev['Transaction Number'] = df_cor['Transaction Number'].apply(replace_cor_with_rev) 
            if 'Invoice Number' in df_rev.columns:
                df_rev['Invoice Number'] = df_cor['Invoice Number'].apply(replace_cor_with_rev)
            df_rev['Transaction Type'] = "MANUAL_ADJ"
            
            # [YÊU CẦU 3C]: Cập nhật Comment cho sheet REV
            if user_comment:
                if 'Comments' in df_rev.columns:
                    df_rev['Comments'] = user_comment
                elif 'Comment' in df_rev.columns:
                    df_rev['Comment'] = user_comment

            # [YÊU CẦU 3A, 3B]: Đảo dấu (Flip Sign) cho tất cả các loại tiền, bao gồm AUD và Native Value
            currency_cols = ['Transaction Amount', 'EUR Value', 'CAD Value', 'GBP Value', 'AUD Value', 'Native Value']
            for col in currency_cols:
                if col in df_rev.columns:
                    df_rev[col] = pd.to_numeric(df_rev[col], errors='coerce') * -1

            # --- GỘP COR VÀ REV ---
            progress_bar.progress(85, text="Đang gộp (Consolidate) và xuất file Output...")
            df_consolidated = pd.concat([df_cor, df_rev], ignore_index=True)

            # --- BƯỚC 9, 14, 15: LƯU FILE ---
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                # [YÊU CẦU 4]: Chỉ xuất duy nhất sheet Consolidated
                df_consolidated.to_excel(writer, sheet_name='Consolidated', index=False)
            st.session_state.output_excel = output_buffer.getvalue()

            # [YÊU CẦU 4]: Đảm bảo Output.csv có dữ liệu y hệt sheet Consolidated
            st.session_state.output_csv = df_consolidated.to_csv(index=False).encode('utf-8-sig')

            st.session_state.processed = True
            progress_bar.progress(100, text="Hoàn tất xử lý!")
            st.success("Tất cả các tệp đã sẵn sàng tải xuống!")

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
            progress_bar.empty()
            st.session_state.processed = False

    # --- HIỂN THỊ NÚT TẢI XUỐNG ---
    if st.session_state.processed:
        st.subheader("📥 Download Results")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        with col_btn1:
            st.download_button(
                label="Download YTD Alta Accessories Review.xlsx",
                data=st.session_state.ytd_excel,
                file_name="YTD Alta Accessories Review.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_btn2:
            st.download_button(
                label="Download Output.xlsx",
                data=st.session_state.output_excel,
                file_name="Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col_btn3:
            st.download_button(
                label="Download Output.csv",
                data=st.session_state.output_csv,
                file_name="Output.csv",
                mime="text/csv"
            )

if __name__ == "__main__":
    main()
