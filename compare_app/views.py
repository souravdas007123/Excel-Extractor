from django.shortcuts import render
import pandas as pd
import io
import base64
import time
import math

# Helper function to generate Base64 Excel URI for individual downloads
def generate_excel_uri(df, sheet_name="Data"):
    if df is None or df.empty:
        return None
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    b64_data = base64.b64encode(output.getvalue()).decode('utf-8')
    return f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_data}"

def excel_dashboard(request):
    context = {'active_tab': 'single'} 
    
    if request.method == 'POST':
        mode = request.POST.get('mode')
        context['active_tab'] = mode
        
        columns_input = request.POST.get('columns', 'Name,Phone,State')
        columns_to_check = [col.strip() for col in columns_input.split(',')]

        try:
            # ==========================================
            # MODE 1: SINGLE EXCEL FILE EXTRACT
            # ==========================================
            if mode == 'single':
                file_single = request.FILES.get('file_single')
                if not file_single:
                    raise ValueError("Please upload an Excel file.")

                remove_duplicates = request.POST.get('remove_duplicates') == 'yes'
                remove_blanks = request.POST.get('remove_blanks') == 'yes'
                remove_invalid_phones = request.POST.get('remove_invalid_phones') == 'yes'
                auto_fix = request.POST.get('auto_fix') == 'yes'

                # Use perf_counter for more accurate internal backend timing
                start_time = time.perf_counter()

                df = pd.read_excel(file_single, engine='calamine')
                total_rows = len(df)
                
                df.insert(0, 'Excel Row', df.index + 2)
                
                missing_cols = [col for col in columns_to_check if col not in df.columns]
                if missing_cols:
                    raise ValueError(f"Columns not found: {missing_cols}")

                total_blanks = int(df[columns_to_check].isnull().sum().sum())
                
                # FIX: Remove .0 from Phone numbers early so previews look clean
                phone_col = None
                for col in columns_to_check:
                    if 'phone' in col.lower() or 'mobile' in col.lower() or 'contact' in col.lower():
                        phone_col = col
                        break
                
                if phone_col:
                    # fillna('') is used so blanks don't become 'nan' strings
                    df[phone_col] = df[phone_col].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)

                df_check = df.copy()
                
                for col in columns_to_check:
                    df_check[col] = df_check[col].fillna('').astype(str).str.strip()
                    
                    if auto_fix:
                        if 'name' in col.lower() or 'state' in col.lower() or 'city' in col.lower():
                            df_check[col] = df_check[col].str.title()
                    else:
                        df_check[col] = df_check[col].str.lower()

                valid_indices = df.index
                
                blanks_df = pd.DataFrame()
                invalid_phones_df = pd.DataFrame()
                duplicates_df = pd.DataFrame()

                has_empty_cells = df_check[columns_to_check].eq('').any(axis=1)
                empty_cells_df = df.loc[has_empty_cells]

                # --- 1. BLANK DATA REMOVAL ---
                if remove_blanks:
                    is_blank = df_check[columns_to_check].eq('').any(axis=1)
                    blank_indices = df_check[is_blank].index
                    blanks_df = df.loc[blank_indices]
                    valid_indices = valid_indices.difference(blank_indices)
                    df_check = df_check.loc[valid_indices]

                # --- 2. SMART PHONE FIXING & REMOVAL ---
                if phone_col:
                    phone_str = df_check[phone_col].astype(str)
                    
                    # Step 1: Pehle saare spaces, dashes ya + signs hata lo, sirf numbers rakho
                    only_numbers = phone_str.str.replace(r'\D', '', regex=True)
                    
                    if auto_fix:
                        # Step 2: '91' ko SIRF tab hatao jab uske theek baad 10 digits hon (total 12 digit number)
                        only_numbers = only_numbers.str.replace(r'^91(?=\d{10}$)', '', regex=True)
                        
                        # Step 3: Starting ka '0' SIRF tab hatao jab uske theek baad 10 digits hon (total 11 digit number)
                        only_numbers = only_numbers.str.replace(r'^0+(?=\d{10}$)', '', regex=True)
                    
                    df_check.loc[valid_indices, phone_col] = only_numbers.loc[valid_indices]

                    if remove_invalid_phones:
                        is_invalid = (only_numbers.str.len() > 0) & (only_numbers.str.len() != 10)
                        invalid_indices = df_check[is_invalid].index
                        invalid_phones_df = df.loc[invalid_indices]
                        valid_indices = valid_indices.difference(invalid_indices)
                        df_check = df_check.loc[valid_indices]

                # --- 3. DUPLICATE DATA REMOVAL ---
                if remove_duplicates:
                    is_duplicate = df_check.duplicated(subset=columns_to_check, keep='first')
                    duplicate_indices = df_check[is_duplicate].index
                    duplicates_df = df.loc[duplicate_indices]
                    valid_indices = valid_indices.difference(duplicate_indices)

                # Final Fresh Data 
                fresh_df = df_check.loc[valid_indices]
                fresh_df_export = fresh_df.drop(columns=['Excel Row'])

                health_score = math.floor((len(fresh_df) / total_rows) * 100) if total_rows > 0 else 0
                
                if health_score >= 90: health_color = 'success'
                elif health_score >= 70: health_color = 'warning'
                else: health_color = 'danger'

                # Main File Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    fresh_df_export.to_excel(writer, sheet_name='Fresh Clean Data', index=False)
                    if not duplicates_df.empty:
                        duplicates_df.to_excel(writer, sheet_name='Found Duplicates', index=False)
                b64_excel = base64.b64encode(output.getvalue()).decode('utf-8')
                main_excel_uri = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}"

                end_time = time.perf_counter()
                raw_seconds = end_time - start_time
                mins, secs = int(raw_seconds // 60), round(raw_seconds % 60, 2)

                context.update({
                    'single_success': True,
                    's_health_score': health_score,
                    's_health_color': health_color,
                    's_total_rows': total_rows,
                    's_duplicate_count': len(duplicates_df),
                    's_blank_rows_removed': len(blanks_df),
                    's_invalid_phones': len(invalid_phones_df),
                    's_total_blanks': total_blanks,
                    's_fresh_count': len(fresh_df),
                    's_download_uri': main_excel_uri,
                    's_execution_time': f"{mins} min {secs} sec", # Backend time fallback
                    
                    's_duplicates_table': duplicates_df.head(100).to_html(classes='table table-warning table-striped mb-0', index=False, justify='left', na_rep='') if not duplicates_df.empty else None,
                    's_blanks_table': blanks_df.head(100).to_html(classes='table table-danger table-striped mb-0', index=False, justify='left', na_rep='') if not blanks_df.empty else None,
                    's_invalid_table': invalid_phones_df.head(100).to_html(classes='table table-info table-striped mb-0', index=False, justify='left', na_rep='') if not invalid_phones_df.empty else None,
                    's_empty_cells_table': empty_cells_df.head(100).to_html(classes='table table-secondary table-striped mb-0', index=False, justify='left', na_rep='') if not empty_cells_df.empty else None,
                    
                    # Individual Download Links
                    's_dup_uri': generate_excel_uri(duplicates_df, "Removed Duplicates"),
                    's_blank_uri': generate_excel_uri(blanks_df, "Blank Rows"),
                    's_invalid_uri': generate_excel_uri(invalid_phones_df, "Invalid Phones"),
                    's_empty_uri': generate_excel_uri(empty_cells_df, "Rows with Empty Cells")
                })

            # ==========================================
            # MODE 2: EXCEL COMPARE DATA
            # ==========================================
            elif mode == 'compare':
                file1 = request.FILES.get('file1')
                file2 = request.FILES.get('file2')
                if not file1 or not file2:
                    raise ValueError("Please upload both files.")

                start_time = time.perf_counter()

                df1 = pd.read_excel(file1, engine='calamine', usecols=columns_to_check)
                df2 = pd.read_excel(file2, engine='calamine', usecols=columns_to_check)

                df1.insert(0, 'File 1 Row', df1.index + 2)
                df2.insert(0, 'File 2 Row', df2.index + 2)
                
                phone_col = None
                for col in columns_to_check:
                    if 'phone' in col.lower() or 'mobile' in col.lower() or 'contact' in col.lower():
                        phone_col = col
                        break
                
                if phone_col:
                    if phone_col in df1.columns:
                        df1[phone_col] = df1[phone_col].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
                    if phone_col in df2.columns:
                        df2[phone_col] = df2[phone_col].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)

                total_blanks = int(df1.isnull().sum().sum() + df2.isnull().sum().sum())

                df1_clean, df2_clean = df1.copy(), df2.copy()
                for col in columns_to_check:
                    df1_clean[col] = df1_clean[col].fillna('').astype(str).str.strip().str.lower()
                    df2_clean[col] = df2_clean[col].fillna('').astype(str).str.strip().str.lower()

                duplicates = pd.merge(df1_clean, df2_clean, on=columns_to_check, how='inner')

                # Main Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if not duplicates.empty:
                        duplicates.to_excel(writer, sheet_name='Exact Duplicates', index=False)
                    else:
                        pd.DataFrame(['No duplicates found'], columns=['Message']).to_excel(writer, index=False)
                b64_excel = base64.b64encode(output.getvalue()).decode('utf-8')
                main_excel_uri = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}"

                end_time = time.perf_counter()
                raw_seconds = end_time - start_time
                mins, secs = int(raw_seconds // 60), round(raw_seconds % 60, 2)

                context.update({
                    'compare_success': True,
                    'c_total_rows_1': len(df1),
                    'c_total_rows_2': len(df2),
                    'c_duplicate_count': len(duplicates),
                    'c_total_blanks': total_blanks,
                    'c_download_uri': main_excel_uri,
                    'c_execution_time': f"{mins} min {secs} sec", # Backend time fallback
                    'c_duplicates_table': duplicates.head(100).to_html(classes='table table-success table-striped mb-0', index=False, justify='left', na_rep='') if not duplicates.empty else None,
                    'c_dup_uri': generate_excel_uri(duplicates, "Exact Duplicates")
                })

        except Exception as e:
            context['error'] = str(e)

    return render(request, 'compare_app/compare.html', context)