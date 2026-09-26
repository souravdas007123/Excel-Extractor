from django.shortcuts import render
import pandas as pd
import io
import base64
import time  # Naya Import

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

                # 🟢 TEENO CHECKBOXES KI VALUES
                remove_duplicates = request.POST.get('remove_duplicates') == 'yes'
                remove_blanks = request.POST.get('remove_blanks') == 'yes'
                remove_invalid_phones = request.POST.get('remove_invalid_phones') == 'yes'

                start_time = time.time()

                df = pd.read_excel(file_single, engine='calamine')
                
                missing_cols = [col for col in columns_to_check if col not in df.columns]
                if missing_cols:
                    raise ValueError(f"Columns not found: {missing_cols}")

                total_blanks = int(df[columns_to_check].isnull().sum().sum())
                
                df_check = df.copy()
                for col in columns_to_check:
                    df_check[col] = df_check[col].fillna('').astype(str).str.strip().str.lower()

                valid_indices = df.index
                blank_rows_count = 0
                duplicate_count = 0
                invalid_phone_count = 0
                duplicates_df = pd.DataFrame()

                # --- 1. BLANK DATA REMOVAL ---
                if remove_blanks:
                    is_blank = df_check[columns_to_check].eq('').any(axis=1)
                    blank_indices = df_check[is_blank].index
                    blank_rows_count = len(blank_indices)
                    valid_indices = valid_indices.difference(blank_indices)
                    df_check = df_check.loc[valid_indices]

                # --- 2. INVALID PHONE REMOVAL (Naya Logic) ---
                if remove_invalid_phones:
                    # Phone column dhoondo (Name mein phone, mobile ya contact ho)
                    phone_col = None
                    for col in columns_to_check:
                        if 'phone' in col.lower() or 'mobile' in col.lower() or 'contact' in col.lower():
                            phone_col = col
                            break
                    
                    if phone_col:
                        # STEP 1: Excel ka trailing '.0' hatao pehle
                        phone_str = df_check[phone_col].astype(str).str.replace(r'\.0$', '', regex=True)
                        
                        # STEP 2: Ab baaki ke non-digits (jaise -, +, spaces) hatao
                        only_numbers = phone_str.str.replace(r'\D', '', regex=True)
                        
                        # STEP 3: Agar digits 0 se zyada hain but 10 se kam hain, toh wo invalid hai
                        is_invalid = (only_numbers.str.len() > 0) & (only_numbers.str.len() != 10)
                        
                        invalid_indices = df_check[is_invalid].index
                        invalid_phone_count = len(invalid_indices)
                        
                        valid_indices = valid_indices.difference(invalid_indices)
                        df_check = df_check.loc[valid_indices]

                # --- 3. DUPLICATE DATA REMOVAL ---
                if remove_duplicates:
                    is_duplicate = df_check.duplicated(subset=columns_to_check, keep='first')
                    duplicate_indices = df_check[is_duplicate].index
                    
                    # Original dataframe se duplicate rows nikalna taaki report me dikha sakein
                    duplicates_df = df.loc[duplicate_indices]
                    duplicate_count = len(duplicates_df)
                    
                    valid_indices = valid_indices.difference(duplicate_indices)

                # Final Fresh Data based on surviving indices
                fresh_df = df.loc[valid_indices]

                # Excel Creation
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    fresh_df.to_excel(writer, sheet_name='Fresh Clean Data', index=False)
                    if not duplicates_df.empty:
                        duplicates_df.to_excel(writer, sheet_name='Found Duplicates', index=False)
                
                b64_excel = base64.b64encode(output.getvalue()).decode('utf-8')
                excel_uri = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}"

                end_time = time.time()
                raw_seconds = end_time - start_time
                mins = int(raw_seconds // 60)
                secs = round(raw_seconds % 60, 2)
                formatted_time = f"{mins} min {secs} sec"

                context.update({
                    'single_success': True,
                    's_duplicate_count': duplicate_count,
                    's_blank_rows_removed': blank_rows_count,
                    's_invalid_phones': invalid_phone_count, # 🟢 Naya variable for frontend
                    's_total_blanks': total_blanks,
                    's_fresh_count': len(fresh_df),
                    's_download_uri': excel_uri,
                    's_execution_time': formatted_time,
                    's_duplicates_table': duplicates_df.head(100).to_html(classes='table table-warning table-striped', index=False) if not duplicates_df.empty else None
                })

            # ==========================================
            # MODE 2: EXCEL COMPARE DATA
            # ==========================================
            elif mode == 'compare':
                file1 = request.FILES.get('file1')
                file2 = request.FILES.get('file2')
                if not file1 or not file2:
                    raise ValueError("Please upload both files.")

                start_time = time.time() # ⏱️ Timer Start

                df1 = pd.read_excel(file1, engine='calamine', usecols=columns_to_check)
                df2 = pd.read_excel(file2, engine='calamine', usecols=columns_to_check)

                total_blanks = int(df1.isnull().sum().sum() + df2.isnull().sum().sum())

                df1_clean = df1.copy()
                df2_clean = df2.copy()
                for col in columns_to_check:
                    df1_clean[col] = df1_clean[col].fillna('').astype(str).str.strip().str.lower()
                    df2_clean[col] = df2_clean[col].fillna('').astype(str).str.strip().str.lower()

                duplicates = pd.merge(df1_clean, df2_clean, on=columns_to_check, how='inner')
                duplicate_count = len(duplicates)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    if not duplicates.empty:
                        duplicates.to_excel(writer, sheet_name='Exact Duplicates', index=False)
                    else:
                        pd.DataFrame(['No duplicates found'], columns=['Message']).to_excel(writer, index=False)
                
                b64_excel = base64.b64encode(output.getvalue()).decode('utf-8')
                excel_uri = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}"

                end_time = time.time()
                raw_seconds = end_time - start_time
                
                # Naya logic Min aur Sec ke liye
                mins = int(raw_seconds // 60)
                secs = round(raw_seconds % 60, 2)
                formatted_time = f"{mins} min {secs} sec"

                context.update({
                    'compare_success': True,
                    'c_duplicate_count': duplicate_count,
                    'c_total_blanks': total_blanks,
                    'c_download_uri': excel_uri,
                    'c_execution_time': formatted_time,  # Ab yahan formatted time jayega
                    'c_duplicates_table': duplicates.head(100).to_html(classes='table table-success table-striped', index=False) if not duplicates.empty else None
                })

        except Exception as e:
            context['error'] = str(e)

    return render(request, 'compare_app/compare.html', context)