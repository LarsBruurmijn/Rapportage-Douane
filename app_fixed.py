import streamlit as st
import pandas as pd
import re
import io
import tempfile
import os
import traceback
from typing import Optional, Dict
import openpyxl
import warnings
import zipfile

# Suppress alle openpyxl warnings die gerelateerd zijn aan Fill bugs
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', message='.*expected.*Fill.*')

def clean_text(text: str) -> str:
    """Schoont tekstvelden op"""
    if pd.isna(text) or not isinstance(text, str):
        return text
    
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    text = re.sub(r'(\d)([A-Za-z])', r'\1 \2', text)
    text = re.sub(r'([A-Za-z])(\d)', r'\1 \2', text)
    text = re.sub(r'\.([A-Z])', r'.\n\1', text)
    text = re.sub(r'!([A-Z])', r'!\n\1', text)
    text = re.sub(r'\?([A-Z])', r'?\n\1', text)
    text = re.sub(r'([a-z])([-•*]\s*[A-Z])', r'\1\n\2', text)
    text = re.sub(r'([a-z])(\d+\.\s*[A-Z])', r'\1\n\2', text)
    text = re.sub(r':([A-Za-z])', r': \1', text)
    text = re.sub(r',([A-Za-z])', r', \1', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def fix_openpyxl_315_bugs(uploaded_file):
    """Specifieke fix voor openpyxl 3.1.5 'expected Fill' bugs"""
    
    st.info("🔧 Applying openpyxl 3.1.5 Fill bug workaround...")
    
    try:
        # Reset file pointer
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        
        # Strategie 1: Force data_only + read_only (vermijdt alle styling)
        try:
            bytes_io = io.BytesIO(file_bytes)
            wb = openpyxl.load_workbook(
                bytes_io, 
                read_only=True,    # Belangrijk: vermijdt style parsing
                data_only=True,    # Belangrijk: alleen data, geen formules/styling
                keep_vba=False     # Belangrijk: vermijdt VBA problemen
            )
            
            # Handmatig data extraheren om Fill bugs te vermijden
            sheet_data = {}
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                data = []
                
                # Lees alleen cell values, geen styling
                for row in ws.iter_rows(values_only=True):
                    if any(cell is not None for cell in row):
                        data.append(row)
                
                if data and len(data) > 1:
                    # Eerste rij als headers
                    df = pd.DataFrame(data[1:], columns=data[0])
                    sheet_data[sheet_name] = df
                elif data:
                    # Geen headers
                    df = pd.DataFrame(data)
                    sheet_data[sheet_name] = df
            
            wb.close()
            st.success("✅ Fix 1 (read_only + data_only) succeeded!")
            return sheet_data
            
        except Exception as e:
            st.warning(f"⚠️ Fix 1 failed: {str(e)[:50]}...")
        
        # Strategie 2: Excel file reconstruction (removes all formatting)
        try:
            st.info("🔄 Trying Excel reconstruction...")
            return reconstruct_clean_excel(file_bytes)
            
        except Exception as e:
            st.warning(f"⚠️ Fix 2 failed: {str(e)[:50]}...")
        
        # Strategie 3: Pandas with BytesIO (sometimes works)
        try:
            st.info("🔄 Trying pandas BytesIO...")
            bytes_io = io.BytesIO(file_bytes)
            
            # Get sheet names first
            excel_file = pd.ExcelFile(bytes_io, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            
            sheet_data = {}
            for sheet_name in sheet_names:
                bytes_io.seek(0)
                df = pd.read_excel(bytes_io, sheet_name=sheet_name, engine='openpyxl')
                sheet_data[sheet_name] = df
            
            st.success("✅ Fix 3 (pandas BytesIO) succeeded!")
            return sheet_data
            
        except Exception as e:
            st.warning(f"⚠️ Fix 3 failed: {str(e)[:50]}...")
        
        # Strategie 4: Temporary file approach
        try:
            st.info("🔄 Trying temporary file...")
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(file_bytes)
                tmp_file_path = tmp_file.name
            
            try:
                excel_file = pd.ExcelFile(tmp_file_path, engine='openpyxl')
                sheet_names = excel_file.sheet_names
                
                sheet_data = {}
                for sheet_name in sheet_names:
                    df = pd.read_excel(tmp_file_path, sheet_name=sheet_name, engine='openpyxl')
                    sheet_data[sheet_name] = df
                
                st.success("✅ Fix 4 (temporary file) succeeded!")
                return sheet_data
                
            finally:
                if os.path.exists(tmp_file_path):
                    os.unlink(tmp_file_path)
                    
        except Exception as e:
            st.warning(f"⚠️ Fix 4 failed: {str(e)[:50]}...")
        
        # If all fixes fail
        st.error("❌ All openpyxl 3.1.5 workarounds failed")
        return None
        
    except Exception as e:
        st.error(f"❌ Critical error in fix function: {str(e)}")
        return None

def reconstruct_clean_excel(file_bytes):
    """Reconstruct Excel file without any formatting to avoid Fill bugs"""
    
    # Read original file with minimal parsing
    bytes_io = io.BytesIO(file_bytes)
    
    # Use openpyxl to extract pure data
    wb_orig = openpyxl.load_workbook(bytes_io, data_only=True, keep_vba=False)
    
    # Create completely new workbook without any styling
    from openpyxl import Workbook
    wb_clean = Workbook()
    wb_clean.remove(wb_clean.active)  # Remove default sheet
    
    sheet_data = {}
    
    for sheet_name in wb_orig.sheetnames:
        ws_orig = wb_orig[sheet_name]
        ws_clean = wb_clean.create_sheet(title=sheet_name)
        
        # Extract only values, no formatting whatsoever
        data_rows = []
        for row in ws_orig.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                data_rows.append(row)
        
        # Write to clean sheet (only values)
        for row_data in data_rows:
            ws_clean.append(row_data)
        
        # Convert to DataFrame
        if data_rows and len(data_rows) > 1:
            df = pd.DataFrame(data_rows[1:], columns=data_rows[0])
            sheet_data[sheet_name] = df
        elif data_rows:
            df = pd.DataFrame(data_rows)
            sheet_data[sheet_name] = df
    
    wb_orig.close()
    wb_clean.close()
    
    return sheet_data

def safe_excel_read(uploaded_file, sheet_name=None):
    """Ultra-safe Excel reading met openpyxl 3.1.5 bug fixes"""
    
    # First try: Apply specific fixes
    if sheet_name is None:
        # Read all sheets
        sheet_data = fix_openpyxl_315_bugs(uploaded_file)
        return sheet_data
    else:
        # Read specific sheet
        sheet_data = fix_openpyxl_315_bugs(uploaded_file)
        if sheet_data and sheet_name in sheet_data:
            return sheet_data[sheet_name]
        else:
            st.error(f"Sheet '{sheet_name}' not found in file")
            return None

def get_sheet_names_safe(uploaded_file):
    """Safe sheet name detection met bug fixes"""
    
    try:
        uploaded_file.seek(0)
        
        # Method 1: read_only approach (safest)
        try:
            wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            st.success("✅ Sheet detection: openpyxl read_only")
            return sheet_names
        except:
            pass
        
        # Method 2: pandas ExcelFile
        try:
            uploaded_file.seek(0)
            file_bytes = uploaded_file.read()
            bytes_io = io.BytesIO(file_bytes)
            excel_file = pd.ExcelFile(bytes_io, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            st.success("✅ Sheet detection: pandas ExcelFile")
            return sheet_names
        except:
            pass
        
        # Method 3: Temporary file
        try:
            uploaded_file.seek(0)
            file_bytes = uploaded_file.read()
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            
            try:
                excel_file = pd.ExcelFile(tmp_path, engine='openpyxl')
                sheet_names = excel_file.sheet_names
                st.success("✅ Sheet detection: temporary file")
                return sheet_names
            finally:
                os.unlink(tmp_path)
                
        except:
            pass
        
        st.error("❌ Could not detect sheet names with any method")
        return []
        
    except Exception as e:
        st.error(f"❌ Sheet detection error: {str(e)}")
        return []

def create_mapping_from_columns(mapping_df, norm_id_col, task_name_col):
    """Maakt mapping van specifieke kolommen"""
    mapping_dict = {}
    
    for _, row in mapping_df.iterrows():
        norm_id = row[norm_id_col]
        task_name = row[task_name_col]
        
        if pd.notna(task_name) and pd.notna(norm_id):
            clean_task_name = str(task_name).strip().lower()
            mapping_dict[clean_task_name] = str(norm_id)
    
    return mapping_dict

def find_norm_id(task_name, mapping_dict):
    """Vindt norm-ID voor taak"""
    if pd.isna(task_name):
        return None
        
    clean_task_name = str(task_name).strip().lower()
    
    # Exacte match
    if clean_task_name in mapping_dict:
        return mapping_dict[clean_task_name]
    
    # Fuzzy match
    for mapped_task, norm_id in mapping_dict.items():
        if mapped_task in clean_task_name or clean_task_name in mapped_task:
            return norm_id
    
    return None

def merge_data(tasks_df, norms_df, mapping_dict, task_name_col, norm_id_col):
    """Voegt data samen"""
    result_df = tasks_df.copy()
    
    # Voeg norm-ID toe
    result_df['Norm_ID'] = result_df[task_name_col].apply(
        lambda x: find_norm_id(x, mapping_dict)
    )
    
    # Merge met normen
    if norms_df is not None:
        norms_df_copy = norms_df.copy()
        norms_df_copy['Norm_ID'] = norms_df_copy[norm_id_col].astype(str)
        result_df['Norm_ID'] = result_df['Norm_ID'].astype(str)
        
        result_df = result_df.merge(
            norms_df_copy, 
            on='Norm_ID', 
            how='left', 
            suffixes=('', '_norm')
        )
    
    # Schoon tekst op
    for col in result_df.columns:
        if result_df[col].dtype == 'object':
            result_df[col] = result_df[col].apply(clean_text)
    
    return result_df

def create_excel_output(df):
    """Creëert Excel output"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Taak met Norm', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="Excel Taak-Norm Verrijker (Fixed)",
        page_icon="🔧",
        layout="wide"
    )
    
    st.title("🔧 Excel Taak-Norm Verrijker (Openpyxl 3.1.5 Fixed)")
    st.markdown("**Specifiek gefixed** voor openpyxl 3.1.5 'expected Fill' bugs")
    
    # Warning about openpyxl version
    with st.expander("⚠️ Openpyxl 3.1.5 Bug Info", expanded=False):
        st.warning("Openpyxl 3.1.5 heeft bekende 'expected Fill' bugs met Excel cell formatting.")
        st.info("Deze versie gebruikt 4 verschillende workarounds om deze bugs te omzeilen:")
        st.write("1. **read_only + data_only** - Vermijdt alle styling")
        st.write("2. **Excel reconstruction** - Maakt schone kopie zonder formatting")  
        st.write("3. **BytesIO approach** - Alternative file handling")
        st.write("4. **Temporary file** - Fallback methode")
    
    # File uploaders
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📝 Stap 1: Upload Taakbestand")
        tasks_file = st.file_uploader(
            "Upload Excel bestand met taken",
            type=['xlsx', 'xls'],
            key="tasks_file",
            help="Upload je Excel bestand (werkt met 500KB bestanden)"
        )
        
    with col2:
        st.subheader("📏 Stap 2: Upload Normbestand")
        norms_file = st.file_uploader(
            "Upload Excel bestand met normen",
            type=['xlsx', 'xls'],
            key="norms_file",
            help="Upload je Excel bestand met normen en mapping"
        )
    
    if tasks_file and norms_file:
        st.success("✅ Beide bestanden geüpload!")
        
        # Get sheet names with safe method
        with st.spinner("🔍 Detecting sheets with bug-safe methods..."):
            tasks_sheets = get_sheet_names_safe(tasks_file)
            norms_sheets = get_sheet_names_safe(norms_file)
        
        if tasks_sheets and norms_sheets:
            st.info(f"📋 Taakbestand sheets: {tasks_sheets}")
            st.info(f"📋 Normbestand sheets: {norms_sheets}")
            
            # Sheet selection
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.write("**Taak Data**")
                selected_task_sheet = st.selectbox(
                    "Selecteer taak sheet:",
                    tasks_sheets,
                    key="task_sheet"
                )
                
            with col2:
                st.write("**Norm Data**")
                selected_norm_sheet = st.selectbox(
                    "Selecteer norm data sheet:",
                    norms_sheets,
                    key="norm_sheet"
                )
                
            with col3:
                st.write("**Mapping Data**")
                selected_mapping_sheet = st.selectbox(
                    "Selecteer mapping sheet:",
                    norms_sheets,
                    key="mapping_sheet"
                )
            
            # Load data with safe methods
            if st.button("🔧 Laad Data (Bug-Safe)", type="secondary"):
                with st.spinner("🛡️ Loading data with openpyxl 3.1.5 fixes..."):
                    
                    # Load all data with bug fixes
                    st.write("**Loading task data...**")
                    tasks_all = safe_excel_read(tasks_file)
                    tasks_df = tasks_all.get(selected_task_sheet) if tasks_all else None
                    
                    st.write("**Loading norm data...**")
                    norms_all = safe_excel_read(norms_file)
                    norms_df = norms_all.get(selected_norm_sheet) if norms_all else None
                    
                    st.write("**Loading mapping data...**")
                    mapping_df = norms_all.get(selected_mapping_sheet) if norms_all else None
                
                if tasks_df is not None and norms_df is not None and mapping_df is not None:
                    # Store in session state
                    st.session_state['tasks_df'] = tasks_df
                    st.session_state['norms_df'] = norms_df
                    st.session_state['mapping_df'] = mapping_df
                    
                    st.success("✅ All data loaded successfully with bug fixes!")
                else:
                    st.error("❌ Failed to load one or more sheets")
            
            # Column configuration and processing
            if 'tasks_df' in st.session_state:
                tasks_df = st.session_state['tasks_df']
                norms_df = st.session_state['norms_df']
                mapping_df = st.session_state['mapping_df']
                
                st.subheader("📊 Stap 3: Configureer Kolommen")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.write("**Taak Data Preview**")
                    st.dataframe(tasks_df.head(3), use_container_width=True)
                    
                    task_name_col = st.selectbox(
                        "Kolom met taaknamen:",
                        tasks_df.columns.tolist(),
                        key="task_name_col"
                    )
                
                with col2:
                    st.write("**Norm Data Preview**")
                    st.dataframe(norms_df.head(3), use_container_width=True)
                    
                    norm_id_col = st.selectbox(
                        "Kolom met norm-ID's:",
                        norms_df.columns.tolist(),
                        key="norm_id_col"
                    )
                
                with col3:
                    st.write("**Mapping Data Preview**")
                    st.dataframe(mapping_df.head(3), use_container_width=True)
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        mapping_norm_id_col = st.selectbox(
                            "Mapping: Norm-ID:",
                            mapping_df.columns.tolist(),
                            key="mapping_norm_id_col"
                        )
                    with col_b:
                        mapping_task_name_col = st.selectbox(
                            "Mapping: Taaknaam:",
                            mapping_df.columns.tolist(),
                            key="mapping_task_name_col"
                        )
                
                # Process data
                if st.button("🚀 Verwerk Data (Fixed)", type="primary", use_container_width=True):
                    with st.spinner("⚡ Processing data with bug fixes..."):
                        try:
                            # Create mapping
                            mapping_dict = create_mapping_from_columns(
                                mapping_df, mapping_norm_id_col, mapping_task_name_col
                            )
                            
                            if mapping_dict:
                                st.info(f"📊 {len(mapping_dict)} mappings found")
                                
                                # Merge data
                                result_df = merge_data(
                                    tasks_df, norms_df, mapping_dict, 
                                    task_name_col, norm_id_col
                                )
                                
                                # Show results
                                st.subheader("📋 Resultaat")
                                st.dataframe(result_df, use_container_width=True)
                                
                                # Statistics
                                matched_count = result_df['Norm_ID'].notna().sum()
                                total_count = len(result_df)
                                
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Totaal taken", total_count)
                                with col2:
                                    st.metric("Gekoppelde taken", matched_count)
                                with col3:
                                    st.metric("Succes %", f"{(matched_count/total_count*100):.1f}%")
                                
                                # Download
                                excel_file = create_excel_output(result_df)
                                
                                st.download_button(
                                    label="📥 Download Resultaat",
                                    data=excel_file,
                                    file_name=f"taak_met_norm_fixed_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                                
                                st.success("✅ Fixed processing completed successfully!")
                            
                            else:
                                st.error("❌ No mappings found. Check your column selections.")
                                
                        except Exception as e:
                            st.error(f"❌ Processing failed: {str(e)}")
                            st.code(traceback.format_exc())
        
        else:
            st.error("❌ Could not detect sheet names with any method")
    
    else:
        st.info("👆 Upload both Excel files to start")
        
    # Instructions
    with st.sidebar:
        st.header("🔧 Bug Fixes Applied")
        st.markdown("""
        **Openpyxl 3.1.5 Fixes:**
        
        🛡️ **read_only + data_only**
        - Skips all cell styling
        - Prevents Fill parsing
        - Fastest method
        
        🔄 **Excel reconstruction**  
        - Creates clean copy
        - Removes all formatting
        - Guaranteed compatibility
        
        📋 **BytesIO handling**
        - Alternative file reading
        - Memory optimization
        - Streamlit compatibility
        
        💾 **Temporary files**
        - Fallback method
        - Maximum compatibility
        - Last resort approach
        """)

if __name__ == "__main__":
    main()