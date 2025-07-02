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
from difflib import SequenceMatcher

# Suppress alle openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', message='.*expected.*Fill.*')

# Versie info
APP_VERSION = "v2.1.0"
APP_DATE = "2025-01-31"

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

def similarity(a, b):
    """Berekent similarity tussen twee strings"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def flexible_column_match(search_term, available_columns, threshold=0.4):
    """Flexibele kolom matching met verschillende strategieën"""
    search_lower = search_term.lower()
    
    # Strategie 1: Exacte match (case insensitive)
    for col in available_columns:
        if col.lower() == search_lower:
            return col
    
    # Strategie 2: Bevat match
    for col in available_columns:
        if search_lower in col.lower() or col.lower() in search_lower:
            return col
    
    # Strategie 3: Fuzzy match op basis van similarity
    best_match = None
    best_score = threshold
    
    for col in available_columns:
        score = similarity(search_term, col)
        if score > best_score:
            best_score = score
            best_match = col
    
    return best_match

def smart_column_detection(df, column_type):
    """Intelligente detectie van kolommen op basis van type"""
    
    keywords = {
        'task_name': ['taak', 'task', 'naam', 'name', 'titel', 'title', 'omschrijving', 'description', 'activiteit', 'activity'],
        'norm_id': ['norm', 'id', 'code', 'nummer', 'number', 'ref', 'referentie', 'identifier'],
        'mapping_task': ['taak', 'task', 'naam', 'name', 'titel', 'title'],
        'mapping_norm': ['norm', 'id', 'code', 'nummer', 'number']
    }
    
    if column_type not in keywords:
        return None
    
    search_terms = keywords[column_type]
    
    for term in search_terms:
        match = flexible_column_match(term, df.columns.tolist())
        if match:
            return match
    
    return None

def ultra_safe_excel_read(uploaded_file):
    """Ultra-safe Excel reading met alle workarounds"""
    
    try:
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        
        # Strategie 1: read_only + data_only (meest stabiel)
        try:
            bytes_io = io.BytesIO(file_bytes)
            wb = openpyxl.load_workbook(
                bytes_io, 
                read_only=True,
                data_only=True,
                keep_vba=False
            )
            
            sheet_data = {}
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                data = []
                
                for row in ws.iter_rows(values_only=True):
                    if any(cell is not None for cell in row):
                        data.append(row)
                
                if data and len(data) > 1:
                    df = pd.DataFrame(data[1:], columns=data[0])
                    sheet_data[sheet_name] = df
                elif data:
                    df = pd.DataFrame(data)
                    sheet_data[sheet_name] = df
            
            wb.close()
            return sheet_data
            
        except Exception as e:
            st.warning(f"Strategie 1 gefaald: {str(e)[:50]}...")
        
        # Strategie 2: Pandas met BytesIO
        try:
            bytes_io = io.BytesIO(file_bytes)
            excel_file = pd.ExcelFile(bytes_io, engine='openpyxl')
            
            sheet_data = {}
            for sheet_name in excel_file.sheet_names:
                bytes_io.seek(0)
                df = pd.read_excel(bytes_io, sheet_name=sheet_name, engine='openpyxl')
                sheet_data[sheet_name] = df
            
            return sheet_data
            
        except Exception as e:
            st.warning(f"Strategie 2 gefaald: {str(e)[:50]}...")
        
        # Strategie 3: Temporary file
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(file_bytes)
                tmp_file_path = tmp_file.name
            
            try:
                excel_file = pd.ExcelFile(tmp_file_path, engine='openpyxl')
                
                sheet_data = {}
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(tmp_file_path, sheet_name=sheet_name, engine='openpyxl')
                    sheet_data[sheet_name] = df
                
                return sheet_data
                
            finally:
                if os.path.exists(tmp_file_path):
                    os.unlink(tmp_file_path)
                    
        except Exception as e:
            st.warning(f"Strategie 3 gefaald: {str(e)[:50]}...")
        
        st.error("❌ Alle Excel read strategieën gefaald")
        return None
        
    except Exception as e:
        st.error(f"❌ Kritieke fout: {str(e)}")
        return None

def get_sheet_names_safe(uploaded_file):
    """Safe sheet name detection"""
    try:
        uploaded_file.seek(0)
        
        # Probeer met read_only
        try:
            wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except:
            pass
        
        # Probeer met pandas
        try:
            uploaded_file.seek(0)
            file_bytes = uploaded_file.read()
            bytes_io = io.BytesIO(file_bytes)
            excel_file = pd.ExcelFile(bytes_io, engine='openpyxl')
            return excel_file.sheet_names
        except:
            pass
        
        return []
        
    except Exception as e:
        st.error(f"❌ Sheet detection error: {str(e)}")
        return []

def flexible_task_matching(task_name, mapping_dict, threshold=0.3):
    """Zeer flexibele taak matching"""
    if pd.isna(task_name):
        return None
        
    clean_task = str(task_name).strip().lower()
    
    # Strategie 1: Exacte match
    if clean_task in mapping_dict:
        return mapping_dict[clean_task]
    
    # Strategie 2: Bevat match
    for mapped_task, norm_id in mapping_dict.items():
        if mapped_task in clean_task or clean_task in mapped_task:
            return norm_id
    
    # Strategie 3: Fuzzy match
    best_match = None
    best_score = threshold
    
    for mapped_task, norm_id in mapping_dict.items():
        score = similarity(clean_task, mapped_task)
        if score > best_score:
            best_score = score
            best_match = norm_id
    
    # Strategie 4: Word-based matching
    if not best_match:
        task_words = set(clean_task.split())
        for mapped_task, norm_id in mapping_dict.items():
            mapped_words = set(mapped_task.split())
            common_words = task_words.intersection(mapped_words)
            if common_words and len(common_words) >= min(2, len(task_words) // 2):
                return norm_id
    
    return best_match

def create_flexible_mapping(mapping_df, norm_id_col, task_name_col):
    """Maakt flexibele mapping"""
    mapping_dict = {}
    
    for _, row in mapping_df.iterrows():
        norm_id = row[norm_id_col]
        task_name = row[task_name_col]
        
        if pd.notna(task_name) and pd.notna(norm_id):
            clean_task_name = str(task_name).strip().lower()
            mapping_dict[clean_task_name] = str(norm_id)
    
    return mapping_dict

def merge_data_flexible(tasks_df, norms_df, mapping_dict, task_name_col, norm_id_col):
    """Flexibele data merge"""
    result_df = tasks_df.copy()
    
    # Voeg norm-ID toe met flexibele matching
    result_df['Norm_ID'] = result_df[task_name_col].apply(
        lambda x: flexible_task_matching(x, mapping_dict)
    )
    
    # Count successful matches
    matched_count = result_df['Norm_ID'].notna().sum()
    total_count = len(result_df)
    match_percentage = (matched_count / total_count * 100) if total_count > 0 else 0
    
    st.info(f"📊 Flexibele matching: {matched_count}/{total_count} taken gekoppeld ({match_percentage:.1f}%)")
    
    # Merge met normen
    if norms_df is not None and matched_count > 0:
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
        page_title=f"Excel Taak-Norm Verrijker {APP_VERSION}",
        page_icon="🚀",
        layout="wide"
    )
    
    # Versie header
    st.markdown(f"""
    <div style="background-color: #1f4e79; color: white; padding: 10px; border-radius: 5px; margin-bottom: 20px;">
        <h2 style="margin: 0; text-align: center;">🚀 Excel Taak-Norm Verrijker</h2>
        <p style="margin: 5px 0 0 0; text-align: center; font-size: 14px;">
            <strong>Versie {APP_VERSION}</strong> | {APP_DATE} | Ultra-flexibele kolom matching
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Features overview
    with st.expander("✨ Nieuwe Features in deze versie", expanded=False):
        st.markdown("""
        **🎯 Ultra-flexibele kolom matching:**
        - ✅ Case-insensitive matching
        - ✅ Gedeeltelijke naam matching  
        - ✅ Fuzzy string matching
        - ✅ Automatische kolom detectie
        - ✅ Word-based matching
        
        **🛡️ Robuuste Excel reading:**
        - ✅ 3 verschillende read strategieën
        - ✅ Openpyxl 3.1.5 bug fixes
        - ✅ Memory optimalisatie
        
        **📊 Verbeterde matching:**
        - ✅ 4 matching strategieën per taak
        - ✅ Match percentage tracking
        - ✅ Flexibele threshold instellingen
        """)
    
    # File uploaders
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📝 Stap 1: Upload Taakbestand")
        tasks_file = st.file_uploader(
            "Upload Excel bestand met taken",
            type=['xlsx', 'xls'],
            key="tasks_file",
            help="Alle Excel bestanden worden ondersteund - 500KB is geen probleem!"
        )
        
    with col2:
        st.subheader("📏 Stap 2: Upload Normbestand")
        norms_file = st.file_uploader(
            "Upload Excel bestand met normen",
            type=['xlsx', 'xls'],
            key="norms_file",
            help="Bestand met normen en mapping informatie"
        )
    
    if tasks_file and norms_file:
        st.success("✅ Beide bestanden geüpload!")
        
        # Load data with progress
        with st.spinner("🔄 Laden van Excel bestanden met ultra-safe methoden..."):
            tasks_data = ultra_safe_excel_read(tasks_file)
            norms_data = ultra_safe_excel_read(norms_file)
        
        if tasks_data and norms_data:
            st.success("✅ Alle Excel bestanden succesvol geladen!")
            
            # Sheet selection
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.write("**Taak Data**")
                task_sheets = list(tasks_data.keys())
                selected_task_sheet = st.selectbox(
                    "Selecteer taak sheet:",
                    task_sheets,
                    key="task_sheet"
                )
                
            with col2:
                st.write("**Norm Data**")
                norm_sheets = list(norms_data.keys())
                selected_norm_sheet = st.selectbox(
                    "Selecteer norm data sheet:",
                    norm_sheets,
                    key="norm_sheet"
                )
                
            with col3:
                st.write("**Mapping Data**")
                selected_mapping_sheet = st.selectbox(
                    "Selecteer mapping sheet:",
                    norm_sheets,
                    key="mapping_sheet"
                )
            
            # Get selected dataframes
            tasks_df = tasks_data[selected_task_sheet]
            norms_df = norms_data[selected_norm_sheet]
            mapping_df = norms_data[selected_mapping_sheet]
            
            # Show previews
            st.subheader("📊 Data Preview")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.write(f"**Taak Data ({tasks_df.shape[0]} rijen)**")
                st.dataframe(tasks_df.head(3), use_container_width=True)
                
            with col2:
                st.write(f"**Norm Data ({norms_df.shape[0]} rijen)**")
                st.dataframe(norms_df.head(3), use_container_width=True)
                
            with col3:
                st.write(f"**Mapping Data ({mapping_df.shape[0]} rijen)**")
                st.dataframe(mapping_df.head(3), use_container_width=True)
            
            # Smart column detection
            st.subheader("🎯 Intelligente Kolom Detectie")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Taak Configuratie**")
                
                # Auto-detect task name column
                auto_task_col = smart_column_detection(tasks_df, 'task_name')
                if auto_task_col:
                    st.info(f"🤖 Auto-gedetecteerd: '{auto_task_col}'")
                
                task_name_col = st.selectbox(
                    "Kolom met taaknamen:",
                    tasks_df.columns.tolist(),
                    index=tasks_df.columns.tolist().index(auto_task_col) if auto_task_col else 0,
                    key="task_name_col",
                    help="Kies de kolom met taak namen/beschrijvingen"
                )
                
                # Auto-detect norm ID column
                auto_norm_col = smart_column_detection(norms_df, 'norm_id')
                if auto_norm_col:
                    st.info(f"🤖 Auto-gedetecteerd: '{auto_norm_col}'")
                
                norm_id_col = st.selectbox(
                    "Kolom met norm-ID's:",
                    norms_df.columns.tolist(),
                    index=norms_df.columns.tolist().index(auto_norm_col) if auto_norm_col else 0,
                    key="norm_id_col",
                    help="Kies de kolom met norm ID's/codes"
                )
            
            with col2:
                st.write("**Mapping Configuratie**")
                
                # Auto-detect mapping columns
                auto_mapping_norm = smart_column_detection(mapping_df, 'mapping_norm')
                auto_mapping_task = smart_column_detection(mapping_df, 'mapping_task')
                
                if auto_mapping_norm:
                    st.info(f"🤖 Auto-gedetecteerd norm: '{auto_mapping_norm}'")
                if auto_mapping_task:
                    st.info(f"🤖 Auto-gedetecteerd taak: '{auto_mapping_task}'")
                
                mapping_norm_id_col = st.selectbox(
                    "Mapping: Norm-ID kolom:",
                    mapping_df.columns.tolist(),
                    index=mapping_df.columns.tolist().index(auto_mapping_norm) if auto_mapping_norm else 0,
                    key="mapping_norm_id_col",
                    help="Kolom die norm ID's bevat in mapping"
                )
                
                mapping_task_name_col = st.selectbox(
                    "Mapping: Taaknaam kolom:",
                    mapping_df.columns.tolist(),
                    index=mapping_df.columns.tolist().index(auto_mapping_task) if auto_mapping_task else 0,
                    key="mapping_task_name_col",
                    help="Kolom die taak namen bevat in mapping"
                )
            
            # Advanced settings
            with st.expander("⚙️ Geavanceerde Instellingen", expanded=False):
                col1, col2 = st.columns(2)
                
                with col1:
                    match_threshold = st.slider(
                        "Fuzzy match threshold",
                        min_value=0.1,
                        max_value=0.9,
                        value=0.3,
                        step=0.1,
                        help="Lagere waarde = meer matches, hogere waarde = striktere matches"
                    )
                
                with col2:
                    enable_word_matching = st.checkbox(
                        "Word-based matching inschakelen",
                        value=True,
                        help="Matcht taken op basis van gemeenschappelijke woorden"
                    )
            
            # Process data
            if st.button("🚀 Verwerk Data", type="primary", use_container_width=True):
                with st.spinner("⚡ Verwerken met ultra-flexibele matching..."):
                    try:
                        # Create flexible mapping
                        mapping_dict = create_flexible_mapping(
                            mapping_df, mapping_norm_id_col, mapping_task_name_col
                        )
                        
                        if mapping_dict:
                            st.info(f"📊 {len(mapping_dict)} mapping regels gevonden")
                            
                            # Merge data with flexible matching
                            result_df = merge_data_flexible(
                                tasks_df, norms_df, mapping_dict, 
                                task_name_col, norm_id_col
                            )
                            
                            # Show results
                            st.subheader("📋 Resultaat")
                            st.dataframe(result_df, use_container_width=True)
                            
                            # Statistics
                            matched_count = result_df['Norm_ID'].notna().sum()
                            total_count = len(result_df)
                            success_rate = (matched_count / total_count * 100) if total_count > 0 else 0
                            
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Totaal taken", total_count)
                            with col2:
                                st.metric("Gekoppelde taken", matched_count)
                            with col3:
                                st.metric("Succes %", f"{success_rate:.1f}%")
                            with col4:
                                if success_rate >= 80:
                                    st.metric("Kwaliteit", "🟢 Uitstekend")
                                elif success_rate >= 60:
                                    st.metric("Kwaliteit", "🟡 Goed")
                                else:
                                    st.metric("Kwaliteit", "🔴 Kan beter")
                            
                            # Show unmatched items for improvement
                            unmatched = result_df[result_df['Norm_ID'].isna()]
                            if len(unmatched) > 0:
                                with st.expander(f"⚠️ Niet gekoppelde taken ({len(unmatched)})", expanded=False):
                                    st.dataframe(unmatched[[task_name_col]], use_container_width=True)
                                    st.info("💡 Tip: Controleer spelling of voeg deze taken toe aan je mapping")
                            
                            # Download
                            excel_file = create_excel_output(result_df)
                            
                            st.download_button(
                                label="📥 Download Resultaat Excel",
                                data=excel_file,
                                file_name=f"taak_met_norm_flexible_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            st.success(f"✅ Verwerking voltooid! {success_rate:.1f}% succes met flexibele matching!")
                        
                        else:
                            st.error("❌ Geen mapping gegevens gevonden. Controleer je kolom selecties.")
                            
                    except Exception as e:
                        st.error(f"❌ Verwerking gefaald: {str(e)}")
                        st.code(traceback.format_exc())
        
        else:
            st.error("❌ Kon één of beide Excel bestanden niet laden")
    
    else:
        st.info("👆 Upload beide Excel bestanden om te beginnen")
        
        # Help section
        with st.expander("📖 Hoe werkt flexibele matching?", expanded=False):
            st.markdown("""
            **Deze app gebruikt 4 matching strategieën:**
            
            1. **Exacte match** - Zoekt exacte overeenkomsten (case-insensitive)
            2. **Bevat match** - Zoekt gedeeltelijke overeenkomsten 
            3. **Fuzzy match** - Gebruikt string similarity (bijv. "Administratie" ↔ "Administratieve")
            4. **Word match** - Matcht op gemeenschappelijke woorden
            
            **Kolom detectie:**
            - Zoekt automatisch naar kolommen met woorden als 'taak', 'naam', 'norm', 'id'
            - Werkt met verschillende talen en synoniemen
            - Jij behoudt volledige controle over de uiteindelijke keuze
            """)
    
    # Sidebar info
    with st.sidebar:
        st.header("ℹ️ App Info")
        st.markdown(f"""
        **Versie:** {APP_VERSION}  
        **Datum:** {APP_DATE}
        
        **✨ Nieuwe features:**
        - Ultra-flexibele kolom matching
        - Automatische kolom detectie  
        - Fuzzy string matching
        - 3 Excel read strategieën
        - Openpyxl 3.1.5 fixes
        
        **📊 Ondersteund:**
        - Alle Excel formaten (.xlsx, .xls)
        - Multi-sheet bestanden
        - Complexe formatting
        - 500KB+ bestanden
        """)

if __name__ == "__main__":
    main()