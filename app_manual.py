import streamlit as st
import pandas as pd
import re
import io
from typing import Optional, Dict
import openpyxl

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

def simple_read_excel(file, sheet_name=None):
    """Eenvoudige Excel lezer zonder complexe error handling"""
    try:
        file.seek(0)
        if sheet_name:
            df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
        else:
            df = pd.read_excel(file, sheet_name=0, engine='openpyxl')
        return df
    except Exception as e:
        try:
            # Probeer zonder engine specificatie
            file.seek(0)
            if sheet_name:
                df = pd.read_excel(file, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file, sheet_name=0)
            return df
        except Exception as e2:
            st.error(f"Kan bestand niet lezen: {str(e2)}")
            return None

def get_sheet_names_simple(file):
    """Eenvoudige sheet naam detectie"""
    try:
        file.seek(0)
        excel_file = pd.ExcelFile(file)
        return excel_file.sheet_names
    except Exception as e:
        try:
            file.seek(0)
            wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception as e2:
            st.error(f"Kan sheet namen niet ophalen: {str(e2)}")
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

def merge_manual_data(tasks_df, norms_df, mapping_dict, task_name_col, norm_id_col):
    """Voegt data samen op basis van handmatige specificaties"""
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
        page_title="Excel Taak-Norm Verrijker (Handmatig)",
        page_icon="🔧",
        layout="wide"
    )
    
    st.title("🔧 Excel Taak-Norm Verrijker (Handmatige Configuratie)")
    st.markdown("**Volledige controle**: Specificeer zelf welke tabbladen en kolommen je wilt gebruiken.")
    
    # File uploaders
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📝 Stap 1: Upload Taakbestand")
        tasks_file = st.file_uploader(
            "Upload Excel bestand met taken",
            type=['xlsx', 'xls'],
            key="tasks_file"
        )
        
    with col2:
        st.subheader("📏 Stap 2: Upload Normbestand")
        norms_file = st.file_uploader(
            "Upload Excel bestand met normen",
            type=['xlsx', 'xls'],
            key="norms_file"
        )
    
    if tasks_file and norms_file:
        st.success("✅ Beide bestanden geüpload!")
        
        # Haal sheet namen op
        tasks_sheets = get_sheet_names_simple(tasks_file)
        norms_sheets = get_sheet_names_simple(norms_file)
        
        if tasks_sheets and norms_sheets:
            st.info(f"📋 Taakbestand sheets: {tasks_sheets}")
            st.info(f"📋 Normbestand sheets: {norms_sheets}")
            
            # Configuratie sectie
            st.subheader("⚙️ Stap 3: Configureer je selecties")
            
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
            
            # Laad preview data
            if st.button("🔍 Laad Preview van Geselecteerde Sheets"):
                with st.spinner("Data wordt geladen..."):
                    tasks_df = simple_read_excel(tasks_file, selected_task_sheet)
                    norms_df = simple_read_excel(norms_file, selected_norm_sheet)
                    mapping_df = simple_read_excel(norms_file, selected_mapping_sheet)
                
                if tasks_df is not None and norms_df is not None and mapping_df is not None:
                    # Bewaar in session state
                    st.session_state['tasks_df'] = tasks_df
                    st.session_state['norms_df'] = norms_df
                    st.session_state['mapping_df'] = mapping_df
                    
                    st.success("✅ Data geladen! Configureer nu de kolommen hieronder.")
                else:
                    st.error("❌ Fout bij laden van een of meer sheets")
            
            # Als data geladen is, toon kolom configuratie
            if 'tasks_df' in st.session_state:
                tasks_df = st.session_state['tasks_df']
                norms_df = st.session_state['norms_df']
                mapping_df = st.session_state['mapping_df']
                
                st.subheader("📊 Stap 4: Configureer Kolommen")
                
                # Toon previews
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
                            "Mapping: Norm-ID kolom:",
                            mapping_df.columns.tolist(),
                            key="mapping_norm_id_col"
                        )
                    with col_b:
                        mapping_task_name_col = st.selectbox(
                            "Mapping: Taaknaam kolom:",
                            mapping_df.columns.tolist(),
                            key="mapping_task_name_col"
                        )
                
                # Verwerk knop
                if st.button("🚀 Verwerk Data", type="primary", use_container_width=True):
                    with st.spinner("Data wordt verwerkt..."):
                        # Maak mapping
                        mapping_dict = create_mapping_from_columns(
                            mapping_df, mapping_norm_id_col, mapping_task_name_col
                        )
                        
                        if mapping_dict:
                            st.info(f"📊 {len(mapping_dict)} mappings gevonden")
                            
                            # Merge data
                            result_df = merge_manual_data(
                                tasks_df, norms_df, mapping_dict, 
                                task_name_col, norm_id_col
                            )
                            
                            # Toon resultaat
                            st.subheader("📋 Resultaat")
                            st.dataframe(result_df, use_container_width=True)
                            
                            # Statistieken
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
                                file_name=f"taak_met_norm_handmatig_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            st.success("✅ Verwerking voltooid!")
                        
                        else:
                            st.error("❌ Geen mappings gevonden. Controleer je kolom selecties.")
        
        else:
            st.error("❌ Kan sheet namen niet detecteren")
    
    else:
        st.info("👆 Upload beide Excel bestanden om te beginnen")
        
    # Instructies in sidebar
    with st.sidebar:
        st.header("📋 Hoe te gebruiken")
        st.markdown("""
        **Stap voor stap:**
        
        1. **Upload** beide Excel bestanden
        2. **Selecteer** de juiste tabbladen
        3. **Klik** "Laad Preview"
        4. **Configureer** welke kolommen wat bevatten
        5. **Verwerk** de data
        6. **Download** het resultaat
        
        **Voordelen:**
        - ✅ Volledige controle
        - ✅ Geen automatische aannames
        - ✅ Werkt met alle Excel formaten
        - ✅ Handmatige kolom mapping
        """)

if __name__ == "__main__":
    main()