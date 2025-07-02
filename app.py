import streamlit as st
import pandas as pd
import re
import io
from typing import Optional, Tuple, Dict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

def clean_text(text: str) -> str:
    """
    Schoont tekstvelden op door spaties toe te voegen waar woorden aan elkaar geplakt zijn
    en regeleinden te herstellen na punten of opsommingstekens.
    """
    if pd.isna(text) or not isinstance(text, str):
        return text
    
    # Vervang multiple spaces door single space
    text = re.sub(r'\s+', ' ', text)
    
    # Voeg spaties toe voor hoofdletters die midden in woorden staan (bijv. "woordWoord" -> "woord Woord")
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    
    # Voeg spaties toe na cijfers gevolgd door letters
    text = re.sub(r'(\d)([A-Za-z])', r'\1 \2', text)
    
    # Voeg spaties toe voor letters gevolgd door cijfers
    text = re.sub(r'([A-Za-z])(\d)', r'\1 \2', text)
    
    # Herstel regeleinden na punten gevolgd door hoofdletters (nieuwe zinnen)
    text = re.sub(r'\.([A-Z])', r'.\n\1', text)
    
    # Herstel regeleinden na uitroeptekens gevolgd door hoofdletters
    text = re.sub(r'!([A-Z])', r'!\n\1', text)
    
    # Herstel regeleinden na vraagtekens gevolgd door hoofdletters
    text = re.sub(r'\?([A-Z])', r'?\n\1', text)
    
    # Herstel regeleinden voor opsommingstekens
    text = re.sub(r'([a-z])([-•*]\s*[A-Z])', r'\1\n\2', text)
    
    # Herstel regeleinden voor genummerde lijsten
    text = re.sub(r'([a-z])(\d+\.\s*[A-Z])', r'\1\n\2', text)
    
    # Voeg spaties toe na dubbele punten als die ontbreken
    text = re.sub(r':([A-Za-z])', r': \1', text)
    
    # Voeg spaties toe na komma's als die ontbreken
    text = re.sub(r',([A-Za-z])', r', \1', text)
    
    # Clean up extra spaces
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    return text

def read_excel_file(uploaded_file) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
    """
    Leest een Excel bestand in en retourneert de eerste twee tabbladen.
    """
    try:
        # Lees beide tabbladen
        sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        
        # Probeer het tweede tabblad te lezen
        try:
            sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
        except:
            sheet2 = None
            
        return sheet1, sheet2
    except Exception as e:
        st.error(f"Fout bij het lezen van het Excel bestand: {str(e)}")
        return None, None

def create_mapping_dict(mapping_df: pd.DataFrame) -> Dict[str, str]:
    """
    Creëert een mapping dictionary van taaknaam naar norm-ID.
    """
    mapping_dict = {}
    
    if mapping_df is not None and len(mapping_df.columns) >= 2:
        # Gebruik de eerste twee kolommen voor mapping
        norm_id_col = mapping_df.columns[0]
        task_name_col = mapping_df.columns[1]
        
        for _, row in mapping_df.iterrows():
            norm_id = row[norm_id_col]
            task_name = row[task_name_col]
            
            if pd.notna(task_name) and pd.notna(norm_id):
                # Clean de taaknaam voor betere matching
                clean_task_name = str(task_name).strip().lower()
                mapping_dict[clean_task_name] = str(norm_id)
    
    return mapping_dict

def find_matching_norm_id(task_name: str, mapping_dict: Dict[str, str]) -> Optional[str]:
    """
    Vindt de bijbehorende norm-ID voor een gegeven taaknaam.
    """
    if pd.isna(task_name):
        return None
        
    clean_task_name = str(task_name).strip().lower()
    
    # Exacte match
    if clean_task_name in mapping_dict:
        return mapping_dict[clean_task_name]
    
    # Fuzzy match - zoek naar gedeeltelijke overeenkomsten
    for mapped_task, norm_id in mapping_dict.items():
        if mapped_task in clean_task_name or clean_task_name in mapped_task:
            return norm_id
    
    return None

def merge_data(tasks_df: pd.DataFrame, norms_df: pd.DataFrame, mapping_dict: Dict[str, str], task_name_column: str) -> pd.DataFrame:
    """
    Voegt normgegevens toe aan het taakbestand op basis van de mapping.
    """
    result_df = tasks_df.copy()
    
    # Voeg norm-ID kolom toe
    result_df['Norm_ID'] = result_df[task_name_column].apply(
        lambda x: find_matching_norm_id(x, mapping_dict)
    )
    
    # Merge met normgegevens
    if norms_df is not None and 'Norm_ID' in result_df.columns:
        # Probeer de norm-ID kolom in het norms_df te vinden
        norm_id_column = None
        for col in norms_df.columns:
            if 'id' in str(col).lower() or 'norm' in str(col).lower():
                norm_id_column = col
                break
        
        if norm_id_column is None and len(norms_df.columns) > 0:
            norm_id_column = norms_df.columns[0]
        
        if norm_id_column:
            # Merge op basis van norm-ID
            norms_df_copy = norms_df.copy()
            norms_df_copy['Norm_ID'] = norms_df_copy[norm_id_column].astype(str)
            result_df['Norm_ID'] = result_df['Norm_ID'].astype(str)
            
            result_df = result_df.merge(
                norms_df_copy, 
                on='Norm_ID', 
                how='left', 
                suffixes=('', '_norm')
            )
    
    # Schoon alle tekstvelden op
    for col in result_df.columns:
        if result_df[col].dtype == 'object':  # String kolommen
            result_df[col] = result_df[col].apply(clean_text)
    
    return result_df

def create_excel_output(df: pd.DataFrame) -> io.BytesIO:
    """
    Creëert een Excel bestand van de output data.
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Taak met Norm', index=False)
        
        # Format de Excel sheet
        workbook = writer.book
        worksheet = writer.sheets['Taak met Norm']
        
        # Header formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # Write headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-adjust column widths
        for i, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.set_column(i, i, min(max_length + 2, 50))
    
    output.seek(0)
    return output

def main():
    st.set_page_config(
        page_title="Excel Taak-Norm Verrijker",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel Taak-Norm Verrijker")
    st.markdown("Upload twee Excel-bestanden om taken te verrijken met normgegevens en tekst op te schonen.")
    
    # Sidebar voor instructies
    with st.sidebar:
        st.header("📋 Instructies")
        st.markdown("""
        **Bestandsformaat:**
        
        **Taakbestand:**
        - Excel bestand met taken
        - Selecteer de kolom met taaknamen
        
        **Normbestand:**
        - **Tabblad 1:** Normgegevens
        - **Tabblad 2:** Mapping (kolom 1: Norm-ID, kolom 2: Taaknaam)
        
        **Functionaliteit:**
        - Automatische koppeling van taken aan normen
        - Tekst opschoning (spaties en regeleinden)
        - Excel export met gecombineerde data
        """)
    
    # File uploaders
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📝 Taakbestand")
        tasks_file = st.file_uploader(
            "Upload het Excel bestand met taken",
            type=['xlsx', 'xls'],
            key="tasks_file"
        )
        
    with col2:
        st.subheader("📏 Normbestand")
        norms_file = st.file_uploader(
            "Upload het Excel bestand met normen (2 tabbladen)",
            type=['xlsx', 'xls'],
            key="norms_file"
        )
    
    if tasks_file and norms_file:
        # Lees bestanden in
        with st.spinner("Bestanden worden ingelezen..."):
            tasks_df, _ = read_excel_file(tasks_file)
            norms_df, mapping_df = read_excel_file(norms_file)
        
        if tasks_df is not None and norms_df is not None and mapping_df is not None:
            st.success("✅ Beide bestanden succesvol ingelezen!")
            
            # Toon preview van data
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("👀 Preview Taakbestand")
                st.dataframe(tasks_df.head(), use_container_width=True)
                
                # Laat gebruiker taaknaam kolom selecteren
                task_name_column = st.selectbox(
                    "Selecteer de kolom met taaknamen:",
                    options=tasks_df.columns.tolist(),
                    key="task_name_col"
                )
            
            with col2:
                st.subheader("👀 Preview Normgegevens")
                st.dataframe(norms_df.head(), use_container_width=True)
                
                st.subheader("🗺️ Preview Mapping")
                st.dataframe(mapping_df.head(), use_container_width=True)
            
            # Genereer uitvoerbestand knop
            if st.button("🚀 Genereer Uitvoerbestand", type="primary", use_container_width=True):
                with st.spinner("Gegevens worden verwerkt en tekst wordt opgeschoond..."):
                    # Creëer mapping dictionary
                    mapping_dict = create_mapping_dict(mapping_df)
                    
                    if mapping_dict:
                        st.info(f"📊 {len(mapping_dict)} mappings gevonden tussen taken en normen")
                        
                        # Merge data
                        result_df = merge_data(tasks_df, norms_df, mapping_dict, task_name_column)
                        
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
                            st.metric("Koppelingspercentage", f"{(matched_count/total_count*100):.1f}%")
                        
                        # Download knop
                        excel_file = create_excel_output(result_df)
                        
                        st.download_button(
                            label="📥 Download Resultaat Excel",
                            data=excel_file,
                            file_name=f"taak_met_norm_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        st.success("✅ Verwerking voltooid! Alle tekstvelden zijn opgeschoond en taken zijn gekoppeld aan normen.")
                    
                    else:
                        st.error("❌ Geen geldige mapping gevonden in het tweede tabblad van het normbestand.")
        
        else:
            st.error("❌ Fout bij het inlezen van één of beide bestanden. Controleer of de bestanden geldig zijn en het normbestand twee tabbladen heeft.")
    
    else:
        st.info("👆 Upload beide Excel-bestanden om te beginnen.")

if __name__ == "__main__":
    main()