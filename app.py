import streamlit as st
import pandas as pd
import re
import io
from typing import Optional, Tuple, Dict
import openpyxl
from excel_cleaner import clean_excel_file

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
    Gebruikt meerdere strategieën om verschillende Excel formaten te ondersteunen.
    """
    # Probeer verschillende engines en methoden
    engines_to_try = ['openpyxl', 'xlrd', None]
    
    for engine in engines_to_try:
        try:
            # Reset file pointer
            uploaded_file.seek(0)
            
            # Lees het eerste tabblad
            if engine:
                sheet1 = pd.read_excel(uploaded_file, sheet_name=0, engine=engine)
            else:
                sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
            
            # Reset file pointer voor tweede lezing
            uploaded_file.seek(0)
            
            # Probeer het tweede tabblad te lezen
            sheet2 = None
            try:
                if engine:
                    sheet2 = pd.read_excel(uploaded_file, sheet_name=1, engine=engine)
                else:
                    sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
            except:
                # Geen tweede tabblad of fout bij lezen
                pass
            
            # Als we hier zijn, was het succesvol
            if engine:
                st.success(f"✅ Bestand gelezen met {engine} engine")
                st.info(f"📊 Sheet 1: {len(sheet1)} rijen, {len(sheet1.columns)} kolommen")
                if sheet2 is not None:
                    st.info(f"📊 Sheet 2: {len(sheet2)} rijen, {len(sheet2.columns)} kolommen")
            else:
                st.success(f"✅ Bestand gelezen met standaard engine")
            
            return sheet1, sheet2
            
        except Exception as e:
            if engine:
                st.warning(f"Engine {engine} mislukt: {str(e)}")
            else:
                st.warning(f"Standaard engine mislukt: {str(e)}")
            continue
    
    # Als alle engines falen, probeer Excel cleaning
    st.info("🔧 Alle standaard methodes gefaald. Probeer Excel bestand te zuiveren...")
    
    try:
        uploaded_file.seek(0)
        cleaned_file, sheet_names = clean_excel_file(uploaded_file)
        
        if cleaned_file and sheet_names:
            st.info(f"📋 Gevonden sheets: {sheet_names}")
            
            # Probeer het gezuiverde bestand te lezen
            sheet1 = pd.read_excel(cleaned_file, sheet_name=0, engine='openpyxl')
            
            # Probeer tweede sheet
            sheet2 = None
            if len(sheet_names) > 1:
                cleaned_file.seek(0)
                try:
                    sheet2 = pd.read_excel(cleaned_file, sheet_name=1, engine='openpyxl')
                except:
                    pass
            
            st.success("✅ Bestand succesvol gezuiverd en gelezen!")
            return sheet1, sheet2
        
    except Exception as e:
        st.error(f"Excel zuivering mislukt: {str(e)}")
    
    # Laatste poging: via BytesIO met tijdelijk bestand
    try:
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        
        # Schrijf naar tijdelijk bestand
        import tempfile
        import os
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(file_bytes)
            tmp_file_path = tmp_file.name
        
        try:
            # Probeer via tijdelijk bestand
            sheet1 = pd.read_excel(tmp_file_path, sheet_name=0)
            try:
                sheet2 = pd.read_excel(tmp_file_path, sheet_name=1)
            except:
                sheet2 = None
            
            st.success("✅ Bestand gelezen via tijdelijk bestand")
            return sheet1, sheet2
            
        finally:
            # Verwijder tijdelijk bestand
            if os.path.exists(tmp_file_path):
                os.unlink(tmp_file_path)
                
    except Exception as e:
        st.error(f"❌ Alle lees-strategieën gefaald. Laatste fout: {str(e)}")
        st.error("💡 Probeer het bestand opnieuw op te slaan als .xlsx in Excel of LibreOffice")
        st.info("🔍 Controleer of het bestand:")
        st.info("   • Daadwerkelijk een Excel bestand is (.xlsx)")
        st.info("   • Niet wachtwoord beveiligd is")
        st.info("   • Geen corrupt formatting bevat")
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
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Taak met Norm', index=False)
        
        # Probeer formatting toe te passen (optioneel)
        try:
            from openpyxl.styles import Font, PatternFill, Alignment
            
            workbook = writer.book
            worksheet = workbook['Taak met Norm']
            
            # Header styling
            header_fill = PatternFill(start_color='D7E4BC', end_color='D7E4BC', fill_type='solid')
            header_font = Font(bold=True)
            
            # Style headers
            for col in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(wrap_text=True, vertical='top')
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
                
        except Exception as e:
            st.warning(f"Formatting warning: {str(e)}")
    
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
        
        st.header("🔧 Problemen?")
        st.info("Bij Excel lees-fouten:\n• Sla bestand opnieuw op als .xlsx\n• Verwijder formatting\n• App heeft ingebouwde Excel cleaner")
        
        if st.button("📖 Volledige Troubleshooting"):
            st.balloons()
            st.success("Bekijk TROUBLESHOOTING.md in de project directory voor uitgebreide hulp!")
    
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