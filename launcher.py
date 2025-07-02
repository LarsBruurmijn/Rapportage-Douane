import streamlit as st

def main():
    st.set_page_config(
        page_title="Excel Taak-Norm Verrijker - Keuze Menu",
        page_icon="🚀",
        layout="centered"
    )
    
    st.title("🚀 Excel Taak-Norm Verrijker")
    st.markdown("Kies de versie die het beste bij jouw situatie past:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🤖 Automatische Versie")
        st.markdown("""
        **Voor eenvoudige bestanden:**
        - Automatische sheet detectie
        - Intelligente kolom detectie
        - Snelle verwerking
        - Geschikt voor standaard Excel files
        
        **Nadelen:**
        - Kan problemen hebben met complexe formatting
        - Minder controle over proces
        """)
        
        if st.button("🤖 Start Automatische Versie", use_container_width=True):
            st.info("🔗 Ga naar: **http://localhost:8501**")
            st.code("streamlit run app.py --server.port 8501", language="bash")
    
    with col2:
        st.subheader("🔧 Handmatige Versie")
        st.markdown("""
        **Voor alle Excel bestanden:**
        - ✅ Volledige controle over tabbladen
        - ✅ Handmatige kolom specificatie
        - ✅ Werkt met complexe bestanden
        - ✅ Stapsgewijze configuratie
        
        **Aanbevolen als:**
        - Je Excel lees-fouten krijgt
        - Je meerdere tabbladen hebt
        - Je volledige controle wilt
        """)
        
        if st.button("🔧 Start Handmatige Versie", use_container_width=True):
            st.info("🔗 Ga naar: **http://localhost:8502**")
            st.code("streamlit run app_manual.py --server.port 8502", language="bash")
    
    st.markdown("---")
    
    st.subheader("📋 Wat is het verschil?")
    
    comparison_data = {
        "Eigenschap": [
            "Excel file ondersteuning",
            "Sheet detectie", 
            "Kolom mapping",
            "Error handling",
            "Complexe formatting",
            "Gebruikerscontrole",
            "Snelheid"
        ],
        "🤖 Automatisch": [
            "Standaard Excel files",
            "Automatisch",
            "Automatisch", 
            "Basis",
            "Problemen mogelijk",
            "Beperkt",
            "Snel"
        ],
        "🔧 Handmatig": [
            "Alle Excel files", 
            "Gebruiker selecteert",
            "Gebruiker specificeert",
            "Robuust",
            "Geen problemen",
            "Volledig",
            "Langzamer"
        ]
    }
    
    import pandas as pd
    df = pd.DataFrame(comparison_data)
    st.dataframe(df, use_container_width=True, hide_index=True)
    
    st.info("💡 **Tip**: Als je problemen hebt met de automatische versie, probeer dan de handmatige versie!")

if __name__ == "__main__":
    main()