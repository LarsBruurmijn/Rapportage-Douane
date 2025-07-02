import streamlit as st
import pandas as pd

def main():
    st.set_page_config(
        page_title="Excel Taak-Norm Verrijker - App Selector",
        page_icon="🚀",
        layout="wide"
    )
    
    st.title("🚀 Excel Taak-Norm Verrijker - App Selector")
    st.markdown("Kies de versie die het beste bij jouw Excel problemen past:")
    
    # App definitions
    apps = [
        {
            "name": "🩹 Fixed App (AANBEVOLEN)",
            "description": "Specifiek gefixed voor openpyxl 3.1.5 'expected Fill' bugs",
            "port": 8505,
            "pros": [
                "✅ Lost 'expected Fill' errors op",
                "✅ 4 verschillende fix strategieën", 
                "✅ Werkt met problematische formatting",
                "✅ read_only + data_only modes",
                "✅ Excel file reconstruction"
            ],
            "cons": [
                "⚠️ Specifiek voor openpyxl 3.1.5 bugs"
            ],
            "best_for": "Als je 'expected Fill' errors krijgt",
            "recommended": True
        },
        {
            "name": "🛡️ Bulletproof App",
            "description": "Ultra-robuuste versie met 6 verschillende read strategieën",
            "port": 8504,
            "pros": [
                "✅ 6 verschillende read methods",
                "✅ Automatische fallback",
                "✅ Memory optimization", 
                "✅ Format cleaning",
                "✅ Guaranteed success"
            ],
            "cons": [
                "⚠️ Kan langzamer zijn",
                "⚠️ Meer complex"
            ],
            "best_for": "Complexe Excel bestanden met onbekende problemen"
        },
        {
            "name": "✋ Manual App",
            "description": "Handmatige configuratie - volledige controle",
            "port": 8502,
            "pros": [
                "✅ Volledige gebruikerscontrole",
                "✅ Stapsgewijze configuratie",
                "✅ Werkt met alle Excel files",
                "✅ Preview van data",
                "✅ Handmatige kolom mapping"
            ],
            "cons": [
                "⚠️ Meer stappen vereist",
                "⚠️ Langzamer proces"
            ],
            "best_for": "Complexe multi-sheet bestanden"
        },
        {
            "name": "🔍 Debug Tool",
            "description": "Diagnostiek tool voor Excel problemen",
            "port": 8503,
            "pros": [
                "✅ Uitgebreide file analyse",
                "✅ Test alle read methods",
                "✅ Identificeert specifieke errors",
                "✅ File corruption check",
                "✅ Memory usage analyse"
            ],
            "cons": [
                "⚠️ Alleen voor diagnostiek"
            ],
            "best_for": "Onderzoeken waarom Excel files niet werken"
        },
        {
            "name": "🔧 Original App",
            "description": "Basis versie met multi-sheet ondersteuning",
            "port": 8501,
            "pros": [
                "✅ Snelle verwerking",
                "✅ Multi-sheet ondersteuning",
                "✅ Automatische detectie",
                "✅ Eenvoudig te gebruiken"
            ],
            "cons": [
                "⚠️ Kan problemen hebben met complex formatting",
                "⚠️ Beperkte error handling"
            ],
            "best_for": "Eenvoudige, goed geformatteerde Excel bestanden"
        }
    ]
    
    # Show apps in grid
    cols = st.columns(2)
    
    for i, app in enumerate(apps):
        with cols[i % 2]:
            # Create card for each app
            with st.container():
                if app.get('recommended'):
                    st.markdown(f"### 🌟 {app['name']}")
                    st.markdown("**🎯 MEEST AANBEVOLEN VOOR JOUW PROBLEEM**")
                else:
                    st.markdown(f"### {app['name']}")
                
                st.markdown(f"*{app['description']}*")
                
                # Pros
                st.markdown("**Voordelen:**")
                for pro in app['pros']:
                    st.markdown(f"- {pro}")
                
                # Cons  
                if app['cons']:
                    st.markdown("**Nadelen:**")
                    for con in app['cons']:
                        st.markdown(f"- {con}")
                
                # Best for
                st.markdown(f"**🎯 Best voor:** {app['best_for']}")
                
                # Launch button
                if app.get('recommended'):
                    button_type = "primary"
                else:
                    button_type = "secondary"
                
                if st.button(f"🚀 Start {app['name']}", key=f"btn_{app['port']}", type=button_type):
                    st.success(f"🔗 Ga naar: **http://localhost:{app['port']}**")
                    st.code(f"App draait op poort {app['port']}", language="text")
                
                st.markdown("---")
    
    # Quick decision guide
    st.subheader("🤔 Welke app moet ik kiezen?")
    
    with st.expander("📋 Snelle beslissingshulp", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Als je deze errors krijgt:**")
            st.code("expected <class 'openpyxl.styles.fills.Fill'>")
            st.markdown("👉 **Gebruik: 🩹 Fixed App (poort 8505)**")
            
            st.markdown("**Als je Excel files niet kunnen worden gelezen:**")
            st.markdown("👉 **Gebruik: 🔍 Debug Tool (poort 8503)** om te onderzoeken")
            st.markdown("👉 **Dan: 🛡️ Bulletproof App (poort 8504)**")
        
        with col2:
            st.markdown("**Als je veel controle wilt:**")
            st.markdown("👉 **Gebruik: ✋ Manual App (poort 8502)**")
            
            st.markdown("**Als je eenvoudige Excel bestanden hebt:**")
            st.markdown("👉 **Gebruik: 🔧 Original App (poort 8501)**")
    
    # Status check
    st.subheader("📊 App Status")
    status_data = []
    
    for app in apps:
        try:
            import requests
            response = requests.get(f"http://localhost:{app['port']}", timeout=1)
            status = "🟢 Online" if response.status_code == 200 else "🔴 Offline"
        except:
            status = "🔴 Offline"
        
        status_data.append({
            "App": app['name'],
            "Port": app['port'],
            "Status": status,
            "URL": f"http://localhost:{app['port']}"
        })
    
    df_status = pd.DataFrame(status_data)
    st.dataframe(df_status, use_container_width=True, hide_index=True)
    
    # Instructions
    with st.sidebar:
        st.header("📖 Instructies")
        st.markdown("""
        **Voor 'expected Fill' errors:**
        1. Gebruik 🩹 **Fixed App** eerst
        2. Als dat niet werkt: 🛡️ **Bulletproof App**
        
        **Voor andere Excel problemen:**
        1. Test met 🔍 **Debug Tool**
        2. Gebruik 🛡️ **Bulletproof App**
        3. Als backup: ✋ **Manual App**
        
        **Voor eenvoudige bestanden:**
        - Gebruik 🔧 **Original App**
        """)
        
        st.markdown("---")
        st.markdown("**🎯 Bestandsgrootte:** 500KB is geen probleem")
        st.markdown("**🔧 Fix:** Openpyxl 3.1.5 bugs aangepakt")

if __name__ == "__main__":
    main()