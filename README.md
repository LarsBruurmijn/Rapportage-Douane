# 📊 Excel Taak-Norm Verrijker

Een interactieve Streamlit webapplicatie voor het verwerken en verrijken van Excel-bestanden door taken te koppelen aan normgegevens en tekstvelden op te schonen.

## 🚀 Functionaliteiten

### Hoofdfuncties
- **Bestand Upload**: Upload twee Excel-bestanden via de browser
- **Automatische Koppeling**: Koppel taken aan normen op basis van taaknamen
- **Tekst Opschoning**: Intelligent opschonen van tekstvelden
- **Excel Export**: Download het verwerkte resultaat als Excel-bestand

### Tekst Opschoning Features
- Voegt spaties toe waar woorden aan elkaar geplakt zijn
- Herstelt regeleinden na punten en opsommingstekens
- Verbetert de leesbaarheid van geëxporteerde teksten
- Schoont interpunctie op (spaties na komma's, dubbele punten, etc.)

## 📋 Bestandsformaten

### Taakbestand
- Excel bestand (.xlsx of .xls)
- Bevat een lijst met taken
- Gebruiker kan de kolom met taaknamen selecteren

### Normbestand
- Excel bestand (.xlsx of .xls) met **2 tabbladen**:
  - **Tabblad 1**: Normgegevens (alle norm informatie)
  - **Tabblad 2**: Mapping (kolom 1: Norm-ID, kolom 2: Taaknaam)

## 🛠️ Installatie

1. Installeer de vereiste dependencies:
```bash
pip install -r requirements.txt
```

2. Start de applicatie:
```bash
streamlit run app.py
```

3. Open je browser en ga naar de aangegeven URL (meestal `http://localhost:8501`)

## 📖 Gebruiksinstructies

1. **Upload Bestanden**: 
   - Upload het taakbestand in het linker paneel
   - Upload het normbestand in het rechter paneel

2. **Selecteer Taaknaam Kolom**:
   - Kies uit de dropdown welke kolom de taaknamen bevat

3. **Preview Data**:
   - Bekijk een preview van beide bestanden
   - Controleer de mapping data

4. **Genereer Resultaat**:
   - Klik op "Genereer Uitvoerbestand"
   - Bekijk statistieken over de koppeling

5. **Download**:
   - Download het verwerkte Excel-bestand
   - Het bestand bevat een tabblad "Taak met Norm"

## 🔧 Dependencies

- `streamlit` - Web interface
- `pandas` - Data verwerking
- `openpyxl` - Excel bestanden lezen
- `xlsxwriter` - Excel bestanden schrijven

## 📊 Output

Het resultaat is een Excel-bestand met:
- Alle oorspronkelijke taakgegevens
- Toegevoegde norm-ID kolom
- Alle bijbehorende normgegevens
- Opgeschoonde tekstvelden voor betere leesbaarheid
- Professionele opmaak met headers

## 🎯 Use Cases

- HR-afdelingen die taken koppelen aan functienormen
- Projectmanagement voor normering van activiteiten
- Kwaliteitsmanagement voor standaardisatie
- Elke situatie waar taken gekoppeld moeten worden aan referentiestandaarden

## 🔍 Matching Algoritme

De applicatie gebruikt zowel exacte als fuzzy matching:
1. **Exacte match**: Directe overeenkomst tussen taaknamen
2. **Gedeeltelijke match**: Zoekt naar overeenkomsten in taaknaam substrings
3. **Case-insensitive**: Hoofdletter ongevoelig matching

## ✨ Extra Features

- Real-time preview van geüploade data
- Interactieve statistieken over koppelingen
- Professionele Excel formatting
- Uitgebreide error handling
- Gebruiksvriendelijke interface met instructies