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
- **🆕 Ondersteunt meerdere tabbladen** - kies welke je wilt gebruiken
- Bevat een lijst met taken
- Gebruiker kan de kolom met taaknamen selecteren

### Normbestand
- Excel bestand (.xlsx of .xls) met **minimaal 2 tabbladen**:
  - **Normgegevens tabblad**: Alle norm informatie
  - **Mapping tabblad**: Mapping tussen Norm-ID en Taaknaam (kolom 1: Norm-ID, kolom 2: Taaknaam)
- **🆕 Ondersteunt 3+ tabbladen** - selecteer precies welke je nodig hebt

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

2. **🆕 Selecteer Tabbladen**:
   - App detecteert automatisch alle beschikbare sheets
   - Kies het taak-tabblad uit de dropdown
   - Kies het normgegevens-tabblad uit de dropdown
   - Kies het mapping-tabblad uit de dropdown
   - Klik "Lees Geselecteerde Sheets"

3. **Selecteer Taaknaam Kolom**:
   - Kies uit de dropdown welke kolom de taaknamen bevat

4. **Preview Data**:
   - Bekijk een preview van alle geselecteerde tabbladen
   - Controleer de mapping data

5. **Genereer Resultaat**:
   - Klik op "Genereer Uitvoerbestand"
   - Bekijk statistieken over de koppeling

6. **Download**:
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

- **🆕 Multi-sheet ondersteuning**: Werk met Excel bestanden met meerdere tabbladen
- **🆕 Sheet selectie**: Kies precies welke tabbladen je wilt gebruiken
- **🆕 Robuuste Excel handling**: Automatische error recovery en file cleaning
- Real-time preview van geüploade data
- Interactieve statistieken over koppelingen
- Professionele Excel formatting
- Uitgebreide error handling met multiple engines
- Gebruiksvriendelijke interface met instructies
- Session state management voor betere UX