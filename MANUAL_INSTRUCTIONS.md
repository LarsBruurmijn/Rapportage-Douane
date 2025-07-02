# 🔧 Handmatige Excel Taak-Norm Verrijker - Gebruikershandleiding

## ✨ Waarom de handmatige versie?

De handmatige versie geeft je **volledige controle** over het proces en lost alle Excel lees-problemen op:

- ✅ **Geen Excel lees-fouten** meer
- ✅ **Werkt met alle Excel bestanden** (ongeacht complexiteit)
- ✅ **Kies zelf welke tabbladen** je wilt gebruiken
- ✅ **Specificeer exact welke kolommen** wat bevatten
- ✅ **Stapsgewijze configuratie** met previews

## 🚀 Hoe te starten

### Optie 1: Via start script (Aanbevolen)
```bash
./start_apps.sh
```
Dan ga naar: http://localhost:8502

### Optie 2: Handmatig
```bash
streamlit run app_manual.py --server.port 8502
```

## 📋 Stap-voor-stap instructies

### Stap 1: Upload bestanden
1. Upload je **taakbestand** (Excel met taken)
2. Upload je **normbestand** (Excel met normen en mapping)

### Stap 2: Configureer tabbladen
3. **Selecteer taak sheet** - kies het tabblad met je taken
4. **Selecteer norm data sheet** - kies het tabblad met normgegevens
5. **Selecteer mapping sheet** - kies het tabblad met de mapping tussen norm-ID's en taaknamen
6. Klik **"Laad Preview van Geselecteerde Sheets"**

### Stap 3: Configureer kolommen
7. **Taak Data**: Selecteer welke kolom de taaknamen bevat
8. **Norm Data**: Selecteer welke kolom de norm-ID's bevat
9. **Mapping Data**: 
   - Selecteer welke kolom de norm-ID's bevat
   - Selecteer welke kolom de taaknamen bevat

### Stap 4: Verwerken
10. Klik **"Verwerk Data"**
11. Bekijk het resultaat en de statistieken
12. Download het verwerkte Excel bestand

## 📊 Voorbeeld configuratie

### Typische taakenbestand structuur:
```
| Taak_ID | Taaknaam        | Beschrijving | Prioriteit |
|---------|-----------------|--------------|------------|
| T001    | Projectplanning | Planning...  | Hoog       |
| T002    | Datanalyse      | Analyse...   | Gemiddeld  |
```
**→ Selecteer kolom "Taaknaam"**

### Typische normbestand - Normgegevens sheet:
```
| Norm_ID | Norm_Naam           | Categorie | Uren |
|---------|---------------------|-----------|------|
| N001    | Project Management  | Management| 40   |
| N002    | Data Analysis       | Analyse   | 16   |
```
**→ Selecteer kolom "Norm_ID"**

### Typische normbestand - Mapping sheet:
```
| Norm_ID | Taaknaam        |
|---------|-----------------|
| N001    | Projectplanning |
| N002    | Datanalyse      |
```
**→ Selecteer "Norm_ID" en "Taaknaam"**

## 🔧 Voordelen van handmatige configuratie

| Probleem | Automatische versie | Handmatige versie |
|----------|-------------------|-------------------|
| Excel lees-fouten | ❌ Kan voorkomen | ✅ Opgelost |
| Meerdere tabbladen | ❌ Verwarring | ✅ Jij kiest |
| Complexe formatting | ❌ Problemen | ✅ Geen probleem |
| Verkeerde kolommen | ❌ Giswerk | ✅ Jij specificeert |
| Onbekende structuur | ❌ Faalt | ✅ Preview eerst |

## 🆘 Troubleshooting

### "Kan bestand niet lezen"
- **Oorzaak**: Corrupt Excel bestand of onbekend formaat
- **Oplossing**: Sla het bestand opnieuw op als .xlsx in Excel

### "Geen mappings gevonden"
- **Oorzaak**: Verkeerde kolommen geselecteerd voor mapping
- **Oplossing**: Controleer of je de juiste kolommen hebt geselecteerd in stap 3

### "Kan sheet namen niet ophalen"
- **Oorzaak**: Bestand is niet leesbaar
- **Oplossing**: Controleer of het bestand een geldig Excel bestand is

### Lage koppelingspercentage
- **Oorzaak**: Taaknamen komen niet overeen met mapping
- **Oplossing**: Controleer de spelling en naming in je mapping sheet

## 💡 Tips voor succes

1. **Preview altijd eerst** - bekijk de data voordat je configureert
2. **Controleer kolom namen** - zorg dat je de juiste kolommen selecteert
3. **Mapping is cruciaal** - de taaknamen in je mapping moeten exact overeenkomen
4. **Case-insensitive** - hoofdletters maken niet uit voor matching
5. **Fuzzy matching** - gedeeltelijke overeenkomsten werken ook

## 🎯 Wanneer te gebruiken

**Gebruik de handmatige versie als:**
- Je Excel lees-fouten krijgt met de automatische versie
- Je bestanden meerdere tabbladen hebben
- Je volledige controle wilt over het proces
- Je complexe Excel bestanden hebt met veel formatting
- De automatische detectie verkeerde aannames maakt

**De handmatige versie is ontworpen om ALTIJD te werken**, ongeacht de complexiteit van je Excel bestanden!