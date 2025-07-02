# 🔧 Troubleshooting Guide - Excel Problemen Oplossen

## ❌ Veel voorkomende fouten

### "expected <class 'openpyxl.styles.fills.Fill'>" fout

Deze fout treedt op wanneer Excel bestanden complexe formatting bevatten die conflicteert met de leesfunctie.

## ✅ Oplossingen (in volgorde van effectiviteit)

### 1. **Bestand opnieuw opslaan**
- Open het Excel bestand in Microsoft Excel of LibreOffice
- Ga naar `Bestand` → `Opslaan als`
- Kies `Excel Werkmap (.xlsx)` formaat
- Sla op met een nieuwe naam
- Upload het nieuwe bestand

### 2. **Formatting verwijderen**
- Selecteer alle cellen (Ctrl+A)
- Ga naar `Start` → `Opmaak wissen` → `Alles wissen`
- Behoud alleen de data, verwijder alle styling
- Sla opnieuw op als .xlsx

### 3. **Kopiëren naar nieuw bestand**
- Selecteer alle data (Ctrl+A)
- Kopieer (Ctrl+C)
- Maak een nieuw Excel bestand
- Plak alleen waarden (`Plakken speciaal` → `Waarden`)
- Sla op als .xlsx

### 4. **Gebruik de ingebouwde Excel Cleaner**
De applicatie heeft nu een automatische Excel cleaner die:
- Meerdere read-engines probeert
- Problematische formatting verwijdert
- Data extraheert zonder styling
- Een schoon bestand regenereert

## 📋 Bestandsvereisten

### Taakbestand:
- ✅ .xlsx formaat
- ✅ Minimaal 1 kolom met taaknamen
- ✅ Headers in de eerste rij

### Normbestand:
- ✅ .xlsx formaat
- ✅ **Tabblad 1**: Normgegevens (alle norm informatie)
- ✅ **Tabblad 2**: Mapping tussen norm-ID en taaknaam
  - Kolom 1: Norm_ID
  - Kolom 2: Taaknaam (moet overeenkomen met taakbestand)

## 🔍 Debug Informatie

De applicatie toont nu debug informatie:
- Welke read-engine succesvol was
- Aantal rijen en kolommen per sheet
- Welke strategie gebruikt werd (standaard/cleaning/tijdelijk bestand)

## 🆘 Als niets werkt

1. **Controleer bestandsintegriteit**:
   - Is het echt een Excel bestand?
   - Is het niet wachtwoord beveiligd?
   - Kan je het openen in Excel?

2. **Maak handmatig een nieuw bestand**:
   - Kopieer data naar kladblok
   - Plak in nieuw Excel bestand
   - Voeg headers toe
   - Sla op als .xlsx

3. **Neem contact op** met de ontwikkelaar met:
   - Exact foutbericht
   - Excel versie gebruikt
   - Screenshot van de data structuur