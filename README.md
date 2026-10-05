# Hapklaar

Weekmenu en boodschappenlijst voor twee. Plan de avonden van zondag tot zaterdag met de agenda ernaast; de boodschappenlijst vult zich vanzelf, per schap gesorteerd.

## Hoe het in elkaar zit

| Onderdeel | Waar | Wat |
|---|---|---|
| App | GitHub Pages (`index.html`, `sw.js`, manifest, iconen) | Werkt offline, installeerbaar op het beginscherm |
| Backend | Google Apps Script (`apps-script/Code.gs`), gebonden aan de sheet | Laden, opslaan in batches, agenda, sortering door Claude |
| Database | Google Sheet "Hapklaar database" | Tabs Recepten, Ingredienten, Planning, Boodschappen, Categorieen |

De web-app-URL en de sleutel staan niet in deze repo. Ze komen via de deel-link (`#k=…&api=…`) of het koppelscherm in de app en blijven op het apparaat.

## Eenmalig opzetten

1. Open de sheet **Hapklaar database** en kies Extensies > Apps Script.
2. Vervang alles in `Code.gs` door de inhoud van [`apps-script/Code.gs`](apps-script/Code.gs) en sla op.
3. Kies bovenin de functie `setup` en klik Uitvoeren. Geef toestemming (Geavanceerd > Ga naar project, want het is je eigen script). Het logboek toont je sleutel.
4. Optioneel: Projectinstellingen > Scripteigenschappen > `CLAUDE_API_KEY` met een sleutel van console.anthropic.com. Zonder sleutel werkt alles; onbekende producten komen dan bij Overig en zet je zelf goed.
5. Implementeren > Nieuwe implementatie > type Web-app. Uitvoeren als: ik. Toegang: iedereen. Kopieer de web-app-URL.
6. Open de app, vul de web-app-URL en de sleutel in en kies Koppelen. Via Instellingen deel je de app met een link of QR-code.

## Wijzigingen

- **App:** push naar `main`. Telefoons laden de nieuwe versie bij de tweede keer openen. Verhoog `VERSION` in `sw.js` en `index.html` bij elke release.
- **Backend:** plak de nieuwe code, dan Implementeren > Implementaties beheren > potlood > Versie: Nieuwe versie. De URL blijft gelijk.
- **Sleutel gelekt:** draai `nieuweSleutel` in Apps Script en deel de nieuwe link. Oude links werken dan niet meer.
