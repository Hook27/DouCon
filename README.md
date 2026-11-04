# Doublurecontrole map

Een tool die een map (inclusief submappen) controleert op dubbele bestanden. De tool is één HTML-bestand, [`index.html`](index.html), en werkt volledig lokaal in de browser. Er wordt niets geüpload of verzonden.

**Online openen:** <https://hook27.github.io/DouCon/> (als GitHub Pages aanstaat)
**Of lokaal:** download `index.html` en open het in Edge of Chrome.

## Wat de tool vindt

| Categorie | Betekenis |
|---|---|
| **Zelfde naam, afwijkende inhoud** | Bestanden met dezelfde naam, maar ze zijn niet gelijk. Beoordeel ze handmatig. |
| **Identieke doublures** | Zelfde naam en exact dezelfde inhoud. Eén exemplaar is genoeg. |
| **Zelfde inhoud, andere naam** | Exact dezelfde inhoud onder een andere bestandsnaam. |
| **Niet leesbaar** | Bestanden die de browser niet kon lezen, bijvoorbeeld omdat het pad langer is dan 260 tekens of het bestand vergrendeld is. |

Of twee bestanden gelijk zijn, bepaalt de tool met een SHA-256-hash: een digitale vingerafdruk van de inhoud. Zelfde hash betekent byte voor byte gelijk.

## Gebruik

1. **Instellingen kiezen.** Bijvoorbeeld hoofdlettergevoeligheid, kopie-aanduidingen als "Rapport (1)" negeren, systeembestanden overslaan, of een uitsluitlijst invullen.
2. **Map kiezen en analyseren.** De browser vraagt of je de bestanden wilt "uploaden". Dat is standaardtekst: de bestanden worden alleen lokaal gelezen.
3. **Resultaten beoordelen.**
   - Bij afwijkingen kun je Word-, Excel- en PowerPoint-bestanden inhoudelijk vergelijken (.docx, .xlsx, .pptx).
   - Met **Voorstel toepassen** kiest de tool per inhoud welk exemplaar behouden blijft. Per bestand kun je de keuze aanpassen.
4. **Opmerkingen toevoegen** per groep. Sla ze op met **Opmerkingen opslaan**; dat geeft een .json-bestand. Daarin staan ook de keuzes behouden/verwijderen, de analyse-instellingen, de uit- en ingeklapte groepen en het actieve tabblad.
   - Later verder werken: analyseer de map en laad het bestand met **Opmerkingen laden**.
   - Wijken de opgeslagen instellingen af van de huidige, dan zie je welke dat zijn. Met één knop neem je ze over en wordt de map opnieuw geanalyseerd.
5. **Exporteren** naar een Word-rapport voor de mapbeheerders, of naar Excel.

### Hercontrole na het opschonen

Laat de mapbeheerders de voorgestelde acties uitvoeren. Kies daarna dezelfde map opnieuw en laad het opmerkingenbestand met het vinkje **Hercontrole uitvoeren** aan. Het tabblad **Hercontrole** toont per bestand:

- **Verwijderen:** uitgevoerd of niet uitgevoerd.
- **Behouden:** aanwezig of ontbreekt. Staat de inhoud nergens meer in de map, dan krijg je een waarschuwing. Staat de inhoud nog elders, dan meldt de tool dat het bestand mogelijk is verplaatst.

De hercontrole komt ook in het Word-rapport en de Excel-export. Wil je alleen verder werken met je opmerkingen, zet het vinkje dan uit. Je kunt het later alsnog aanzetten.

> **Let op:** kies na het opschonen de map opnieuw via **Map kiezen en analyseren**. De knop **Opnieuw analyseren** gebruikt de bestandslijst van het moment waarop de map werd gekozen. Verwijderde bestanden staan daarin nog en geven dan een leesfout. De tool waarschuwt als dat gebeurt.

### Uitsluitlijst

Gescheiden door komma of puntkomma; hoofdletters maken niet uit.

| Invoer | Sluit uit |
|---|---|
| `Archief` | Elke map met die naam, op elk niveau |
| `Projecten/Oud` | Alleen die map, gerekend vanaf de gekozen map |
| `.tmp` | Alle bestanden met die extensie |
| `Oud*` | Jokerteken: elke map die met "Oud" begint |

## Privacy

- **Alles gebeurt lokaal.** De bestanden worden in de browser gelezen en verlaten je computer niet. De rapporten worden ook in de browser gemaakt.
- **Netwerkverkeer is geblokkeerd.** Een beveiligingsregel in het bestand (Content Security Policy) verbiedt de pagina om verbinding te maken met internet.
- **Bij openen via GitHub Pages** haalt de browser alleen de pagina zelf bij GitHub op. Je bestanden gaan niet naar GitHub.
- **Het enige wat de browser onthoudt** is de uitsluitlijst.

Zelf controleren: druk op F12, open het tabblad **Netwerk** en analyseer een map. Er verschijnen geen verzoeken naar buiten.

## Browsers

Aanbevolen: een actuele versie van **Microsoft Edge** of **Google Chrome**.
- In andere browsers worden bestanden opgeslagen via een gewone download in plaats van een opslaan-venster.
- Open de tool niet in de voorbeeldweergave van OneDrive of SharePoint, want daar blokkeert de browser het opslaan. Download het bestand en open het lokaal.

## Beperkingen

- **Oudere formaten.** Inhoudelijk vergelijken werkt niet voor .doc, .xls, .ppt en pdf; open die bestanden via de paden. Ook bestanden met een wachtwoord kunnen niet worden vergeleken.
- **Grote bestanden.** Bestanden vanaf 128 MB worden in delen gehasht. Het geheugengebruik blijft daardoor laag, maar ze worden wel langzamer verwerkt.
- **OneDrive.** Bestanden die alleen online staan, downloadt OneDrive eerst naar de pc zodra de tool ze leest.
- **Opnieuw analyseren.** Gebruikt de bestandslijst van het moment van kiezen; zie de waarschuwing bij de hercontrole.
- **Opmerkingenbestanden van v8.** Die bevatten geen bestandsgrootte. Bij de hercontrole kan de tool daardoor niet altijd nagaan of de inhoud van een ontbrekend bestand elders nog staat.

## Testen

In de map [`tests/`](tests) staat een testsuite die de tool in een echte browser (Chromium) doorloopt. Hij controleert onder andere de analyse, de Word-vergelijking, de exports, het opmerkingenbestand, de hercontrole en de snelheid bij 3000 groepen. De tests maken hun testmappen aan in een tijdelijke map en ruimen die daarna op.

Eenmalig installeren (Node.js nodig):

```bash
cd tests
npm install
npx playwright install chromium
```

Tests draaien:

```bash
npm test
```

Met `npm test -- --groot` loopt ook een test met twee bestanden van 130 MB mee. Die controleert het hashen in blokken en duurt wat langer. Draai de tests na elke wijziging aan `index.html`.

## Versies

| Versie | Wijzigingen |
|---|---|
| **v9.3** | Verbeteringen na een code-audit: waarschuwing als "Opnieuw analyseren" een verouderde bestandslijst gebruikt (ook in de hercontrole), groepen worden pas opgebouwd bij uitklappen en het filter wacht tot je stopt met typen (veel sneller bij grote mappen), kopie-aanduidingen van de Mac en `(1).txt` correct, LibreOffice-vergrendelbestanden overgeslagen, tegels met het toetsenbord te bedienen, getallen in Excel als getal, en een testsuite in `tests/`. |
| **v9.2** | Compactere indeling: instellingen klappen in na de analyse, samenvatting op één rij en een vast blok met samenvatting, tabbladen en werkbalk dat bij het scrollen in beeld blijft. Klik op een tegel in de samenvatting om naar het bijbehorende tabblad te gaan. |
| **v9.1** | Opmerkingenbestand bewaart ook uit- en ingeklapte groepen en het actieve tabblad. Afwijkende instellingen zijn met één knop over te nemen. Vinkje om de hercontrole wel of niet uit te voeren. |
| **v9** | Hercontrole na het opschonen, uitsluitlijst, grote bestanden in delen hashen, versienummer in rapport/Excel/opmerkingenbestand, één berekening voor overbodige opslag, jaartallen als "(2023)" niet meer als kopie-aanduiding, bediening met het toetsenbord. |
| **v8** | Eerste versie in deze repository. |
