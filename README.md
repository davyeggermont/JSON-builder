# JSON-builder Cursussen

Zet `Cursussen.xlsx` om naar de cursussen-JSON (zelfde structuur als `example/Cursussen.json`).
Je hoeft niets te installeren en er is geen internet nodig. Er wordt ook niets verstuurd.

## Gebruik
1. Vul `Cursussen.xlsx` in (zie hieronder).
2. Dubbelklik op `json-builder.html`. Het bestand opent in Edge of Chrome.
3. Sleep de Excel in het venster of klik om hem te kiezen.
4. Zijn er geen fouten, dan kies je **Kopiëren** of **Download Cursussen.json**.
   Zijn er wel fouten, dan zie je per fout het blad en de cel. Pas de Excel aan en laad hem opnieuw.

Om de tool op een andere pc te gebruiken, kopieer je gewoon `json-builder.html` en `Cursussen.xlsx`.

## De Excel
**Blad Cursussen**: één rij per cursus.
- Titel, Categorie, Categorie URL, Vormingsorganisatie
- Taal en Statuut (keuzelijst)
- Begin/Einde inschrijving (datum)
- Vormingsjaar (`2026-2027`)
- Credits (leeg = 0)
- Verlofstelsels, mandaten en regio's: zet een **X** in de kolom van elke code die geldt.

**Blad Activiteiten**: één rij per activiteit.
- Cursus: de titel, exact zoals in het blad Cursussen (keuzelijst).
- Gemeente
- Residentieel en Actief: JA of NEE.
- Datum 1 … Datum 8: optioneel. Lege datums worden overgeslagen. Wil je meer datums, voeg dan een kolom "Datum 9" toe, enz.

**Blad Lijsten**: de toegelaten waarden.
- Een nieuwe Taal of een nieuw Statuut voeg je hier gewoon toe.
- Bij een nieuw verlofstelsel, mandaat of nieuwe regio moet je ook een kolom met exact die code toevoegen in het blad Cursussen.

De volgorde van de kolommen maakt niet uit: de tool zoekt ze op naam (rij 1).

## Sleutels (worden automatisch gemaakt)
- **Cursus**: de titel in CamelCase, zonder accenten of leestekens, dan `_` en het vormingsjaar zonder streepje.
  `Rechten en Plichten op het werk 32u` + `2026-2027` → `RechtenEnPlichtenOpHetWerk32u_20262027`
- **Activiteit**: de cursussleutel, dan `_PLAATS_` en een volgnummer per plaats.
  `..._AALST_01`, `..._AALST_02`, `..._SINTNIKLAAS_01`
- Activiteiten staan in de JSON alfabetisch op gemeente, ongeacht de volgorde in de Excel. Binnen dezelfde gemeente bepaalt de eerste datum de volgorde. Datums binnen een activiteit worden chronologisch gezet.
- Dubbele sleutels geven een fout. Twee cursussen met dezelfde titel en hetzelfde vormingsjaar kan dus niet.

## Ontwikkeling
`tools/maak-sjabloon.ps1` maakt `Cursussen.xlsx` opnieuw aan vanuit een JSON (via Excel). Met `-Leeg` krijg je een leeg sjabloon.
Voor de tool zelf is dit script niet nodig.
