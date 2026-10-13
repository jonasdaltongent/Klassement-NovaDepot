# Handleiding voor de leraar — Les 07: Het digitale klassement

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, fictief logistiek bedrijf en groothandel — module 2 van het jaarplan
(`Leerplannen/Lessenplan_TOINFO_ORLO_arbeidsmarkt_2026-2027.md`)
**Toestel:** **lokaal 18, Windows 11** — Google Workspace in Chrome en de Verkenner
**Lesdag:** uur 1 van week 7: 4ORLOb di 13 oktober (8ste uur) · 3ORLO di 13 oktober (9de uur) · 4ORLOa wo 14 oktober (4de uur)
· **tweede uur** (sinds 13-10-2026: les 08 vervalt): 3ORLO wo 14 oktober (2de uur) · 4ORLOb wo 14 oktober (3de uur) · 4ORLOa do 15 oktober (7de uur)
**Kernleerplandoelen:** `BK2_02.09` (`.01`–`.03`, klassementen) — toepassen
**Deadline:** vrijdag 16 oktober 2026, 20.00 uur · 20 punten

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 09-10-2026:
     klassen 3ORLO, 4ORLOa, 4ORLOb · les 07 van W07 · toestel lokaal 18, Windows 11 ·
     lesstart = ophalen uit les 06 · wie les 06 miste, moet snel mee kunnen (de mappen staan in hun eigen Drive) ·
     werkdocument: ja · daarna les 08 en dan de lessen van 3MWb en 3MWWE. -->

De leerling zet zes nieuwe documenten van twee nieuwe leveranciers in de mappen van les 06 en maakt er een
**klassement** van: de leveringsbonnen **op datum** (jaar-maand-dag vooraan in de naam), de facturen **op
nummer** (het factuurnummer vooraan), en de prijslijsten staan vanzelf **op naam**. De naam van les 06
(*Wat - Van wie*) blijft gewoon staan; er komt alleen een sleutel vooraan bij.

> [!NOTE]
> **Wie les 06 miste.** Hun map staat in hun eigen Drive: daar kan jij niet aan, en ook geen script. Daarom
> hangt er een **kant-en-klare map** aan de opdracht: `NovaDepot.zip`. Uitpakken in de Verkenner en in Drive
> zetten met **Nieuw › Map uploaden** duurt ± 2 minuten. Het staat in stap 1, handeling 4 (dicht tot de leerling ze
> opent), met een schermafbeelding. Wie een map heeft die niet klopt, geeft die eerst de naam *NovaDepot oud*.

---

## 1. Inhoud van het pakket

```text
W07 - Les 07 - ORLO - Het digitale klassement/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 11 dia's; dia 2 alleen in het tweede uur, dia 8 controle, dia 9 samen stap 1, dia 10 kloktijden
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/
│   ├── novadepot-logo.svg/.png, novadepot-icon.svg, dalton-gent-logo.png
│   ├── fonts/                   # Atkinson Hyperlegible + Montserrat (OFL, zelf gehost)
│   ├── mascotte/                # Dalton de bever, de poses uit het sjabloon (.webp)
│   ├── video/                   # stap3-verplaatsen, stap4-datum-vooraan (09-10-2026) + zip-downloaden (uit les 06)
│   └── screenshots/             # zes schermafbeeldingen (09-10-2026) + LEESMIJ.md
├── werkdocument/
│   ├── WP2_Het-digitale-klassement.docx   # het werkdocument dat de leerling INLEVERT
│   ├── Nieuwe-documenten.zip              # de zes nieuwe documenten: aan de opdracht hangen
│   ├── NovaDepot.zip                      # de map van les 06, kant-en-klaar: ook aan de opdracht hangen
│   ├── nieuw/                             # de zes losse bestanden uit Nieuwe-documenten.zip
│   ├── inhaal/NovaDepot/                  # de inhoud van NovaDepot.zip
│   ├── maak_werkdocumenten.py             # maakt alles opnieuw (python-docx, openpyxl)
│   └── huisstijl.py, go-dalton-gent.png   # de huisstijl van het werkdocument (uit _sjablonen/lespakket/werkdocument/)
├── lesvoorbereiding.md
├── dalton-lesfiche.html         # openen, Kopieer de fiche, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

### 1b. Hoe de lespagina werkt

Dezelfde opbouw als les 06: links de **route** (zes stappen in twee groepen: *Je map aanvullen* · *Je
klassement*), midden **één stap** met vijf vaste blokken, rechts de **checklist** met 18 taken. Op een smal
venster staat alles onder elkaar en zie je alleen de taken van de huidige stap; in stap 6 staat de hele lijst open.

- 2 tot 5 handelingen per stap, één tip per stap (de fout die het vaakst gebeurt).
- **Stap 4** heeft een filmpje van 13 seconden dat vanzelf herhaalt, met nummers die bij de handelingen horen.
  **Stap 3** toont sinds 13-10-2026 tekeningen van het slepen (het filmpje van *Ordenen › Verplaatsen* staat niet meer op de pagina). Stap 4, 5 en 6 tonen ook een schermafbeelding van het resultaat: *zo ziet je map er daarna uit*.
- De **naamkaart** in stap 4 en 5 toont in kleur wat er vooraan komt (de datum, het nummer) en wat blijft staan.
- De **zelftest** in stap 6 heeft drie vragen, waarvan één de valkuil is: van Z tot A staat `28-09-2026` vóór
  `02-10-2026`, terwijl 2 oktober nieuwer is.
- **Volgorde: de nieuwste bovenaan** (Jonas, 12-10-2026), zoals de klasseerregel in Lift 3A Thema 4, p. 113. In Drive
  is dat **Sorteren › Naam › Z tot A** (stap 4, handeling 5). Drive onthoudt die keuze voor alle mappen (nagekeken in
  je demomap, 12-10-2026): ook de facturen staan dan met het hoogste nummer bovenaan, en de prijslijsten van Z tot A.
  Dat zegt de pagina eerlijk: het blijft alfabetisch, alleen omgekeerd.
- Windows 11 (lokaal 18). `localStorage` bewaart alleen de vinkjes en de laatste stap (voorvoegsel
  `novadepot_klassement_v1_`), met een wisknop.

- **Stap voor stap / Kort** (sinds 11-10-2026, zoals MW les 03; Jonas: "Pas je de lessen van ORLO ook aan volgens
  de nieuwe afspraken?"): elke handeling is één vette regel; open je ze, dan zie je de kleine stapjes, elk met één
  beeld (`assets/stapjes/`). Bovenaan *Wat moet je doen?* zet *Stap voor stap* alles open en *Kort* alles dicht. De
  les begint in *Stap voor stap* (`data-modus-start="stap"`). *Hulp nodig?* is weg: wat erin stond, staat als grijze
  regel onder het stapje of de handeling waar het misloopt; het tipvak bevat de meest waarschijnlijke fout. De
  hulpvolgorde is korter: *Toon stapjes* → je buur → jij.
- **De beelden bij de stapjes** (11-10-2026): uitsneden uit de filmpjes van stap 3 en 4 (Verplaatsen, Naam
  wijzigen, Home, OK; de rode kaders zaten al in de filmpjes) en `stap1-map-uploaden.png` en
  `stap2-bestand-uploaden.png`. Voor Classroom, de tabbladen, de lijst links in Drive en de **Verkenner** zijn het
  tekeningen in HTML en CSS (`.mock-opdracht`, `.teken-tabs`, `.teken-ui`): van de Verkenner zijn er geen
  schermafbeeldingen. Handeling 4 van stap 1 (*les 06 gemist*) heeft `data-dicht`: ze blijft dicht, ook in
  *Stap voor stap*.
- **Bijgestuurd na het eerste uur** (13-10-2026, Jonas na 4ORLOb en 3ORLO: tijd te krap, niemand snel klaar; les 08 vervalt):
  stap 3 met **slepen** (`_afspraken/klikpaden.md`); in stap 1 en 2 *Weergeven in map* in de downloads van Chrome, zodat ze
  hun zip-bestand ook in een volle map Downloads vinden; de startpagina groter (vier beelden in 2 × 2, grotere logo's);
  dia 2 *Waar ben je gebleven?* voor het tweede uur; dia 11 kondigt *Een veilig account* (les 09) aan.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 09-10-2026*

Repository [`jonasdaltongent/Klassement-NovaDepot`](https://github.com/jonasdaltongent/Klassement-NovaDepot),
Pages op branch `main`, map `/ (root)`. De lespagina staat op
[jonasdaltongent.github.io/Klassement-NovaDepot](https://jonasdaltongent.github.io/Klassement-NovaDepot/), de dia's op
[…/presentatie.html](https://jonasdaltongent.github.io/Klassement-NovaDepot/presentatie.html); dat adres staat
ook op dia 10, in `classroom.json` en in `lesdoelen.json`. Live bestanden nagekeken: gelijk aan de lokale.

### Stap 2 — De opdracht in Classroom, met de koppeling ✅ *concept op 09-10-2026*

Als **concept** klaargezet in [3ORLO](https://classroom.google.com/c/MTYyNjY1NjU1OTEx), [4ORLOa](https://classroom.google.com/c/MjUzMTgzNDM2NjZa) en [4ORLOb](https://classroom.google.com/c/MjUzMTc2OTkwNzNa) (teruggelezen: onderwerp, deadline, 20 punten, werkdocument
`STUDENT_COPY`, beide zips `VIEW`). Toewijzen doe je zelf met **Toewijzen**.

`classroom.json` zet de opdracht klaar in 3ORLO, 4ORLOa en 4ORLOb (zie `_tools/CLASSROOM-KOPPELING.md`):

| | Opdracht: **Werkplek 2 — Het digitale klassement** |
|---|---|
| **Onderwerp** | Digitale competenties (bestaat al) |
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `WP2_Het-digitale-klassement` (Google-document) — **Een kopie maken voor elke leerling** |
| **Bijlage 3** | `Nieuwe-documenten.zip` — **Leerlingen kunnen bestand bekijken** |
| **Bijlage 4** | `NovaDepot.zip` — **Leerlingen kunnen bestand bekijken** (alleen voor wie les 06 miste) |
| **Punten** | 20 |
| **Deadline** | vrijdag 16 oktober 2026, 20.00 uur |

Proef zonder Google (09-10-2026: in orde):

```bash
python3 "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/_tools/zet_opdracht_klaar.py" "/Volumes/Littlecisboy/Google drive/Toegepaste Informatica/2026-2027/W07 - Les 07 - ORLO - Het digitale klassement" --proef
```

Daarna `--maak` (concept). De AI publiceert niet in Classroom; dat doe jij ([classroom-koppeling](../../_afspraken/classroom-koppeling.md)).

Instructietekst (vult het script in):

```text
1. Open de lespagina (link) en je werkdocument WP2_Het-digitale-klassement.
2. Nieuwe-documenten.zip heb je nodig in stap 2. NovaDepot.zip is alleen voor wie les 06 miste (stap 1).
3. Klik op Inleveren, ten laatste vrijdag 16 oktober 2026 om 20.00 uur.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload `werkdocument/WP2_Het-digitale-klassement.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**,
   met *Uploads converteren* aan) en voeg het toe met **Bijvoegen** › **Drive** › **Een kopie maken voor elke leerling**.
2. Voeg de twee zip-bestanden toe met **Bijvoegen** › **Uploaden** › **Leerlingen kunnen bestand bekijken**.
3. Voeg de lespagina toe met **Link**.

### De demo klaarzetten (2 minuten)

Je doet op dia 7 de bon van Bakkerij Korstjes voor, in een eigen map NovaDepot. Pak `werkdocument/NovaDepot.zip`
uit en zet de map in je Drive met **Nieuw** › **Map uploaden**. Zo heb je dezelfde map als een leerling na les 06,
en heb je meteen de inhaalweg zelf eens gedaan.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 10 klopt.
- [ ] De opdracht staat in de drie klassen, met vier bijlagen en de juiste instelling per bijlage.
- [ ] Een testleerling krijgt een eigen kopie van het werkdocument, met de eigen naam in de titel.
- [ ] Met een **leerlingaccount** getest: de zip downloaden uit de opdracht (Openen met › Openen in nieuw tabblad › pijl linksboven).
- [ ] Je map NovaDepot voor de demo staat klaar, met de bon van Bakkerij Korstjes in *Leveringen*.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

Opgenomen en nagekeken op **jouw scherm** (Chrome, profiel Jonas, 09-10-2026) in de map *Demo voor
schermafbeeldingen (Claude)*, en in de Nederlandse help:

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Zip downloaden uit de opdracht | klik op het bestand › **Openen met** › **Openen in nieuw tabblad** › linksboven de pijl (**Downloaden**) | Jonas' schermopname (27-09-2026), `_afspraken/klikpaden.md` |
| Verkenner | **Windows** + **E**, links **Downloads** | [Microsoft](https://support.microsoft.com/nl-nl/windows/sneltoetsen-in-windows-dcc61a57-8ff0-cffe-9796-cb9706c75eec) |
| Uitpakken (Windows 11) | rechtermuisknop op het zip-bestand › **Alles uitpakken…** en de instructies volgen; de pagina zegt *bevestig in het venster* (de knopnaam staat niet in de help) | [Microsoft](https://support.microsoft.com/nl-nl/windows/bestanden-comprimeren-en-uitpakken-8d28fa72-f2f9-712f-67df-f80cf89fd4e5), nagelezen 09-10-2026 |
| Uploaden | **Nieuw** › **Bestand uploaden** · **Map uploaden** | jouw scherm (04-10 en 09-10-2026) |
| Verplaatsen | **slepen**: sleep het bestand naar de map en laat los (sinds 13-10-2026). In de tip, om een bestand uit de verkeerde map te halen: rechtermuisknop › **Ordenen** › **Verplaatsen**; het venster opent op **Voorgesteld**: klik naast *Huidige locatie* op **NovaDepot**, kies de map › **Verplaatsen** | [Drive 2375091](https://support.google.com/drive/answer/2375091?hl=nl), nagelezen 13-10-2026 · jouw scherm (08-10 en 09-10-2026) |
| Download terugvinden | rechts naast de adresbalk de downloads › muis op de bestandsnaam › **Weergeven in map** | [Chrome 95759](https://support.google.com/chrome/answer/95759?hl=nl&co=GENIE.Platform%3DDesktop), nagelezen 13-10-2026 (niet op een scherm gezien) |
| Naam wijzigen | rechtermuisknop › **Naam wijzigen**; de hele naam staat geselecteerd; **Home** zet de cursor vooraan en de naam blijft staan › **OK** (of **Annuleren**) | jouw scherm (09-10-2026) |
| Sorteren | rechts boven de lijst **Sorteren** › onder *Sorteren op* **Naam**, onder *Sorteerrichting* **A tot Z** (of **Z tot A**); onder *Mappen* **Bovenaan** · **Tussen de bestanden** | jouw scherm (09-10-2026) |
| Volgorde van Z tot A | letters komen vóór cijfers: de foto zonder datum staat boven de bonnen; de keuze geldt in alle mappen | jouw scherm (09-10-2026) · je demomap (12-10-2026) |
| Inleveren | **Inleveren** · **Inleveren ongedaan maken** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl) |

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1 terwijl de computers opstarten (dia 2 sla je over: die is voor het tweede uur). Dia 3 lesstart (3'): *waar zit de factuur van Vellekens, en hoe heet ze?* Dia 4–7 (5'30"): lesdoel, drie mappen en drie manieren, de puzzel *waarom jaar-maand-dag?*, de demo met **Home**. Dia 8 controle (1'): *wat doe je eerst?* (Home). |
| **Samen** | **3'** | Dia 9: iedereen doet stap 1 terwijl jij de schermen ziet; duim op als ze klaar zijn. |
| **Keuzewerktijd** | **tot het einde** | Dia 10 blijft staan, met **kloktijden** uit de planning (`_tools/zet_lesmomenten.py`; T rekent opnieuw); stap 2 tot 6, inleveren inbegrepen. Dia 11 in de laatste minuut. |

**Tweede uur** (sinds 13-10-2026): dia 1 terwijl ze aanmelden, dan dia 2 *Waar ben je gebleven?*. In hun map zien ze bij welke
stap ze verder moeten. Laat dia 2 staan; dia 11 in de laatste minuut. Bij het eerste uur sla je dia 2 over.

Je doet alleen **de datum vooraan** voor. Het zip-bestand, het uitpakken en het verplaatsen kennen ze uit les 06;
nieuw is alleen dat ze uitpakken in de **Verkenner** in plaats van in de app Bestanden van de Chromebook.

**Eerste rondgang, kijk naar twee dingen:**

1. Heeft iedereen een map NovaDepot met drie mappen? Wie niet: stap 1, handeling 4 (`NovaDepot.zip`).
2. Staan in stap 2 de zes **losse** bestanden in NovaDepot, en niet het zip-bestand?

**Tweede rondgang, rond stap 4 en 5:** laat een nieuwe naam luidop lezen. *"Staat het jaar vooraan? Staat de
nieuwste bon bovenaan?"* · *"Welk nummer koos je: de datum of het factuurnummer?"*

Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert.

---

## 4. Verbetersleutel

De leerling levert alleen het werkdocument in. De mappen zie je tijdens de rondgang. (Les 08, waarin ze hun map
NovaDepot met jou zouden delen, vervalt sinds 13-10-2026.)

### Deel 1 — Leveringsbonnen: op datum

| Van wie? | Datum | Nieuwe naam |
|---|---|---|
| Bakkerij Korstjes | 2026-10-13 | `2026-10-13 Leveringsbon - Bakkerij Korstjes.docx` (voorbeeld, al ingevuld) |
| Fruitgroothandel Appelbloesem | 2026-09-28 | `2026-09-28 Leveringsbon - Fruitgroothandel Appelbloesem.docx` |
| Zuivelhoeve De Melkkan | 2026-10-02 | `2026-10-02 Leveringsbon - Zuivelhoeve De Melkkan.docx` |

In Drive (Z tot A): de foto, dan Korstjes, Melkkan, Appelbloesem. De foto staat bovenaan omdat ze geen datum heeft
(van Z tot A komen letters vóór cijfers); in *Extra* krijgt ze er een.

### Deel 2 — Facturen: op nummer

| Van wie? | Factuurnummer | Nieuwe naam |
|---|---|---|
| Papierhandel Vellekens | 2026-0417 | `2026-0417 Factuur - Papierhandel Vellekens.docx` (voorbeeld, al ingevuld) |
| Fruitgroothandel Appelbloesem | 2026-0398 | `2026-0398 Factuur - Fruitgroothandel Appelbloesem.docx` |
| Zuivelhoeve De Melkkan | 2026-0409 | `2026-0409 Factuur - Zuivelhoeve De Melkkan.docx` |

Op de facturen staat ook een datum (29-09-2026 en 05-10-2026). Wie die gebruikt, heeft de sleutel niet begrepen.

Kleine verschillen zijn **goed**: hoofdletters, een spatie meer of minder, `.docx` die al ontbrak sinds les 06.
Wat moet kloppen: de sleutel vooraan (jaar-maand-dag, of het factuurnummer), en de naam van les 06 erachter.

### Deel 3 — De vragen

| Vraag | Waar het om gaat |
|---|---|
| 1 | Een **alfabetisch** klassement. Alle namen beginnen met *Prijslijst -*, dus beslist de naam van de leverancier. Die namen kregen ze al in les 06 (Wat - Van wie). |
| 2 | Drive leest de naam van links naar rechts. Met dag-maand-jaar komt `02-10` vóór `28-09`, terwijl 28 september ouder is. Met het jaar eerst, dan de maand, dan de dag, klopt de volgorde: met Z tot A staat de nieuwste bovenaan. Ook goed: *"Van Z tot A komt 28-09 vóór 02-10, maar 2 oktober is nieuwer."* |
| 3 | Eén reden: je vindt een document snel terug (ook een collega) · je ziet wat er ontbreekt · een bedrijf moet zijn facturen jaren bewaren. |

### Extra

- a) `2026-10-13 Foto - Speelgoed Tolletje.png` (de datum staat op het etiket).
- b) `2026-10-08 Leveringsbon - Sapfabriek Fris.docx`, tussen de bon van Zuivelhoeve De Melkkan (2026-10-02) en die
  van Bakkerij Korstjes (2026-10-13).

### Essentiële fouten — geef hier altijd feedback op

- De datum vooraan als dag-maand-jaar (`28-09-2026 …`): de volgorde klopt niet.
- Bij een factuur de datum gebruikt in plaats van het factuurnummer.
- De naam van les 06 is weg: alleen nog een datum of een nummer.
- Het zip-bestand geüpload in plaats van de zes bestanden.

---

## 5. Schermafbeeldingen en filmpjes

Alles is op 09-10-2026 opgenomen in jouw Chrome (profiel Jonas), in een demomap: het account rechtsboven staat niet
in beeld, Mac-sneltoetsen in de menu's zijn afgedekt. De lijst staat in `assets/screenshots/LEESMIJ.md`.

| Bestand | Waar |
|---|---|
| `video/stap3-verplaatsen.mp4` | niet meer op de pagina sinds 13-10-2026: stap 3 werkt met slepen (tekeningen) |
| `video/stap4-datum-vooraan.mp4` | stap 4: Naam wijzigen › Home › datum › OK |
| `video/zip-downloaden.mp4` | stap 2, onder de handelingen (uit les 06; een ander zip-bestand, dezelfde klikken) |
| `screenshots/stap1-map-uploaden.png` | stap 1, handeling 4 (inhalen); ook als stapje `stapjes/map-uploaden.png` |
| `screenshots/stap2-bestand-uploaden.png` | stap 2 |
| `screenshots/stap4-leveringen-klaar.png`, `stap5-facturen-klaar.png`, `stap6-prijslijsten.png` | stap 4, 5 en 6: zo ziet de map er daarna uit |
| `screenshots/stap4-sorteren.png` | stap 4, handeling 5: Sorteren › Naam › Z tot A (opnieuw opgenomen op 12-10-2026, net als de drie beelden van het resultaat) |

Er zijn geen beelden van de **Verkenner** (uitpakken): die kan ik vanaf je Mac niet maken. Je vroeg op 13-10-2026 een
opname van het hele proces (downloaden, uitpakken, uploaden): neem ze één keer op in lokaal 18, dan maak ik er een
filmpje met nummers van. Wil je ze toch, maak
dan in lokaal 18 een knipsel (**Windows** + **Shift** + **S**) en zet het in `assets/screenshots/`.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt het werkdocument, de zes nieuwe documenten, `Nieuwe-documenten.zip` en `NovaDepot.zip` opnieuw. Vereist
`python-docx` en `openpyxl`. Lees eerst de opmerkingen bovenaan het script: de datums en nummers zijn zo gekozen
dat de volgorde zonder sleutel fout is en met sleutel juist.

---

## 7. Leerplandoelen in je jaaroverzicht

De repository krijgt dezelfde `pre-push` hook als de andere lessen: bij elke push roept hij
`_tools/update_leerdoelen.py` aan met `lesdoelen.json`. Het blad **ORLO** kent alle codes van deze les
(`BK2_02.09`, `.01`, `.02`, `.03`, `BV2_04.03`, `BK2_02.06.02`, `BK2_01.02`; nagekeken op 09-10-2026).

---

## 8. Wat nog moet blijken in de klas

- **Bijgestuurd op 13-10-2026**: lukt **slepen** met de muis in lokaal 18, en belanden er bestanden in de verkeerde map?
  Staat **Weergeven in map** in de downloads van Chrome op de computers van lokaal 18 (help, niet op een scherm gezien),
  en opent de Verkenner bij het zip-bestand? Volgen ze de **grotere startpagina** beter? Vinden ze met **dia 2** snel
  waar ze waren?

- **De nieuwste bovenaan** (12-10-2026): vinden ze *Z tot A*? Storen ze zich aan de foto bovenaan en aan de prijslijsten
  van Z tot A?
- **Stap voor stap / Kort** (overgenomen van MW les 03, 11-10-2026): kiezen leerlingen zelf *Kort*? Wordt de
  pagina met alle beelden te lang? Vinden wie les 06 miste handeling 4 van stap 1?
- Komen ze vlotter op gang na de controledia en *Iedereen doet stap 1*?
- **Het nieuwe werkdocument** (12-10-2026): geen naam/klas/datum en geen *Zo werk je* meer, antwoorden in lichtblauwe
  vakken. Typen ze vlot in de vakken van deel 1 en 2? Missen ze de uitleg die nu alleen op de lespagina staat?


1. **Uitpakken in lokaal 18.** De help zegt alleen *Alles uitpakken… en volg de instructies*. De pagina zegt
   daarom *bevestig in het venster dat opengaat*. Heet de nieuwe map in Downloads echt `Nieuwe-documenten`?
2. **De toets Home.** Op een Belgisch toetsenbord staat er soms een pijl ↖ of *Début* op. Onder stapje b van
   stap 4 staat de terugweg: de pijl naar links ←. Sinds 12-10-2026 staat er een foto van het toetsenbord in
   het lokaal bij (`assets/stapjes/home-toets.jpg`, Home in een rood kader): in stap 4, stap 5 en op de demodia.
3. **Map uploaden (inhalen).** Chrome vraagt of je de bestanden wil uploaden; de pagina zegt *bevestig*. Ga na
   of dat bij een leerlingaccount ook zo loopt.
4. **Twee schrijfwijzen van de datum.** De bon van Korstjes (uit les 06) schrijft *13 oktober 2026*, de nieuwe bonnen
   *28-09-2026*. Dat is bewust: rij 1 en de demo tonen de eerste, de leerling doet de tweede zelf.
5. **4ORLOa** heeft les 07 op woensdag (eerste uur) en donderdag (tweede uur, sinds 13-10-2026).
6. **Haalbaarheid.** Stap 2 (downloaden, uitpakken, uploaden) is het riskante stuk. Loopt het vast, laat de
   leerlingen dan in stap 5 maar één factuur doen.
7. **Nieuw sinds 09-10-2026, voor alle klassen**: de vaste startdia *Zo start je* (dia 1) en de startpagina met
   *Wat heb je nodig?* in beelden en *Vast?* als rij. Geraken de leerlingen zo zonder hulp tot bij de opdracht,
   de lespagina en het werkdocument? (`_afspraken/presentatie.md`, `_afspraken/lespagina.md`)
