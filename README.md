# Handleiding voor de leraar — Les 07: Het digitale klassement

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, fictief logistiek bedrijf en groothandel — module 2 van het jaarplan
(`Leerplannen/Lessenplan_TOINFO_ORLO_arbeidsmarkt_2026-2027.md`)
**Toestel:** **lokaal 18, Windows 11** — Google Workspace in Chrome en de Verkenner
**Lesdag:** uur 1 van week 7: 4ORLOb di 13 oktober (8ste uur) · 3ORLO di 13 oktober (9de uur) · 4ORLOa wo 14 oktober (4de uur)
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
> zetten met **Nieuw › Map uploaden** duurt ± 2 minuten. Het staat bij *Hulp nodig?* in stap 1, met een
> schermafbeelding. Wie een map heeft die niet klopt, geeft die eerst de naam *NovaDepot oud*.

---

## 1. Inhoud van het pakket

```text
W07 - Les 07 - ORLO - Het digitale klassement/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 8 dia's voor de lesstart en "Ik doe"
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
│   └── maak_werkdocumenten.py             # maakt alles opnieuw (python-docx, openpyxl)
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

- 4 tot 6 handelingen per stap, één tip en één *Hulp nodig?* per stap.
- **Stap 3** en **stap 4** hebben een filmpje van 13 seconden dat vanzelf herhaalt, met nummers die bij de
  handelingen horen. Stap 4, 5 en 6 tonen ook een schermafbeelding van het resultaat: *zo ziet je map er daarna uit*.
- De **naamkaart** in stap 4 en 5 toont in kleur wat er vooraan komt (de datum, het nummer) en wat blijft staan.
- De **zelftest** in stap 6 heeft drie vragen, waarvan één de valkuil is: van A tot Z staat `02-10-2026` vóór
  `28-09-2026`.
- Windows 11 (lokaal 18). `localStorage` bewaart alleen de vinkjes en de laatste stap (voorvoegsel
  `novadepot_klassement_v1_`), met een wisknop.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 09-10-2026*

Repository [`jonasdaltongent/Klassement-NovaDepot`](https://github.com/jonasdaltongent/Klassement-NovaDepot),
Pages op branch `main`, map `/ (root)`. De lespagina staat op
[jonasdaltongent.github.io/Klassement-NovaDepot](https://jonasdaltongent.github.io/Klassement-NovaDepot/), de dia's op
[…/presentatie.html](https://jonasdaltongent.github.io/Klassement-NovaDepot/presentatie.html); dat adres staat
ook op dia 7, in `classroom.json` en in `lesdoelen.json`. Live bestanden nagekeken: gelijk aan de lokale.

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

Je doet op dia 6 de bon van Bakkerij Korstjes voor, in een eigen map NovaDepot. Pak `werkdocument/NovaDepot.zip`
uit en zet de map in je Drive met **Nieuw** › **Map uploaden**. Zo heb je dezelfde map als een leerling na les 06,
en heb je meteen de inhaalweg zelf eens gedaan.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 7 klopt.
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
| Verplaatsen | rechtermuisknop › **Ordenen** › **Verplaatsen**; het venster opent op **Voorgesteld**: klik naast *Huidige locatie* op **NovaDepot**, kies de map › **Verplaatsen** | jouw scherm (08-10 en 09-10-2026) |
| Naam wijzigen | rechtermuisknop › **Naam wijzigen**; de hele naam staat geselecteerd; **Home** zet de cursor vooraan en de naam blijft staan › **OK** (of **Annuleren**) | jouw scherm (09-10-2026) |
| Sorteren | rechts boven de lijst **Sorteren** › onder *Sorteren op* **Naam**, onder *Sorteerrichting* **A tot Z** (of **Z tot A**); onder *Mappen* **Bovenaan** · **Tussen de bestanden** | jouw scherm (09-10-2026) |
| Volgorde van A tot Z | cijfers komen vóór letters: de bonnen met een datum staan boven de foto | jouw scherm (09-10-2026) |
| Inleveren | **Inleveren** · **Inleveren ongedaan maken** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl) |

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1 terwijl de computers opstarten. Dia 2 lesstart (3'): *waar zit de factuur van Vellekens, en hoe heet ze?* Dia 3–6 (5'30"): lesdoel, drie mappen en drie manieren, de puzzel *waarom jaar-maand-dag?*, de demo met **Home**. Dia 7 *Zo werk je verder* (30"). |
| **Keuzewerktijd** | **40'** | Dia 7 blijft staan; de leerlingen werken stap 1 tot 6 af, inleveren inbegrepen. Dia 8 in de laatste minuut. |

Je doet alleen **de datum vooraan** voor. Het zip-bestand, het uitpakken en het verplaatsen kennen ze uit les 06;
nieuw is alleen dat ze uitpakken in de **Verkenner** in plaats van in de app Bestanden van de Chromebook.

**Eerste rondgang, kijk naar twee dingen:**

1. Heeft iedereen een map NovaDepot met drie mappen? Wie niet: stap 1, *Hulp nodig?* (`NovaDepot.zip`).
2. Staan in stap 2 de zes **losse** bestanden in NovaDepot, en niet het zip-bestand?

**Tweede rondgang, rond stap 4 en 5:** laat een nieuwe naam luidop lezen. *"Staat het jaar vooraan? Staat de
oudste bon bovenaan?"* · *"Welk nummer koos je: de datum of het factuurnummer?"*

Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert.

---

## 4. Verbetersleutel

De leerling levert alleen het werkdocument in. De mappen zie je tijdens de rondgang; in les 08 delen de
leerlingen hun map NovaDepot met jou (als kijker).

### Deel 1 — Leveringsbonnen: op datum

| Van wie? | Datum | Nieuwe naam |
|---|---|---|
| Bakkerij Korstjes | 2026-10-13 | `2026-10-13 Leveringsbon - Bakkerij Korstjes.docx` (voorbeeld, al ingevuld) |
| Fruitgroothandel Appelbloesem | 2026-09-28 | `2026-09-28 Leveringsbon - Fruitgroothandel Appelbloesem.docx` |
| Zuivelhoeve De Melkkan | 2026-10-02 | `2026-10-02 Leveringsbon - Zuivelhoeve De Melkkan.docx` |

In Drive: Appelbloesem, Melkkan, Korstjes, en onderaan de foto. Zonder datum stonden ze op naam, met de nieuwste
(Korstjes) bovenaan.

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
| 2 | Drive leest de naam van links naar rechts. Met dag-maand-jaar komt `02-10` vóór `28-09`, terwijl 28 september ouder is. Met het jaar eerst, dan de maand, dan de dag, staat de oudste bovenaan. |
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
| `video/stap3-verplaatsen.mp4` | stap 3: Ordenen › Verplaatsen › NovaDepot › Leveringen › Verplaatsen |
| `video/stap4-datum-vooraan.mp4` | stap 4: Naam wijzigen › Home › datum › OK |
| `video/zip-downloaden.mp4` | stap 2, *Hulp nodig?* (uit les 06; een ander zip-bestand, dezelfde klikken) |
| `screenshots/stap1-map-uploaden.png` | stap 1, *Hulp nodig?* (inhalen) |
| `screenshots/stap2-bestand-uploaden.png` | stap 2 |
| `screenshots/stap4-leveringen-klaar.png`, `stap5-facturen-klaar.png`, `stap6-prijslijsten.png` | stap 4, 5 en 6: zo ziet de map er daarna uit |
| `screenshots/stap4-sorteren.png` | stap 4, *Hulp nodig?*: Sorteren › Naam › A tot Z |

Er zijn geen beelden van de **Verkenner** (uitpakken): die kan ik vanaf je Mac niet maken. Wil je ze toch, maak
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

1. **Uitpakken in lokaal 18.** De help zegt alleen *Alles uitpakken… en volg de instructies*. De pagina zegt
   daarom *bevestig in het venster dat opengaat*. Heet de nieuwe map in Downloads echt `Nieuwe-documenten`?
2. **De toets Home.** Op een Belgisch toetsenbord staat er soms een pijl ↖ of *Début* op. Bij *Hulp nodig?* in
   stap 4 staat de terugweg: de pijl naar links ←.
3. **Map uploaden (inhalen).** Chrome vraagt of je de bestanden wil uploaden; de pagina zegt *bevestig*. Ga na
   of dat bij een leerlingaccount ook zo loopt.
4. **Twee schrijfwijzen van de datum.** De bon van Korstjes (uit les 06) schrijft *13 oktober 2026*, de nieuwe bonnen
   *28-09-2026*. Dat is bewust: rij 1 en de demo tonen de eerste, de leerling doet de tweede zelf.
5. **4ORLOa** heeft les 07 op woensdag en les 08 al op donderdag: wie stap 6 niet haalt, werkt thuis af vóór
   vrijdag 20.00 uur.
6. **Haalbaarheid.** Stap 2 (downloaden, uitpakken, uploaden) is het riskante stuk. Loopt het vast, laat de
   leerlingen dan in stap 5 maar één factuur doen.
7. **Nieuw sinds 09-10-2026, voor alle klassen**: de vaste startdia *Zo start je* (dia 1) en de startpagina met
   *Wat heb je nodig?* in beelden en *Vast?* als rij. Geraken de leerlingen zo zonder hulp tot bij de opdracht,
   de lespagina en het werkdocument? (`_afspraken/presentatie.md`, `_afspraken/lespagina.md`)
