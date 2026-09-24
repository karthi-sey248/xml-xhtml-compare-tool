# XML Compare Tool

Rendu XML files (Before & After) compare pannitu, changes ah kandu pidichu
kaatura tool. Rendu parts irukku:

1. **`index.html`** — Browser la open panni use pannura, side-by-side visual
   diff viewer (left = Before, right = After, changes highlight, click
   pannina popup varum).
2. **`xml_compare_to_excel.py`** — Python script, detailed Excel report
   (S.No, line number, attribute/text level changes) + annotated XML copies
   generate pannum.

No installation required for the HTML viewer — vera edhavadhu library venaam,
server venaam.

---

## 1. Browser Viewer (`index.html`) — Local Deploy

**New:** Ippo `index.html` la irundhe rendu download buttons + oru "Pretty
Print" button + "Ignore id changes" checkbox irukku — rendu files um upload
panna odane (Compare button click pannanum nu kudathillai, download buttons
ready aagidum):

- **"✦ Pretty Print"** — Notepad++ oda "XML Tools → Pretty Print" plugin
  mari, loaded panna BEFORE/AFTER XML rendaiyum properly indent panni
  reformat pannidum (2-space indent, ovoru element separate line-la, leaf
  values mattum single-line-la). Click pannina odane, compare-um auto-a
  refresh aagidum. Ithu use panna reason: rendu files-um structurally
  same-a irundhaalum, indentation/line-break mattum difference irundha
  (oru file minified-a irundhaalum, vera one formatted-a irundhaalum),
  namba raw line-by-line view-la avlo "changes" fake-a kaatum — pretty
  print panna appuram andha kind of fake difference varathu.
- **"Ignore id changes"** checkbox — check pannina, `id="MPS_d1e5"` mari
  auto-generated id attribute-oda difference-ah **element/attribute-level
  detailed compare** (Excel + HTML report rendaiyum) la ignore pannidum.
  (Idhu visual line-by-line view-ah affect pannathu, andha view raw text
  mattum.)
- **"Download Excel Report"** — click pannina, browser lendhu directly
  Excel (.xlsx) file download panniidalam. Idhu `xml_compare_to_excel.py`
  script oda **same detailed element/attribute-level report** thaan
  generate pannum (S.No, Before/After Line No, XML Path, Change Type,
  Before/After Value) — vera Python script run pannanum nu kudathillai
  ippo. Add/Remove type mistakes-ku full row highlight (green/red), aana
  **Modified** type mistakes-ku (oru value vera oru value-a maarina) row
  full-a color pannama, exact-a **enna word maarichirukko andha word
  mattum** highlight pannum.
- **"Download HTML Report"** — click pannina, oru standalone, colour-coded
  `.html` mistake report file download aagum. Idhu Excel/library edhuvum
  illama, edhavadhu browser la open panni parkalam, illana team member
  kitta email/share pannalam. Mela stats (total mistakes, added/removed/
  attribute/text change counts) + niche same detailed table, Modified
  mistakes-ku exact word-level highlight (green = added word, red-strike =
  removed word).

(Rendu engine-um same logic — Python vs JavaScript la separately port
pannirukom, verify pannitum test pannirukom, same result varum. Pretty
Print button mattum index.html-ku mattum specific, Python script-la
kidayathu.)

### Option A — Direct-a open pannunga (unga system la mattum use pannanumna)
`index.html` file ah double-click pannunga, adhu unga default browser la
open aagum. File upload pannitu "Compare" click pannunga.

### Option B — Unga system ah SERVER-a run panni, VERA SYSTEMS la irundhu access pannanumna

Idhu thaan unga ku venum use-case — unga computer ah oru chinna web server-a
run pannitu, same network (LAN / office WiFi) la irukura vera computers
"http://<unga-IP>:8000" nu browser la type panni open pannalam.

**Step 1 — Unga computer oda IP address ah kandu pidichunga:**

Windows:
```bash
ipconfig
```
(`IPv4 Address` nu paarunga, e.g. `192.168.1.25`)

Mac / Linux:
```bash
ifconfig
```
or
```bash
ip addr
```

**Step 2 — Project folder kulla poi, server ah start pannunga:**

```bash
cd xml-compare-tool
python3 -m http.server 8000
```

Idhu run aagum bothu, terminal open-a vachikonga (window close pannina
server nikkum).

**Step 3 — Vera system la irundhu access pannunga:**

Same network la irukura vera edhavadhu computer/laptop la, browser thirandhu
type pannunga:

```
http://<unga-computer-IP>:8000/index.html
```

Example: `http://192.168.1.25:8000/index.html`

**Note:**
- Rendu computers um **same network / WiFi** la irukanum (office LAN, same
  router).
- Unga system la Firewall irundha, port 8000 ah allow pannanum (Windows
  Firewall → Allow an app → port 8000 add pannunga, illana simple-a
  "Private network" ku allow pannunga).
- Server ah "always-on" ah vekkanumna (unga terminal close panna kooda
  server run aaganum), Linux/Mac la `nohup python3 -m http.server 8000 &`
  use pannunga, illana Windows la Task Scheduler / NSSM mari tool vachi
  background service-a run pannalam.
- Idhu chinna teams/office internal use ku sufficient. Company-wide,
  permanent hosting venumna, IT team kitta oru proper internal web server
  (IIS/Nginx) la idha deploy panna sollunga — `index.html` copy pannitu
  podradhu than, konjam kooda setup venaam.

### Option C — Company server / intranet la deploy pannanumna
`index.html` ah edho oru static file hosting (Nginx, Apache, IIS, internal
web server, SharePoint, illana simple S3/static bucket) la mattum copy
pannitu podunga — idhu pure client-side HTML/CSS/JS, backend venaam,
database venaam, API venaam.

---

## 2. Python Excel Report Script (`xml_compare_to_excel.py`)

### Requirements
- Python 3.8+
- `openpyxl` library

```bash
pip install openpyxl
```

### Run pannurathu

Script open pannunga, top la irukura **INPUT / OUTPUT SETTINGS** section la
unga file paths edit pannunga:

```python
BEFORE_XML_PATH = "path/to/before.xml"
AFTER_XML_PATH  = "path/to/after.xml"
OUTPUT_EXCEL_PATH = "path/to/output_report.xlsx"
```

Appuram run pannunga:

```bash
python3 xml_compare_to_excel.py
```

Output:
- Excel report (`.xlsx`) — S.No, Before/After Line No, XML Path, Change
  Type, Before Value, After Value. Element Added/Removed type mistakes -
  FULL row highlight (green/red). **Modified** type mistakes (oru value
  vera oru value-a maarina, e.g. Text Modified / Attribute Modified) -
  row full-a color pannama, exact-a **enna word maarichirukko andha word
  mattum** (rich text, multiple colors same cell-la) highlight pannum -
  e.g. "Rylander P.N." → "P.N. Rylander" nu maarina, "Rylander"/"P.N."
  mattum red/green-a highlight aagum, mothama sentence-um illa.
- Standalone colour-coded HTML mistake report (`.html`) — same rows, same
  word-level highlight logic, ஆனா Excel venaam edhavadhu browser la open
  panni parkalam/share pannalam. `OUTPUT_HTML_PATH` variable-la path set
  pannalam (script top la INPUT/OUTPUT SETTINGS section), `None` vachu
  disable-um pannalam.
- Annotated copies of before/after XML — ovoru change irukura line mela
  `<!-- >>> CHANGE #S.No ... -->` comment potu kaatum

`IGNORE_ATTRIBUTES = ["id"]` nu set pannina, `id="MPS_d1e5"` mari
auto-generated id difference report la varathu — `index.html` la irukura
"Ignore id changes" checkbox-oda same effect.

---

## Folder Structure

```
xml-compare-tool/
├── index.html                 # Browser-based visual diff viewer
├── xml_compare_to_excel.py    # Python Excel report generator
└── README.md                  # This file
```

## Notes
- `index.html` ku, comparison work panna internet connection venaam
  (client-side thaan) — aana Excel download button ku, oru chinna library
  (SheetJS) cdnjs.cloudflare.com lendhu load aagum, athanaala andha ஒரு
  feature ku mattum internet/network access venum (unga office network
  la internet irundha, problem illa).
- Files unga computer vittu vela poagathu — server ah host pannalum,
  processing browser kulla thaan nadakum, files server ku upload aagathu.
- Rendu tool um independent-a use pannalam, illana together um use
  pannalam (Browser la Excel download pannitu, detailed script venumna
  andha `.py` file um separate-a use pannalam).
- `index.html` la irukura "Download Excel Report" (browser, SheetJS
  library) - Modified rows-ku exact word-level highlight kidayathu
  (SheetJS free/community library-la, oru cell-kulla multiple colors
  vecha rich-text feature paid "Pro" version-la thaan irukku), adhanaala
  andha Excel-la Modified rows full-row light color-ah mattum kaatum.
  Word-level exact highlight venumna, "Download HTML Report" button
  (rendu-um la irukku) illana `xml_compare_to_excel.py` (openpyxl rich
  text support pannum) use pannunga.
