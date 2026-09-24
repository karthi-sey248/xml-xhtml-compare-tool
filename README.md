# XML Compare Tool

A tool designed to compare two XML files (Before & After), identify differences, and display changes clearly. It consists of two independent components:

1. **`index.html`** — A browser-based, side-by-side visual diff viewer (Left = Before, Right = After, highlighted changes with click-to-view popups).
2. **`xml_compare_to_excel.py`** — A Python script that generates a detailed Excel report (S.No, line numbers, attribute/text-level changes) along with annotated XML copies.

> **No installation required for the HTML viewer** — No additional local libraries or dedicated backend servers are needed.

---

## 1. Browser Viewer (`index.html`) — Local & Server Deployment

**Key Features:**
* **Instant Action:** As soon as both files are uploaded, download buttons become ready (no need to click the Compare button first).
* **✦ Pretty Print:** Works like Notepad++'s "XML Tools → Pretty Print" plugin. It indents and reformats loaded BEFORE/AFTER XML files (2-space indent, each element on a separate line, leaf values kept single-line) and automatically refreshes the comparison. Use this to avoid fake line-break/indentation differences when comparing minified vs formatted files.
* **Ignore ID Changes Checkbox:** When checked, it ignores differences in auto-generated ID attributes (like `id="MPS_d1e5"`) in element/attribute-level detailed reports (Excel & HTML). *(Note: Does not affect the raw line-by-line visual diff).*
* **Download Excel Report:** Generates and downloads a detailed `.xlsx` report directly from the browser without needing Python. Added/Removed changes highlight the full row (green/red), while Modified types highlight the row with a subtle color.
* **Download HTML Report:** Generates a standalone, color-coded `.html` report that can be opened in any browser or shared via email. Features top statistics (total mistakes, added/removed/attribute/text counts) and detailed tables with **exact word-level highlighting** (green = added word, red strikethrough = removed word) for modified changes.
