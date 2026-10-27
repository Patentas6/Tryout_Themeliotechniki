# AGENTS_v1.md — Ροή Εργασίας Pricing Αγκυρώσεων & Berlin Wall (Berliner Wall)

> **Έκδοση: v1.0 (2026-10) — Επίσημο Πρότυπο Εταιρείας για Berlin Wall & Αγκυρώσεις**
> Αυστηρή τήρηση της αρχιτεκτονικής 2+1 Φύλλων, διατήρηση των 4 Πυλώνων και απόλυτη συμμόρφωση με τους κανόνες μορφοποίησης Excel.

Read this whole file and `RULES.md` (ή `RULES_v1.md`) before doing anything.

---

## 0. IMMUTABILITY & REPOSITORY PROTOCOL

* **Κεντρικό Αποθετήριο:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Rules_Berlin_Wall`
* **Αυστηρή Απαγόρευση Αλλοίωσης:** Κανένας AI agent σε οποιοδήποτε νέο chat δεν επιτρέπεται να αλλάξει τη μεθοδολογία, τα πλάτη των στηλών, τα χρώματα, τους τύπους, τα script ή τους 4 πυλώνες. Όλα τα άλλα chats αντιμετωπίζουν αυτόν τον φάκελο ως **αυστηρά READ-ONLY νόμο**.

---

## Your role (and its limits)

| You DO | You NEVER |
|---|---|
| **ΠΑΝΤΑ δημιουργείς Excel με 2+1 Φύλλα:** Φύλλο 1 = Στοιχεία & Φυσικές Ποσότητες (ΧΩΡΙΣ €), Φύλλο 2 = Υπολογισμοί Ευρώ (€), Φύλλο 3 = Βιβλιοθήκη Διατομών Χάλυβα | Mix raw inputs with calculated formulas or mix physical units with € on Sheet 1 |
| **Τιμές Μονάδας (€):** ΠΑΝΤΑ με **ΚΙΤΡΙΝΟ γέμισμα `#FFF2CC`** (Στήλη D Φύλλου 2) | Leave unit price cells white or unformatted without the mandatory yellow fill |
| **Ελλείπουσες / Εκτιμώμενες Τιμές:** ΠΑΝΤΑ με **ΚΟΚΚΙΝΟ γέμισμα `#FFC7CE` / `#9C0006`** και επεξήγηση στη Στήλη 6 | Crash, halt, leave empty cells or display `#N/A` if a material price isn't in the pricelist |
| **Κουγκρί Πάκτωσης (120 €/m³):** Διακριτό χρώμα `#FFE699` για εύκολη αναγνώριση & αλλαγή | Hardcode concrete price without making it visually identifiable for master review |
| **Πολλαπλές Ζώνες:** Αθροίζεις όλες τις οριζόντιες σειρές αγκυρώσεων ($\sum \text{Σειρές}$). Στράτες αγκυρώσεων = $\sum \text{Σειρές}$, Shotcrete = $\sum \text{Σειρές} + 1$ | Calculate trips for each zone separately without cumulative summing |
| **Καθαρό Σύνολο:** Άθροισμα 4 Πυλώνων = Καθαρό Σύνολο Προσφοράς (χωρίς ξεχωριστά Γενικά Έξοδα & Κέρδος) | Add separate overhead/profit rows when unit rates already include them |
| **KPIs:** Υπολογίζεις και εμφανίζεις: €/m² μετώπου, €/αγκύρωση, €/m² για Σχάρες+Shotcrete, €/οπή για H-Beams | Skip unit cost KPIs requested by the master and structural engineers |
| Keep versioned copies of `RULES_v*.md` and `AGENTS_v*.md` whenever rules change, never altering older versions | Overwrite or delete past version files |

---

## Υποχρεωτικό Python Boilerplate Μορφοποίησης Openpyxl

Κάθε script Python που δημιουργεί το Excel **πρέπει υποχρεωτικά να περιέχει και να εφαρμόζει τα εξής στυλ**:

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Στυλ Γραμματοσειρών
F_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
F_CHAP_HDR = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_SEC_HDR = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_SUBSEC_HDR = Font(name="Calibri", size=10, bold=True, color="1F4E78")
F_TBL_HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_BOLD = Font(name="Calibri", size=11, bold=True, color="000000")
F_REGULAR = Font(name="Calibri", size=11, color="000000")
F_ITALIC = Font(name="Calibri", size=10, italic=True, color="595959")
F_TIP = Font(name="Calibri", size=9, italic=True, color="555555")

# Κόκκινο Alert για ελλείπουσες / εκτιμώμενες τιμές
F_RED_ALERT = Font(name="Calibri", size=11, bold=True, color="9C0006")
FILL_RED_ALERT = PatternFill("solid", fgColor="FFC7CE")

# Χρώματα Γεμίσματος
FILL_TITLE = PatternFill("solid", fgColor="1F4E78")
FILL_CHAP_A = PatternFill("solid", fgColor="1E4D2B")      # Σκούρο πράσινο Κεφαλαίου Α
FILL_CHAP_B = PatternFill("solid", fgColor="2F5597")      # Σκούρο μπλε Κεφαλαίου Β
FILL_SEC = PatternFill("solid", fgColor="41719C")         # Μπλε ενοτήτων
FILL_SUBSEC = PatternFill("solid", fgColor="D9E1F2")      # Απαλό μπλε υποενοτήτων
FILL_TBL_HDR = PatternFill("solid", fgColor="305496")     # Ναυτικό μπλε πινάκων
FILL_EDITABLE = PatternFill("solid", fgColor="E2EFDA")    # Πράσινο: Εισαγωγή Μελέτης (Single Point of Truth)
FILL_CALC = PatternFill("solid", fgColor="F2F2F2")        # Γκρι: Υπολογισμοί
FILL_TOTAL = PatternFill("solid", fgColor="C6EFCE")       # Έντονο πράσινο: Σύνολα
FILL_YELLOW = PatternFill("solid", fgColor="FFF2CC")      # ΥΠΟΧΡΕΩΤΙΚΟ ΚΙΤΡΙΝΟ: Τιμές Μονάδας
FILL_SPECIAL = PatternFill("solid", fgColor="FFE699")     # Διακριτό κίτρινο/πορτοκαλί: Κουγκρί Πάκτωσης 120€

# Περιγράμματα
THIN_GRAY = Side(style="thin", color="D9D9D9")
MED_BLUE = Side(style="medium", color="1F4E78")
DOUBLE_BOTTOM = Side(style="double", color="1F4E78")

BORDER_CELL = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
BORDER_HEADER = Border(left=THIN_GRAY, right=THIN_GRAY, top=MED_BLUE, bottom=MED_BLUE)
BORDER_TOTAL = Border(top=THIN_GRAY, bottom=DOUBLE_BOTTOM, left=THIN_GRAY, right=THIN_GRAY)

# Πλάτη Στηλών
COL_W_1 = {"A": 48, "B": 18, "C": 14, "D": 22, "E": 22, "F": 22}
COL_W_2 = {"A": 44, "B": 18, "C": 14, "D": 22, "E": 22, "F": 26, "G": 28}
COL_W_3 = {"A": 20, "B": 16, "C": 14, "D": 14, "E": 14, "F": 16, "G": 18}

# Υποχρεωτικά Gridlines
ws1.views.sheetView[0].showGridLines = True
ws2.views.sheetView[0].showGridLines = True
ws3.views.sheetView[0].showGridLines = True
```

---

## Workflow — follow these steps in order

### Step 0 — Create the job folder & files
`jobs/<YYYY-MM>_<short_name>/` with `input/` (save email text / PDFs / plans there unchanged).

### Step 1 — Sheet 1: 2 Big Chapters (Physical Data ONLY)
- **ΚΕΦΑΛΑΙΟ Α: ΜΟΝΟ ΣΤΟΙΧΕΙΑ ΠΟΥ ΔΙΑΒΑΣΤΗΚΑΝ ΑΠΟ ΤΗ ΜΕΛΕΤΗ**
  - Τοποθεσία, Έδρα (Άγιος Συλάς), Χιλιόμετρα, Σειρές Αγκυρώσεων ανά Ζώνη.
  - Γεωμετρία: Μήκη $L$, Ύψος Εκσκαφής $H_{\text{exc}}$.
  - Σχάρες: Διάμετρος mm, Βήματα mm.
  - Shotcrete: Ονομαστικό πάχος $m$ (π.χ. 0.12), Πάγια απώλεια $0.05\text{ m}$.
  - H-Beams: Πλήθος οπών $N_{\text{holes}}$, Διάμετρος οπής (800 mm default), Ύψος πάκτωσης $H_{\text{embed}}$, Διατομή προφίλ & πλήθος ανά οπή, Λάμες ($W \times L_{\text{pl}} \times t$ & βήμα), Γωνιές (προαιρετικά).
  - Αγκυρώσεις: Σειρές, Βάθος ανά σειρά, Πλήθος ανά σειρά, Strands ανά οπή, Διάμετρος οπής 12cm, Grout $\times 1.20$, $+1.0\text{ m}$ τένοντα.
  - Όλα σε πράσινα κελιά `#E2EFDA` (Single Point of Truth).
- **ΚΕΦΑΛΑΙΟ Β: ΟΛΟΙ ΟΙ ΥΠΟΛΟΓΙΣΜΟΙ ΦΥΣΙΚΩΝ ΜΕΓΕΘΩΝ**
  - Στράτες: Μεγάλο τρυπάνι = 1, Αγκυρώσεις = $\sum \text{Σειρές}$, Shotcrete = $\sum \text{Σειρές} + 1$.
  - Σχάρες & Shotcrete: $A = L \times H_{\text{exc}}$, $kg$ σχάρας, ονομαστικά $m^3$ shotcrete (χωρίς $+5\text{ cm}$), % απώλειας.
  - H-Beams: Μέτρα διάτρησης $M_{\text{drill}}$, $m^3$ διάτρησης, $kg$ προφίλ νέφκας, $kg$ λαμών, $m^3$ κουγκριού πάκτωσης, $kg$ γωνιών.
  - Αγκυρώσεις: Συνολικά τεμάχια, συνολικά μέτρα διάτρησης, $m^3$ grout, μέτρα & κιλά τενόντων, τεμάχια προέντασης & κεφαλών.
- **ΚΑΝΕΝΑ ΕΥΡΩ (€) ΣΤΟ ΦΥΛΛΟ 1.**

### Step 2 — Sheet 2: Pricing across the 4 Pillars
- **Πυλώνας 1: Μετακινήσεις**
  - Μεγάλο τρυπάνι: 1,500 € Λεμεσός (+500 € 40-80km, +1500 € >80km).
  - Εξοπλισμός αγκυρώσεων: 300 € / στράτα (+50 € / 40km).
  - Συνεργείο Shotcrete: 200 € / στράτα (+50 € / 40km).
- **Πυλώνας 2: Σχάρες & Shotcrete**
  - Αγορά χάλυβα σχάρας (2.50 €/kg) & Τοποθέτηση (15.00 €/m²).
  - Shotcrete C30/37 αγορά (90.00 €/m³) & Τοποθέτηση με απώλεια (€/m³).
- **Πυλώνας 3: H-Beams (Νέφκες)**
  - Διάτρηση μεγάλου τρυπανιού (65–90 €/m, default 75.00 €/m).
  - Νέφκες χάλυβας (1.40 €/kg = 1.00 € αγορά + 0.40 € επεξεργασία).
  - Λάμες σύνδεσης (1.60 €/kg = 1.00 € αγορά + 0.60 € τοποθέτηση).
  - Κουγκρί πάκτωσης (120.00 €/m³ [διακριτό `#FFE699`]).
  - Γωνιές στήριξης (1.60 €/kg = 1.00 € αγορά + 0.60 € επεξεργασία - αν ζητούνται).
- **Πυλώνας 4: Αγκυρώσεις**
  - Διάτρηση αγκυρώσεων (25–35 €/m, default 30.00 €/m).
  - Grout τσιμεντένεμα (500.00 €/m³).
  - Τένοντες (2.28 €/m τένοντα).
  - Προένταση (18.00 €/τεμ).
  - Κεφαλή αγκύρωσης (55.00 €/τεμ).
- **Τελικά Σύνολα & KPIs:**
  - Καθαρό Σύνολο Προσφοράς = Άθροισμα 4 Πυλώνων (`=E8+E13+E19+E27` κ.λπ.).
  - Φ.Π.Α. (19%) & Τελικό Πληρωτέο Ποσό.
  - KPI 1: Συνολικό Κόστος ανά m² Μετώπου (€/m²).
  - KPI 2: Συνολικό Κόστος ανά Αγκύρωση (€/τεμ).
  - KPI 3: Συνολικό Κόστος Σχαρών & Shotcrete ανά m² Μετώπου (€/m²).
  - KPI 4: Συνολικό Κόστος H-Beams ανά Πάσσαλο/Οπή (€/οπή).

### Step 3 — Sheet 3: Steel Profiles Database
- Ενσωματωμένος πλήρης πίνακας διατομών HEA, HEB, UPN και Γωνιών (L).

### Step 4 — Report to User
- Παρουσίαση στον χρήστη με καθαρή δομή:
  1. Συνοπτική εικόνα Φύλλου 1 (Κεφάλαιο Α & Κεφάλαιο Β).
  2. Συνοπτική εικόνα Φύλλου 2 (Ανάλυση Κόστους 4 Πυλώνων & Τελικά Σύνολα).
  3. Πίνακας KPIs (€/m², €/αγκύρωση, €/οπή).
  4. Σύνδεσμος προς το νέο αρχείο Excel.
