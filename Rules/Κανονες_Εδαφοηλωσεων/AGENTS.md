# RULES_v8.md — Pricing Εδαφοηλώσεων (Soil Nailing) — Έκδοση v8

> **Έκδοση: v8.0 (2026-10)**
> *Ιστορικό Εκδόσεων:*
> - `AGENTS_v1.md`: Αρχική έκδοση ροής εργασίας (Ιστορικό / Ανέπαφο).
> - `AGENTS_v2.md`: Πρότυπο 2 Φύλλων Excel & Κανόνας Versioning (Ιστορικό / Ανέπαφο).
> - `AGENTS_v3.md`: Single Point of Truth δυναμικών τίτλων + Value Engineering (Ιστορικό / Ανέπαφο).
> - `AGENTS_v4.md`: Κεφάλαια Α & Β στο Φύλλο 1 + [Χ %] Μάστρου (Ιστορικό / Ανέπαφο).
> - `AGENTS_v5.md`: Ονομαστικά κυβικά Shotcrete χωρίς απώλεια + Ακριβείς στράτες 7 και 6 (Ιστορικό / Ανέπαφο).
> - `AGENTS_v6.md`: Ενσωμάτωση εξόδων/κέρδους στις τιμές μονάδας + Κόκκινο κελί ελλειπουσών τιμών + Immutability (Ιστορικό / Ανέπαφο).
> - `AGENTS_v7.md`: **Τρέχουσα Επίσημη Έκδοση Βάσης Εταιρείας** με:
>   1. **Απαράβατοι Κανόνες Μορφοποίησης Excel**: Αυστηρά προκαθορισμένα πλάτη στηλών, ύψη γραμμών και χρωματισμοί. Όλες οι τιμές μονάδας υποχρεωτικά με **ΚΙΤΡΙΝΟ `#FFF2CC`** (ή **ΚΟΚΚΙΝΟ `#FFC7CE` / `#9C0006`** αν λείπει τιμή).
>   2. **Πολλαπλές Ζώνες Εκσκαφής (Multi-Zone)**: Αθροιστικός υπολογισμός σειρών ($\sum \text{Σειρές}$). Στράτες ήλων = $\sum \text{Σειρές}$, Στράτες shotcrete = $\sum \text{Σειρές} + 1$.
>   3. **Απόλυτη Απαγόρευση Αλλοίωσης από άλλα Chats (Immutability Protocol)**.

Read this whole file and `RULES.md` (ή `RULES_v7.md`) before doing anything.

---

## 0. IMMUTABILITY, REPOSITORY & ACTIVE VERSION PROTOCOL

* **Κεντρικό Αποθετήριο:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Rules\Κανονες_Εδαφοηλωσεων`
* **Επίσημος Φάκελος Έργων:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Edafoilosis\<PROJECT_NAME>\`
* **Ανάγνωση Ενεργής Έκδοσης (Active Version Discovery):** Κάθε AI agent που ξεκινάει νέο chat **ΟΦΕΙΛΕΙ ΠΡΩΤΑ ΝΑ ΔΙΑΒΑΖΕΙ ΤΟ `ACTIVE_VERSION.json`** στον φάκελο `Rules\Κανονες_Εδαφοηλωσεων`. Το αρχείο αυτό καθορίζει ποια έκδοση είναι σε ισχύ (π.χ. `"active_version": "v7"`), ώστε το AI να φορτώνει το `RULES_v7.md` & `AGENTS_v7.md` και όχι λάθος έκδοση (v6 ή v8).
* **Κανόνας Απαγόρευσης Τροποποίησης Ενεργής Έκδοσης (Copy-on-Change):** ΠΟΤΕ δεν τροποποιείται η τρέχουσα ενεργή έκδοση in-place. Για οποιαδήποτε αλλαγή ή δοκιμή, αντιγράφουμε το αρχείο σε νέο αριθμό έκδοσης (π.χ. `RULES_v8.md`, `AGENTS_v8.md`) και δουλεύουμε αποκλειστικά εκεί.
* **Προώθηση σε Παραγωγή (Promotion):** Μόνο αφού ελεγχθεί και εγκριθεί η νέα έκδοση από τον χρήστη, ενημερώνεται το `ACTIVE_VERSION.json` σε `"active_version": "v8"` και συγχρονίζονται τα `RULES.md` / `AGENTS.md`.
* **Απαγόρευση Τροποποίησης από άλλα Chats:** Όλα τα άλλα chats αντιμετωπίζουν αυτόν τον φάκελο ως **READ-ONLY ΝΟΜΟ**.
* **ΑΠΟΛΥΤΟΣ ΚΑΝΟΝΑΣ ΑΠΟΜΟΝΩΣΗΣ ΕΡΓΟΥ & ZERO CROSS-PROJECT CONTAMINATION (BLIND TEST):** Το AI απαγορεύεται αυστηρά να ψάχνει, να διαβάζει ή να ανοίγει άλλα projects ή παλαιότερα αρχεία excel/προσφορών (για να μην «κλέβει»). Περιορίζεται αποκλειστικά στα αρχεία εισόδου της τρέχουσας μελέτης και στους φακέλους Rules\.
* **ΟΡΓΑΝΩΣΗ ΦΑΚΕΛΟΥ OLDER_VERSIONS:** Στον κεντρικό φάκελο Κανονες_Εδαφοηλωσεων μένουν ΜΟΝΟ τα αρχεία της ενεργής έκδοσης (v8) και το ACTIVE_VERSION.json. Όλα τα παλαιότερα μεταφέρονται στο older_versions\.

---

## Your role (and its limits)

| You DO | You NEVER |
|---|---|
| **Ελέγχεις ΠΡΩΤΑ το `ACTIVE_VERSION.json`:** Διαβάζεις ποια έκδοση είναι σε παραγωγή και ακολουθείς πιστά τα αντίστοιχα αρχεία (`RULES_v7.md` / `AGENTS_v7.md`) | Assume or guess rule versions without checking `ACTIVE_VERSION.json` |
| **Απαγόρευση In-Place Edit:** Αν ζητηθούν αλλαγές κανόνων, αντιγράφεις σε νέα έκδοση (`v8`) και κάνεις εκεί τις αλλαγές | Overwrite or alter the active version files (`v7`) directly |
| **ΠΑΝΤΑ εφαρμόζεις τις ακριβείς προδιαγραφές μορφοποίησης:** Πλάτη στηλών (Sheet 1: 48,18,14,22,22,22 / Sheet 2: 44,18,14,22,22,26,28) και ύψη γραμμών | Alter column widths, cell sizes, or invent your own layout |
| **Τιμές Μονάδας (€):** ΠΑΝΤΑ με **ΚΙΤΡΙΝΟ γέμισμα `#FFF2CC`** | Leave unit price cells white or unformatted without the mandatory yellow fill |
| **Πολλαπλές Ζώνες Εκσκαφής:** Αθροίζεις όλες τις οριζόντιες σειρές ($\sum \text{Σειρές}$). Ήλοι = $\sum$, Shotcrete = $\sum + 1$ | Calculate trips for each zone separately or miss the cumulative row count |
| **Φύλλο 1 σε 2 Μεγάλα Κεφάλαια:** Κεφάλαιο Α = ΜΟΝΟ ό,τι διαβάστηκε από τη μελέτη, Κεφάλαιο Β = Όλοι οι υπολογισμοί | Mix raw inputs with calculated formulas in the same section |
| **Shotcrete Κυβικά:** Υπολογίζεις κυβικά ΜΟΝΟ με ονομαστικό πάχος μελέτης ($V = A \times \text{πάχος}$), χωρίς $+5\text{ cm}$ | Add waste thickness to physical $m^3$ when it is already charged in the unit price (€/m³) |
| **Καθαρό Σύνολο:** Άθροισμα 4 Πυλώνων = Καθαρό Σύνολο (χωρίς ξεχωριστά Γενικά Έξοδα & Κέρδος) | Add separate overhead/profit rows when unit rates already include them |
| **Ελλείπουσες Τιμές (Red Alert):** Χρησιμοποιείς πλησιέστερη τιμή, βάφεις το κελί **ΚΟΚΚΙΝΟ (`#FFC7CE` / `#9C0006`)** και βάζεις επεξήγηση στη Στήλη 6 | Crash, halt, leave empty cells or display `#N/A` if a material price isn't in the pricelist |
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

# Κόκκινο Alert για ελλείπουσες τιμές
F_RED_ALERT = Font(name="Calibri", size=11, bold=True, color="9C0006")
FILL_RED_ALERT = PatternFill("solid", fgColor="FFC7CE")

# Χρώματα Γεμίσματος
FILL_TITLE = PatternFill("solid", fgColor="1F4E78")
FILL_CHAP_A = PatternFill("solid", fgColor="1E4D2B")
FILL_CHAP_B = PatternFill("solid", fgColor="2F5597")
FILL_SEC = PatternFill("solid", fgColor="41719C")
FILL_SUBSEC = PatternFill("solid", fgColor="D9E1F2")
FILL_TBL_HDR = PatternFill("solid", fgColor="305496")
FILL_EDITABLE = PatternFill("solid", fgColor="E2EFDA")    # Πράσινο: Εισαγωγή Μελέτης
FILL_CALC = PatternFill("solid", fgColor="F2F2F2")        # Γκρι: Υπολογισμοί
FILL_TOTAL = PatternFill("solid", fgColor="C6EFCE")       # Έντονο πράσινο: Σύνολα
FILL_YELLOW = PatternFill("solid", fgColor="FFF2CC")      # ΥΠΟΧΡΕΩΤΙΚΟ ΚΙΤΡΙΝΟ: Τιμές Μονάδας
FILL_VE = PatternFill("solid", fgColor="FCE4D6")          # Απαλό πορτοκαλί: Value Engineering

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

# Υποχρεωτικά Gridlines
ws1.views.sheetView[0].showGridLines = True
ws2.views.sheetView[0].showGridLines = True
```

---

## Workflow — follow these steps in order

### Step 0 — Create the job folder & files
`jobs/<YYYY-MM>_<short_name>/` with `input/` (save email text / PDFs / plans there unchanged).

### Step 1 — Sheet 1: 2 Big Chapters (Physical Data ONLY)
- **ΚΕΦΑΛΑΙΟ Α: ΜΟΝΟ ΣΤΟΙΧΕΙΑ ΠΟΥ ΔΙΑΒΑΣΤΗΚΑΝ ΑΠΟ ΤΗ ΜΕΛΕΤΗ**
  - Αν υπάρχουν πολλαπλές ζώνες εκσκαφής: Καταγράφονται οι σειρές κάθε ζώνης χωριστά (π.χ. Ζώνη 1: 4 σειρές, Ζώνη 2: 3 σειρές).
  - Όλα τα πρωτογενή δεδομένα συγκεντρωμένα στην αρχή σε πράσινα κελιά `#E2EFDA` (Single Point of Truth).
- **ΚΕΦΑΛΑΙΟ Β: ΟΛΟΙ ΟΙ ΥΠΟΛΟΓΙΣΜΟΙ ΦΥΣΙΚΩΝ ΜΕΓΕΘΩΝ**
  - Συνολικές Σειρές Μελέτης: `=SUM(Σειρές Ζωνών)` (π.χ. $4 + 3 = 7$).
  - Στράτες Ήλων: `= Συνολικές Σειρές` ($7$).
  - Στράτες Shotcrete: `= Συνολικές Σειρές + 1` ($7 + 1 = 8$).
  - Σχάρες: $A = L \times H$, βάρος $kg/m^2$ δυναμικά από διάμετρο, συνολικό βάρος $kg$.
  - Shotcrete: Ονομαστικός όγκος $V = A \times \text{Ονομαστικό Πάχος}$, % απώλειας $= \text{απώλεια} / \text{πάχος}$.
  - Εδαφοηλώσεις: Τεμάχια, μέτρα διάτρησης, grout $m^3$ ($\times 1.20$), βάρος ράβδων, πλακών, τεμάχια nuts.
- **ΚΑΝΕΝΑ ΕΥΡΩ (€) ΣΤΟ ΦΥΛΛΟ 1.**

### Step 2 — Sheet 2: Pricing, Yellow Unit Rates, Red Alert Fallbacks & Value Engineering
- **Εφαρμογή τιμών μονάδας ανά Πυλώνα:**
  - **Όλες οι τιμές μονάδας (Στήλη D) έχουν ΥΠΟΧΡΕΩΤΙΚΑ ΚΙΤΡΙΝΟ γέμισμα `#FFF2CC`**.
  - Αν υπάρχει υλικό χωρίς ακριβή τιμή (π.χ. άλλη κατηγορία σκυροδέματος/χάλυβα): Χρήση πλησιέστερης τιμής, **Κόκκινο κελί (`#FFC7CE` / `#9C0006`)** και σημείωση στη Στήλη 6.
  - Πυλώνας 1: Στράτες Shotcrete ($=\sum \text{Σειρές}+1$) και Στράτες Ήλων ($=\sum \text{Σειρές}$).
  - Πυλώνας 2: Χάλυβας σχάρας + Τοποθέτηση.
  - Πυλώνας 3: Shotcrete με ονομαστικά κυβικά, βάση 200€ (>7m³) ή 250€ (<=7m³), προσαύξηση $+(\text{απώλεια}/\text{πάχος})\%$, αγορά 90€, τοποθέτηση το υπόλοιπο.
  - Πυλώνας 4: Διάτρηση, Grout, Χάλυβας ράβδων, Πάσο, Πλάκες, Nuts.
- **Τελικά Σύνολα (Όλα μέσα στις τιμές):**
  - **Καθαρό Σύνολο Προσφοράς:** Άθροισμα Πυλώνων 1 + 2 + 3 + 4 (`=E8+E13+E18+E27`).
  - **Φ.Π.Α. (19%):** `=ROUND(Καθαρό*0.19, 2)`.
  - **Τελικό Πληρωτέο Ποσό:** `=Καθαρό + ΦΠΑ`.
  - **KPIs:** €/m² μετώπου και €/ήλο.
- **Value Engineering:** Πλήρης πίνακας σύγκρισης αναλογίας κόστους ανά ζώνη/διάταξη.

### Step 3 — Report to User
- Παρουσίαση στον χρήστη με καθαρή δομή:
  1. Συνοπτική εικόνα Φύλλου 1 (Κεφάλαιο Α & Κεφάλαιο Β).
  2. Συνοπτική εικόνα Φύλλου 2 (Ανάλυση Κόστους € & Τελικά Σύνολα).
  3. Επισημάνσεις Κόκκινων Κελιών (αν υπάρχουν ελλείπουσες τιμές).
  4. Πίνακας Value Engineering.
  5. Σύνδεσμος προς το νέο αρχείο Excel.

---

## Folder Map & Versioning Policy
```
RULES_v1.md                  Ιστορικό v1
RULES_v2.md                  Ιστορικό v2
RULES_v3.md                  Ιστορικό v3
RULES_v4.md                  Ιστορικό v4
RULES_v5.md                  Ιστορικό v5
RULES_v6.md                  Ιστορικό v6
RULES_v7.md                  Τρέχον v7
RULES.md                     Active (αντίγραφο του v7)
AGENTS_v1.md                 Ιστορικό v1
AGENTS_v2.md                 Ιστορικό v2
AGENTS_v3.md                 Ιστορικό v3
AGENTS_v4.md                 Ιστορικό v4
AGENTS_v5.md                 Ιστορικό v5
AGENTS_v6.md                 Ιστορικό v6
AGENTS_v7.md                 Τρέχον v7
AGENTS.md                    Active (αντίγραφο του v7)
```

---

## 4. CAD / Engineering Plan Object Counting Methodology

# CAD / Engineering Plan Object Counting Methodology

When asked to detect and count structural elements in a drawing (piles, tie rods/holes, columns, openings) when NOT explicitly given in a BOQ, drawing text, or study:

1. **NEVER count manually by looking at the image (risk of hallucination). Always write and execute a Python OpenCV script.**
2. **PIPELINE PROTOCOL:**
   - **a. CONTEXT EXTRACTION (ROI):** 
     - Locate the hosting geometry first (e.g. wall contour, perimeter beam, grid axis).
     - Mask out irrelevant regions (text annotations, dimension arrows, hatching patterns) using color thresholding or bounding contours.
   - **b. ELEMENT TEMPLATE SELECTION:**
     - If circular (piles, core drills, holes): Use a synthetic dark disc kernel (`cv2.circle`) or `cv2.HoughCircles` with tuned radius and minimal blur (`ksize=3`).
     - If rectangular (columns, footings): Use `cv2.findContours` on thresholded lines, filtering by aspect ratio and area, OR a rectangular synthetic template.
     - If complex CAD block: Crop one clear sample from the image and apply `cv2.matchTemplate` with rotation invariance (or multi-angle if not symmetrical).
   - **c. NON-MAXIMUM SUPPRESSION (NMS):**
     - Always pass bboxes as `[x, y, w, h]` to `cv2.dnn.NMSBoxes`.
     - Set `score_threshold` and `nms_threshold` based on the spacing between elements.
   - **d. GEOMETRIC & PHYSICAL SANITY CHECKS:**
     - Filter candidates by verifying that the core/center pixel intensity matches the object type.
     - Verify that consecutive distances between detected objects conform to standard engineering spacings (e.g. c/c distance).
   - **e. OUTPUT:**
     - Draw red circles/boxes with consecutive numeric labels (1, 2, 3...) following the natural physical path (chainage).
     - Report the exact deterministic count and save the annotated image.
