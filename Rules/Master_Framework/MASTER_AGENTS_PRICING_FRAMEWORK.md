# MASTER_AGENTS_PRICING_FRAMEWORK.md — Ενιαία Ροή Εργασίας AI Agents για Κοστολόγηση Γεωτεχνικών Έργων

> **Έκδοση: Master Agent Framework v2.0 (2026-10)**  
> **Σκοπός:** Αυστηρή τυποποίηση της διαδικασίας ανάγνωσης σχεδίων, κατασκευής Excel και οικονομικής κοστολόγησης για όλα τα είδη εργασιών της εταιρείας ΘΕΜΕΛΙΟΤΕΧΝΙΚΗ ΛΤΔ.

---

## 0. IMMUTABILITY, REPOSITORY & FOLDERS POLICY

* **Κεντρικοί Φάκελοι Έργων:**
  * Εδαφοηλώσεις: `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Edafoilosis\<PROJECT_NAME>\`
  * Αγκυρώσεις: `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Agkirosis\<PROJECT_NAME>\`
  * Πάσσαλοι: `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Piles\<PROJECT_NAME>\`
* **Κεντρικό Αποθετήριο Κανόνων:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Rules\` (`Κανονες_Εδαφοηλωσεων\`, `Κανονες_Αγκυρωσεων\`, `Κανονες_Πασσαλων\`, `Master_Framework\`)
* **Ανάγνωση Ενεργής Έκδοσης (Active Version Discovery):** Κάθε AI agent που ανοίγει νέο chat **ΟΦΕΙΛΕΙ ΠΡΩΤΑ ΝΑ ΔΙΑΒΑΖΕΙ ΤΟ `ACTIVE_VERSION.json`** στον σχετικό φάκελο κανόνων. Αυτό ορίζει ποια έκδοση κανόνων/agent είναι σε παραγωγή.
* **Κανόνας Απαγόρευσης Τροποποίησης Ενεργής Έκδοσης (Copy-on-Change):** ΠΟΤΕ δεν τροποποιούμε την ενεργή έκδοση in-place. Για οποιαδήποτε αλλαγή ή δοκιμή, αντιγράφουμε σε νέο αρχείο με αυξημένο αριθμό έκδοσης (π.χ. `_v2.md` ή `_v8.md`) και κάνουμε τις αλλαγές αποκλειστικά εκεί.
* **Προώθηση σε Παραγωγή (Promotion):** Μόνο αφού ελεγχθεί και εγκριθεί η νέα έκδοση από τον χρήστη, ενημερώνεται το `ACTIVE_VERSION.json` ώστε όλα τα επόμενα chats να ακολουθούν τη νέα έκδοση.
* **ΑΠΑΓΟΡΕΥΣΗ ONEDRIVE:** Το AI δεν διαβάζει ούτε γράφει ποτέ σε διαδρομές του OneDrive.
* **ΑΠΟΛΥΤΟΣ ΚΑΝΟΝΑΣ ΑΠΟΜΟΝΩΣΗΣ ΕΡΓΟΥ & ZERO CROSS-PROJECT CONTAMINATION:** Το AI απαγορεύεται αυστηρά να διαβάζει ή να ανοίγει άλλα projects ή παλαιότερα αρχεία excel/προσφορών. Περιορίζεται αποκλειστικά στα αρχεία εισόδου της τρέχουσας μελέτης και στους φακέλους Rules\.
* **ΟΡΓΑΝΩΣΗ ΦΑΚΕΛΩΝ OLDER_VERSIONS:** Στον κεντρικό φάκελο κάθε κατηγορίας κανόνων μένουν ΜΟΝΟ τα αρχεία της ενεργής έκδοσης και το ACTIVE_VERSION.json. Όλα τα παλαιότερα μεταφέρονται στο older_versions\.

---

## Your role (and its limits)

| You DO | You NEVER |
|---|---|
| **Ελέγχεις ΠΡΩΤΑ το `ACTIVE_VERSION.json`:** Διαβάζεις ποια έκδοση είναι σε παραγωγή και ακολουθείς πιστά τα αντίστοιχα αρχεία | Assume or guess rule versions without checking `ACTIVE_VERSION.json` |
| **Απαγόρευση In-Place Edit:** Αν ζητηθούν αλλαγές κανόνων, αντιγράφεις σε νέα έκδοση (`v2`, `v8` κτλ.) και κάνεις εκεί τις αλλαγές | Overwrite or alter the active version files directly |
| **ΠΑΝΤΑ δημιουργείς Excel με 3 Φύλλα:** Φύλλο 1 = Στοιχεία Μελέτης & Φυσικές Ποσότητες (ΧΩΡΙΣ €), Φύλλο 2 = Υπολογισμοί Ευρώ (€), Φύλλο 3 = Τεχνική Βιβλιοθήκη | Mix raw inputs with calculated formulas or mix physical units with € on Sheet 1 |
| **Αποθήκευση Έργων:** Στον ειδικό φάκελο `Projects\<PROJECT_NAME>\` με εκδόσεις `_v1`, `_v2`, `_v3`, `_v4` | Overwrite or delete past project versions or alter old rules |
| **Single Source of Truth:** Εισαγωγή παραμέτρων ΜΙΑ ΦΟΡΑ στο Κεφάλαιο Α. Εξαρτημένες παράμετροι συνδέονται με φόρμουλες (π.χ. `=B34`) | Hardcode dependent quantities that break when the user alters an input |
| **100% Δυναμικές Σημειώσεις & Επεξηγήσεις:** Κατασκευή ΟΛΩΝ των σημειώσεων με συναρτήσεις κειμένου Excel (`="=" & B34 & ...`) | Write static hardcoded numbers in notes (e.g. `"98 οπές..."`) that don't update |
| **Citations Μελέτης:** Υποχρεωτική αναγραφή σχεδίου/σελίδας στη Στήλη D για κάθε πρωτογενές δεδομένο | Leave numbers without blueprint citations for engineering cross-checking |
| **Καταγραφή Κανόνων:** Σήμανση `[DEFAULT]`, `[OVERRIDE]`, `[ΣΥΓΚΡΟΥΣΗ]` στη Στήλη E | Apply rule changes or defaults silently without audit trails |
| **Τιμές Μονάδας (€):** ΠΑΝΤΑ με **ΚΙΤΡΙΝΟ γέμισμα `#FFF2CC`** (Στήλη E Φύλλου 2) | Leave unit price cells white or unformatted without the mandatory yellow fill |
| **Ειδικές Τιμές:** Πορτοκαλί γέμισμα `#FFE699` για τιμές υπό διερεύνηση ή σύγκρουση | Hardcode special prices without making them visually identifiable for master review |
| **Καθαρό Σύνολο:** Άθροισμα 4 Πυλώνων = Καθαρό Σύνολο Προσφοράς (χωρίς ξεχωριστά Γενικά Έξοδα & Κέρδος) | Add separate overhead/profit rows when unit rates already include them |
| **KPIs:** Υπολογίζεις και εμφανίζεις: €/m² μετώπου, €/τρέχον μέτρο, €/πάσσαλο ή ήλο, €/αγκύρωση | Skip unit cost KPIs requested by the master and structural engineers |

---

## 4-Step Action Sequence for Any Project

### Step 1 — Sheet 1: Blueprint Extraction & Physical Quantities (NO €)
1. **Κεφάλαιο Α (Raw Inputs, Citations & Rule Auditing):**
   - Στήλη A: Περιγραφή στοιχείου μελέτης.
   - Στήλη B: Τιμή μελέτης (`#E2EFDA`).
   - Στήλη C: Μονάδα μέτρησης.
   - Στήλη D: Πηγή Μελέτης / Σχέδιο (π.χ. `Σχέδιο Σ01.1, Κάτοψη -12.50m`).
   - Στήλη E: Τεχνική Σημείωση / Έλεγχος Κανόνα (`[DEFAULT]`, `[OVERRIDE]`, `[ΣΥΓΚΡΟΥΣΗ]`).
   - Όλες οι εξαρτημένες παράμετροι συνδέονται δυναμικά με τύπο (π.χ. `=B34`).
2. **Κεφάλαιο Β (Calculations with 100% Dynamic Text Notes):**
   - Στήλη A: Περιγραφή υπολογισμού.
   - Στήλη B: Αποτέλεσμα με μαθηματικό τύπο (`=ROUND(...)`, `=SUM(...)`, `=PI()`).
   - Στήλη C: Μονάδα μέτρησης.
   - Στήλη D: Αναλυτικός τύπος και επαλήθευση (`="=" & B34 & " οπές x " & TEXT(B77, "0.00") & " m"`).
   - Στήλη E: Δυναμική σημείωση εφαρμογής κανόνα (`="=" & B34 & " οπές x " & TEXT(B77, "0.00") & "m = " & TEXT(B78, "#,##0.00") & " m"`).

### Step 2 — Sheet 2: Financial Pricing (4 Pillars & Totals)
- **Πυλώνας 1:** Μετακινήσεις Εξοπλισμού & Συνεργείων.
- **Πυλώνας 2:** Σχάρες & Shotcrete.
- **Πυλώνας 3:** Κύριο Δομικό Στοιχείο Αντιστήριξης (H-Beams, Ήλοι, Πάσσαλοι).
- **Πυλώνας 4:** Συστήματα Αγκύρωσης / Ενίσχυσης.
- **Σημειώσεις Στήλης G:** Επίσης 100% δυναμικές φόρμουλες κειμένου.
- **Τελικά Σύνολα:** Καθαρό Σύνολο + ΦΠΑ 19% + Τελικό Πληρωτέο.
- **KPIs:** €/m² μετώπου, €/τρέχον μέτρο, €/πάσσαλο ή ήλο, €/αγκύρωση.

### Step 3 — Sheet 3: Technical Library
- Ενσωματωμένος πλήρης πίνακας τυποποιημένων διατομών ή υλικών που αφορούν το συγκεκριμένο είδος έργου.

### Step 4 — Project Saving & Verification
- Αποθήκευση στο `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\<PROJECT_NAME>\` τόσο ως `<PROJECT>_v*.xlsx` όσο και ως τρέχον master.
- Έλεγχος για απουσία σφαλμάτων `#REF!`, `#VALUE!`, `#NAME?`.
- Παρουσίαση στον χρήστη με αναλυτικό πίνακα 4 Πυλώνων, KPIs και συνδέσμους στα αρχεία.

---

## 5. CAD / Engineering Plan Object Counting Methodology

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
