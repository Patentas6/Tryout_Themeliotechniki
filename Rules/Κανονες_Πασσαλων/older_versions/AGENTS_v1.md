# AGENTS_v1.md — Ροή Εργασίας Pricing Πασσάλων & Πασσαλοστοιχιών (Bored Piles)

> **Έκδοση: v1.0 (2026-10) — Επίσημο Πρότυπο Εταιρείας για Πασσάλους**  
> Αυστηρή τήρηση της αρχιτεκτονικής 3 Φύλλων, διατήρηση των 4 Πυλώνων, απόλυτη συμμόρφωση με τους κανόνες v1:  
> 1. Απαγόρευση OneDrive (Zero-Touch Policy).  
> 2. Επίσημος φάκελος έργων `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Piles\<PROJECT_NAME>\` & Versioning (`_v1`, `_v2`, `_v3`).  
> 3. Single Source of Truth & 100% Δυναμική Σύνδεση Excel.  
> 4. Υποχρεωτική αναγραφή πηγής μελέτης (Citations / Σχέδια) στη Στήλη D για κάθε αριθμό.  
> 5. Αυτόματη καταμέτρηση με Python OpenCV όταν δεν δίνεται ρητά το πλήθος πασσάλων.

Read this whole file and `RULES.md` (ή `RULES_v1.md`) before doing anything.

---

## 0. IMMUTABILITY, REPOSITORY, ACTIVE VERSION & FOLDERS POLICY

* **Κεντρικό Αποθετήριο Κανόνων:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Rules\Κανονες_Πασσαλων` (Αυστηρά READ-ONLY).
* **Ανάγνωση Ενεργής Έκδοσης (Active Version Discovery):** Κάθε AI agent που ξεκινάει νέο chat **ΟΦΕΙΛΕΙ ΠΡΩΤΑ ΝΑ ΔΙΑΒΑΖΕΙ ΤΟ `ACTIVE_VERSION.json`** στον φάκελο `Rules\Κανονες_Πασσαλων`. Το αρχείο αυτό ορίζει ποια έκδοση είναι σε παραγωγή (π.χ. `"active_version": "v1"`), ώστε το AI να εφαρμόζει το `RULES_v1.md` / `AGENTS_v1.md`.
* **Κανόνας Απαγόρευσης Τροποποίησης Ενεργής Έκδοσης (Copy-on-Change):** ΠΟΤΕ δεν τροποποιείται η τρέχουσα ενεργή έκδοση in-place. Για οποιαδήποτε αλλαγή ή δοκιμή, αντιγράφουμε το αρχείο σε νέο αριθμό έκδοσης (π.χ. `RULES_v2.md`, `AGENTS_v2.md`) και δουλεύουμε αποκλειστικά εκεί.
* **Προώθηση σε Παραγωγή (Promotion):** Μόνο αφού ελεγχθεί και εγκριθεί η νέα έκδοση από τον χρήστη, ενημερώνεται το `ACTIVE_VERSION.json` σε `"active_version": "v2"` και συγχρονίζονται τα `RULES.md` / `AGENTS.md`.
* **Επίσημος Φάκελος Έργων:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Piles\<PROJECT_NAME>\` (Αποθήκευση όλων των Excel έργου ανά project με εκδόσεις `_v1.xlsx`, `_v2.xlsx` κ.ο.κ.).
* **Κεντρικός Τιμοκατάλογος:** `ΤΙΜΟΚΑΤΑΛΟΓΟΣ_ΕΤΑΙΡΕΙΑΣ_ΠΑΣΣΑΛΩΝ.xlsx`.
* **ΑΠΑΓΟΡΕΥΣΗ ONEDRIVE:** Το AI δεν διαβάζει ούτε γράφει ποτέ σε διαδρομές του OneDrive.

---

## Your role (and its limits)

| You DO | You NEVER |
|---|---|
| **Ελέγχεις ΠΡΩΤΑ το `ACTIVE_VERSION.json`:** Διαβάζεις ποια έκδοση είναι σε παραγωγή και ακολουθείς πιστά τα αντίστοιχα αρχεία | Assume or guess rule versions without checking `ACTIVE_VERSION.json` |
| **Απαγόρευση In-Place Edit:** Αν ζητηθούν αλλαγές κανόνων, αντιγράφεις σε νέα έκδοση (`v2`) και κάνεις εκεί τις αλλαγές | Overwrite or alter the active version files directly |
| **ΠΑΝΤΑ δημιουργείς Excel με 3 Φύλλα:** Φύλλο 1 = Στοιχεία & Φυσικές Ποσότητες (ΧΩΡΙΣ €), Φύλλο 2 = Υπολογισμοί Ευρώ (€), Φύλλο 3 = Βιβλιοθήκη Οπλισμού | Mix raw inputs with calculated formulas or mix physical units with € on Sheet 1 |
| **Αποθήκευση Έργων:** Στον ειδικό φάκελο `Projects\Piles\<PROJECT_NAME>\` με εκδόσεις `_v1`, `_v2` | Overwrite or delete past project versions or alter old rules |
| **Single Source of Truth:** Εισαγωγή παραμέτρων ΜΙΑ ΦΟΡΑ στο Κεφάλαιο Α. Εξαρτημένες παράμετροι συνδέονται με φόρμουλες | Hardcode dependent quantities that break when the user alters an input |
| **Δυναμικές Σημειώσεις & Επεξηγήσεις:** Κατασκευή κειμένων υπολογισμών με δυναμικές φόρμουλες Excel (`="=" & B34 & ...`) | Write static hardcoded numbers in notes that don't update |
| **Citations Μελέτης:** Υποχρεωτική αναγραφή σχεδίου/σελίδας στη Στήλη D για κάθε πρωτογενές δεδομένο | Leave numbers without blueprint citations for engineering cross-checking |
| **Τιμές Μονάδας (€):** ΠΑΝΤΑ με **ΚΙΤΡΙΝΟ γέμισμα `#FFF2CC`** (Στήλη D Φύλλου 2) | Leave unit price cells white or unformatted without the mandatory yellow fill |
| **Σκυρόδεμα +10%:** Προσθέτεις υποχρεωτικά +10% απώλεια στα κυβικά ($V_{\text{conc}} = V_{\text{exc}} \times 1.10$) | Calculate concrete without the required +10% overbreak |
| **Χάλυβας Κλωβών B500C:** 1.10 €/kg ενιαία (κάθετες ράβδοι + σούστα), με 50Φ μάτιση αν $L > 12\text{ m}$ | Forget 50Φ lap length or rebar weight calculation |
| **Καθαρό Σύνολο:** Άθροισμα 4 Πυλώνων = Καθαρό Σύνολο Προσφοράς | Add separate overhead/profit rows when unit rates already include them |
| **KPIs:** Υπολογίζεις και εμφανίζεις: €/m διάτρησης, €/πάσσαλο, €/m³ σκυροδέματος | Skip unit cost KPIs requested by the master and structural engineers |

---

## 4-Step Action Sequence for Every Piles Project

### Step 1 — Sheet 1: Extraction & Physical Quantities (NO €)
1. **Κεφάλαιο Α (Raw Inputs & Citations):** Πράσινα κελιά `#E2EFDA`.
   - Πλήθος πασσάλων, διάμετρος $D$ (mm), βάθος $L$ (m), ποιότητα σκυροδέματος, οπλισμός (κάθετος & σούστα).
2. **Κεφάλαιο Β (Υπολογισμοί με Δυναμικές Σημειώσεις):** Γκρι `#F2F2F2` / Πράσινα `#C6EFCE`.
   - Συνολικά τρέχοντα μέτρα διάτρησης, όγκος εκσκαφής, όγκος σκυροδέματος ($+10\%$), βάρος χάλυβα κλωβών (kg).

### Step 2 — Sheet 2: Financial Pricing (4 Pillars & Totals)
- **Πυλώνας 1:** Μετακίνηση Γεωτρύπανου (Έδρα Άγιος Συλάς).
- **Πυλώνας 2:** Διάτρηση (ανά τρέχον μέτρο).
- **Πυλώνας 3:** Σκυρόδεμα (κουγκρί ανά m³ με +10% απώλεια).
- **Πυλώνας 4:** Χάλυβας Κλωβών B500C (ανά kg στα 1,10 €/kg).
- **Τελικά Σύνολα:** Καθαρό Σύνολο + ΦΠΑ 19% + Τελικό Πληρωτέο.
- **KPIs:** €/τρέχον μέτρο, €/πάσσαλο, €/m³ σκυροδέματος.

### Step 3 — Sheet 3: Steel Rebar Database
- Ενσωματωμένος πίνακας βαρών ράβδων B500C (Y6–Y40).

### Step 4 — Project Saving & Report
- Αποθήκευση στο `C:\Users\nikth\Desktop\NIKOLAS TRY\Projects\Piles\<PROJECT_NAME>\`.

---

## 5. CAD / Engineering Plan Object Counting Methodology

# CAD / Engineering Plan Object Counting Methodology

When asked to detect and count structural elements in a drawing (piles, tie rods/holes, columns, openings) when NOT explicitly given in a BOQ, drawing text, or study:

1. **NEVER count manually by looking at the image (risk of hallucination). Always write and execute a Python OpenCV script.**
2. **PIPELINE PROTOCOL:**
   - **a. CONTEXT EXTRACTION (ROI):** Locate hosting geometry, mask out irrelevant text/hatching.
   - **b. ELEMENT TEMPLATE SELECTION:** Circular disc kernel (`cv2.circle`) or `cv2.HoughCircles` for piles.
   - **c. NON-MAXIMUM SUPPRESSION (NMS):** `cv2.dnn.NMSBoxes` with [x, y, w, h].
   - **d. GEOMETRIC & PHYSICAL SANITY CHECKS:** Verify center intensity & c/c spacing.
   - **e. OUTPUT:** Draw red circles with sequential numbers (1, 2, 3...) along chainage and report exact count.
