# AGENTS_v2.md — Ροή Εργασίας Pricing Αγκυρώσεων & Berlin Wall (Berliner Wall)

> **Έκδοση: v2.0 (2026-10) — Επίσημο Πρότυπο Εταιρείας για Berlin Wall & Αγκυρώσεις**  
> Αυστηρή τήρηση της αρχιτεκτονικής 3 Φύλλων, διατήρηση των 4 Πυλώνων, απόλυτη συμμόρφωση με τους νέους κανόνες v2 και Zero-Touch OneDrive.

Read this whole file and `RULES.md` (ή `RULES_v2.md`) before doing anything.

---

## 0. IMMUTABILITY & ONEDRIVE ZERO-TOUCH POLICY

* **Κεντρικό Αποθετήριο:** `C:\Users\nikth\Desktop\NIKOLAS TRY\Rules_Berlin_Wall` (Αυστηρά READ-ONLY).
* **Κεντρικός Τιμοκατάλογος:** `ΤΙΜΟΚΑΤΑΛΟΓΟΣ_ΕΤΑΙΡΕΙΑΣ_BERLIN_WALL.xlsx` (Master τιμές, βάρη, υπολογιστές, πολιτική).
* **Πρότυπο Έργου:** `Πρότυπο_Κοστολόγησης_Berlin_Wall_Αγκυρώσεων.xlsx` (Master template έργου 3 φύλλων).
* **ΑΠΑΓΟΡΕΥΣΗ ONEDRIVE:** Το AI δεν διαβάζει ούτε γράφει ποτέ σε διαδρομές του OneDrive. Εάν ο χρήστης δώσει path από το OneDrive, το AI του ζητά να αποθηκεύσει το αρχείο σε τοπικό φάκελο εκτός OneDrive.

---

## Your role (and its limits)

| You DO | You NEVER |
|---|---|
| **ΠΑΝΤΑ δημιουργείς Excel με 3 Φύλλα:** Φύλλο 1 = Στοιχεία & Φυσικές Ποσότητες (ΧΩΡΙΣ €), Φύλλο 2 = Υπολογισμοί Ευρώ (€), Φύλλο 3 = Βιβλιοθήκη Διατομών Χάλυβα | Mix raw inputs with calculated formulas or mix physical units with € on Sheet 1 |
| **Τιμές Μονάδας (€):** ΠΑΝΤΑ με **ΚΙΤΡΙΝΟ γέμισμα `#FFF2CC`** (Στήλη D Φύλλου 2) | Leave unit price cells white or unformatted without the mandatory yellow fill |
| **Ελάχιστη Διάμετρος Διάτρησης:** Αν ζητείται $< 800\text{ mm}$, βάζεις **ΠΑΝΤΑ 800 mm** ($0.80\text{ m}$) | Use drilling diameter smaller than 800 mm for soldier pile boreholes |
| **Λάμες Σύνδεσης:** Υπολογίζεις ως $L_{\text{beam}} / s_{\text{plate}}$ ανά πάσσαλο | Multiply by arbitrary number of plates per level |
| **Αντιφάσεις Σκυροδέματος:** Επιλέγεις πάντα το ανώτερο στατικά σκυρόδεμα και βάζεις επεξηγηματική σημείωση | Pick inferior grade or fail to explain discrepancy in notes |
| **Υπολογισμός Πασσάλων/Περιμέτρου:** Μετράς ακριβή διαστήματα ($L = \text{διαστήματα} \times s$) όταν δεν δίνονται μέτρα | Guess or hallucinate dimensions without interval counting |
| **Κουγκρί Πάκτωσης (120 €/m³):** Διακριτό χρώμα `#FFE699` για εύκολη αναγνώριση & αλλαγή | Hardcode concrete price without making it visually identifiable for master review |
| **Καθαρό Σύνολο:** Άθροισμα 4 Πυλώνων = Καθαρό Σύνολο Προσφοράς (χωρίς ξεχωριστά Γενικά Έξοδα & Κέρδος) | Add separate overhead/profit rows when unit rates already include them |
| **KPIs:** Υπολογίζεις και εμφανίζεις: €/m² μετώπου, €/αγκύρωση, €/m² για Σχάρες+Shotcrete, €/οπή για H-Beams | Skip unit cost KPIs requested by the master and structural engineers |
| Keep versioned copies of `RULES_v*.md` and `AGENTS_v*.md` whenever rules change, never altering older versions | Overwrite or delete past version files |

---

## 4-Step Action Sequence for Every Project

### Step 1 — Sheet 1: Extraction & Physical Quantities (NO €)
1. **Κεφάλαιο Α (Raw Inputs):**
   - Έδρα, Τοποθεσία, Απόσταση km, Σειρές αγκυρώσεων.
   - Μήκη τοίχων μέσω διαστημάτων ($L = \text{διαστήματα} \times s$), $H_{\text{exc}}$, $H_{\text{embed}}$.
   - Σχάρα (διάμετροι, βήματα, $t_{\text{nom}}$, απώλεια $+0.05\text{ m}$).
   - H-Beams (πλήθος $N_{\text{holes}}$, διάμετρος $\ge 800\text{ mm}$, διατομή, βάρος $kg/m$, διαστάσεις λαμών, $s_{\text{plate}}$, σκυρόδεμα πάκτωσης).
   - Αγκυρώσεις (σειρές, βάθη, strands/οπή, διάμετρος, μήκος grout).
2. **Κεφάλαιο Β (Calculations):**
   - Στράτες (1 τρυπάνι, σειρές αγκυρώσεων, σειρές shotcrete $+ 1$).
   - Περίμετρος ($L = \sum L_i$), Εμβαδόν ($A = L \times H_{\text{exc}}$), Βάρος σχάρας ($kg$), Όγκος shotcrete ($m^3$).
   - H-Beams: Μέτρα διάτρησης, κυβικά εκσκαφής (με $D \ge 0.80\text{ m}$), βάρος νέφκων, πλήθος λαμών ($L/s$), βάρος λαμών, κυβικά πάκτωσης (με $D \ge 0.80\text{ m}$).
   - Αγκυρώσεις: Πλήθος ($N_{\text{holes}} \times \text{σειρές}$), μέτρα διάτρησης, κυβικά grout, μήκος τενόντων, σετ προέντασης & κεφαλών.

### Step 2 — Sheet 2: Financial Pricing (4 Pillars & Totals)
- **Πυλώνας 1:** Μετακινήσεις (3 είδη).
- **Πυλώνας 2:** Σχάρες & Shotcrete.
- **Πυλώνας 3:** H-Beams.
- **Πυλώνας 4:** Αγκυρώσεις.
- **Τελικά Σύνολα:** Καθαρό Σύνολο + ΦΠΑ 19% + Τελικό Πληρωτέο.
- **KPIs:** €/m² μετώπου, €/αγκύρωση, €/m² για Σχάρες+Shotcrete, €/οπή για H-Beams.

### Step 3 — Sheet 3: Steel Profiles Database
- Ενσωματωμένος πλήρης πίνακας διατομών HEA, HEB, UPN και Γωνιών (L).

### Step 4 — Report to User
- Καθαρή παρουσίαση στον χρήστη με αναλυτικό πίνακα 4 Πυλώνων, KPIs και συνδέσμους στα αρχεία.
