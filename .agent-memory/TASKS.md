# TASKS — NIKOLAS TRY (ΘΕΜΕΛΙΟΤΕΧΝΙΚΗ ΛΤΔ)

> **Last Updated:** 2026-10-27
> **Rule:** Update this file automatically whenever a task is completed, added, or reprioritized, and before ending any session.

---

## Completed Tasks

- [x] **Multi-PC Workspace Sync Bridge Setup (2026-10-27)**
  - [x] Create `.agent-memory/` directory with `PROJECT_STATE.md`, `TASKS.md`, `IMPLEMENTATION_PLAN.md`, and `WALKTHROUGH.md`.
  - [x] Configure workspace rules in `.antigravityrules`, `.gemini/GEMINI.md`, `GEMINI.md`, and `.agents/rules/agent-memory-sync.md`.
  - [x] Create `sync_push.bat` and `sync_pull.bat` helper scripts (with brain artifact sync + Git commit/push/pull).
  - [x] Initialize local Git repository with `.gitignore` configured for Excel lock files (`~$*`) and Python cache (`__pycache__/`).
- [x] **Core Rules & Engine Versioning**
  - [x] Establish `Master_Framework` `v2` (`MASTER_RULES_PRICING_FRAMEWORK_v2.md` & `MASTER_AGENTS_PRICING_FRAMEWORK_v2.md`).
  - [x] Upgrade Bored Piles engine to `v7` (`pile_engine_v7.py`, `RULES_v7.md`, `AGENTS_v7.md`, `ΤΙΜΟΚΑΤΑΛΟΓΟΣ_ΕΤΑΙΡΕΙΑΣ_ΠΑΣΣΑΛΩΝ_v4.xlsx`) with 3-Sheet architecture, Police Escort (`8.00 €/km` from Agios Sylas HQ), interactive YES/NO toggles, and Section 2B/2C breakdowns.
  - [x] Upgrade Anchors & Berlin Wall rules to `v8` (`Κανονες_Αγκυρωσεων/RULES_v8.md`).
  - [x] Upgrade Soil Nailing rules to `v8` (`Κανονες_Εδαφοηλωσεων/RULES_v8.md`).
- [x] **Recent Project Cost Estimations (`Projects/`)**
  - [x] `Projects/Agkirosis/Emeralda` — Anchors & shoring cost offer (`S1480_WALLS_DD_01-S01/S02.pdf`).
  - [x] `Projects/Piles/Atlantic_Parekklishia` — Bored piles costing (`PILING DRAWINGS - PLOT 1.pdf`).
  - [x] `Projects/Piles/BOLTON_TOWER` & `Projects/Agkirosis/BOLTON_TOWER` — Piles & Anchors costing.
  - [x] `Projects/Piles/TheHub` & `Projects/Agkirosis/TheHub` — Perimeter verification & costing.
  - [x] `Projects/Agkirosis/TOFARCO_ATHENION` — Shoring & anchors report/costing.

---

## Active / Pending Tasks

### Priority 1: Multi-PC Git Remote Connection
- [ ] Create a private GitHub repository (e.g., `nikolas-try-workspace`) and link it on PC 1 (`git remote add origin <URL>` + `sync_push.bat`).
- [ ] Clone the repository on PC 2 and verify seamless pickup in Antigravity (`sync_pull.bat`).

### Priority 2: Bored Piles (`Κανονες_Πασσαλων`) Open TODOs (from `ACTIVE_VERSION.json`)
- [ ] **TODO 1 (Drilling Cost Breakdown):** Add detailed analytical cost breakdown slide/section per pile hole / drilling (`Κόστος ανά Τρύπα / Διάτρηση Πασσάλων`) in the company catalog.
- [ ] **TODO 3 (Other Catalog Breakdowns):** Complete remaining analytical breakdowns in the catalog (`Χάλυβας`, `Κεφαλοκοπή`, `Εκτοξευόμενο`, etc.).
- [ ] **TODO 4 (Excavated Soil Removal Formula):** Replace provisional `15.00 €/m³` unit rate for `Απομάκρυνση Εκσκαπτόμενου Χώματος & Λατομικά Τέλη` with the detailed internal cost formula once finalized.
- [ ] **Minor Sync Check:** Update line 9 of `Rules/GEMINI.md` so the text summary lists `v8` (instead of `v5`) for `Αγκυρώσεις & Berlin Wall`, matching `Rules/Κανονες_Αγκυρωσεων/ACTIVE_VERSION.json`.

### Priority 3: 5-Stage End-to-End Automation Roadmap (Master Plan)
- [ ] **Stage 1 (Outlook Quotation Scanner & Auto-Costing):** Connect Outlook email scanner to automatically detect incoming RFQs for Piles, Anchors, and Soil Nailing, classify Case A (no study/general quote) vs. Case B (engineering study PDF/DWG), run the corresponding pricing rules/engine, flag out-of-scope BOQ items, and notify for review.
- [ ] **Stage 2 (Automated Concrete Supplier RFQs by District):** Auto-draft/send concrete price requests (`C30/37`, m³, district suppliers in Limassol/Larnaca/Nicosia/Paphos) when pricing Case B studies.
- [ ] **Stage 3 (Concrete Quote Parsing & `_v2.xlsx` Auto-Update):** Parse supplier replies and automatically generate updated `_v2.xlsx` costings with the best concrete price.
- [ ] **Stage 4 (Boss Review & One-Click PDF Offer Export):** Export Sheet 1 (`Οικονομική Προσφορά`) to official PDF and prepare client reply draft.
- [ ] **Stage 5 (Post-Contract Auto-Submittals):** Trigger automated emails upon contract signing for Mix Design, Mill Certificates, Method Statements, and Health & Safety Plan (ΣΑΥ).
