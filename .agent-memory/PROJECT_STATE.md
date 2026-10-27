# PROJECT STATE — NIKOLAS TRY (ΘΕΜΕΛΙΟΤΕΧΝΙΚΗ ΛΤΔ)

> **Last Updated:** 2026-10-27
> **Workspace Root:** `C:\Users\nikth\Desktop\NIKOLAS TRY`
> **Sync Bridge:** `.agent-memory/` + Git (`sync_push.bat` / `sync_pull.bat`)

---

## 1. System Architecture & Overview

This workspace (`NIKOLAS TRY`) is the central engineering and automated geotechnical costing environment for **ΘΕΜΕΛΙΟΤΕΧΝΙΚΗ ΛΤΔ (Themeliotechniki Ltd)**. It houses the production pricing rules, deterministic Python calculation/Excel engines, company price catalogs, bending schedules, and individual project offers across three core geotechnical domains:

1. **Πάσσαλοι (Bored Piles)**
2. **Αγκυρώσεις & Berlin Wall (Anchors, Tiebacks & Shoring)**
3. **Εδαφοηλώσεις (Soil Nailing & Shotcrete)**

### Directory Structure

```text
NIKOLAS TRY/
├── .agent-memory/                 # Cross-PC agent memory & context sync bridge
│   ├── PROJECT_STATE.md           # Current architecture, active versions, and key decisions
│   ├── TASKS.md                   # Active checklist of completed and pending tasks
│   ├── IMPLEMENTATION_PLAN.md     # Current technical & automation implementation plan
│   └── WALKTHROUGH.md             # Latest session walkthrough & verification log
├── .antigravityrules              # Workspace root agent rules (auto-loaded)
├── .gemini/
│   └── GEMINI.md                  # Workspace root Gemini rules (auto-loaded)
├── GEMINI.md                      # Root hierarchical rule file
├── sync_push.bat                  # One-click script: Brain -> .agent-memory + git commit & push
├── sync_pull.bat                  # One-click script: git pull + status check on second PC
├── Bending Schedule/              # Bar bending schedule examples & templates
├── Projects/                      # Individual client project folders (strictly isolated)
│   ├── Agkirosis/                 # Athenion, BOLTON_TOWER, Emeralda, TheHub, TOFARCO_ATHENION
│   ├── Edafoilosis/               # ALDO_SQ1, ST788_Lexum
│   └── Piles/                     # Atlantic_Parekklishia, BOLTON_TOWER, C+ Studio Khaldoon and Omar, TheHub, Tryout
└── Rules/                         # Official versioned pricing rules & engines
    ├── GEMINI.md                  # Core pricing governance & isolation protocol
    ├── Master_Framework/          # Cross-domain pricing & Excel formatting standards
    ├── Κανονες_Αγκυρωσεων/        # Anchors & Berlin Wall rules
    ├── Κανονες_Εδαφοηλωσεων/      # Soil Nailing rules
    └── Κανονες_Πασσαλων/          # Bored Piles rules, Python engine, and Excel catalog
```

---

## 2. Official Active Production Versions (`ACTIVE_VERSION.json`)

Before any costing task, the agent **must** verify `ACTIVE_VERSION.json` in the corresponding `Rules/` subfolder. As of **2026-10-27**, the official active production versions are:

| Domain | Folder | Active Version | Active Files |
| :--- | :--- | :---: | :--- |
| **Master Framework** | `Rules/Master_Framework/` | **`v2`** | `MASTER_RULES_PRICING_FRAMEWORK_v2.md`, `MASTER_AGENTS_PRICING_FRAMEWORK_v2.md` |
| **Πάσσαλοι (Bored Piles)** | `Rules/Κανονες_Πασσαλων/` | **`v7`** | `RULES_v7.md`, `AGENTS_v7.md`, `pile_engine_v7.py`, `ΤΙΜΟΚΑΤΑΛΟΓΟΣ_ΕΤΑΙΡΕΙΑΣ_ΠΑΣΣΑΛΩΝ_v4.xlsx` |
| **Αγκυρώσεις & Berlin Wall** | `Rules/Κανονες_Αγκυρωσεων/` | **`v8`** | `RULES_v8.md`, `AGENTS_v8.md` *(Note: `Rules/GEMINI.md` still lists v5 in text, but `ACTIVE_VERSION.json` is authoritative at `v8`)* |
| **Εδαφοηλώσεις (Soil Nailing)** | `Rules/Κανονες_Εδαφοηλωσεων/` | **`v8`** | `RULES_v8.md`, `AGENTS_v8.md` |

---

## 3. Core Architectural & Governance Decisions

1. **Mandatory `ACTIVE_VERSION.json` Check:**
   - Every new session performing a cost estimation must read `ACTIVE_VERSION.json` in the target domain folder first.
   - Loading or referencing anything inside `older_versions/` is strictly forbidden.
2. **Copy-on-Change (No In-Place Rule Edits):**
   - Production rules, engines (`pile_engine_v7.py`), and catalogs are never modified in-place. Any rule update requires creating the next version (`vN+1`), testing it, archiving the previous version into `older_versions/`, and updating `ACTIVE_VERSION.json`.
3. **Zero Cross-Project Contamination (Blind Test Protocol):**
   - When pricing a project in `Projects/<Domain>/<ProjectName>`, the agent must **never** read or copy data from other folders under `Projects/`. Only the official `Rules/` files and the current project's input drawings/documents may be used.
4. **3-Sheet Deterministic Pile Engine (`pile_engine_v7.py`):**
   - For Bored Piles (`v7`), the agent does not write ad-hoc Excel scripts. It extracts quantities from the engineering drawings into `project_spec.json` and runs `pile_engine_v7.py`.
   - Key `v7` features:
     - **3-Sheet Architecture:** Sheet 1 (Οικονομική Προσφορά), Sheet 2 (Αναλυτική Κοστολόγηση & Πράξεις including Section 2B All-in-One Summary and Section 2C Per-Pile Breakdown), Sheet 3 (`Τιμές & Έλεγχος`).
     - **Police Escort (`Συνοδεία Αστυνομίας`):** `8.00 €/km` always calculated from Headquarters (`Σπύρου Κυπριανού 32, Άγιος Συλάς`) to the project site (both inside and outside Limassol).
     - **Interactive YES/NO Toggles on Sheet 3:** Pile Breaker (`Κεφαλοκοπή` @ `150.00 €/τεμ.`) and Excavated Soil Removal (`Απομάκρυνση Εκσκαπτόμενου Χώματος & Λατομικά Τέλη` @ `15.00 €/m³` provisional).
     - **Body Spiral Formula:** `N_turns = ROUNDUP(L_pile / s, 0) + 1` below Cut-Off; Starter links above Cut-Off do not add `+1`.
5. **Multi-PC Sync Bridge (`.agent-memory/` + Git):**
   - Global `~/.gemini` history is intentionally **not** synced across PCs so unrelated personal/other-project chats remain separate.
   - All project state, tasks, and implementation plans live inside `.agent-memory/` in this repository and are synced via `sync_push.bat` and `sync_pull.bat`.
