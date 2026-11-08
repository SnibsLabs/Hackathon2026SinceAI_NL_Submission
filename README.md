# Traficom Space & Terrestrial RF Interference Screening Engine

> **Since AI Hackathon 2026 — Traficom Challenge:** *"Saving Earth from Space Interferers Using Radio Frequencies"*

An automated, zero-hallucination regulatory screening platform designed for the **Finnish Transport and Communications Agency (Traficom)**. It screens bi-weekly **ITU BR IFIC** satellite and cross-border terrestrial frequency filings against Finland's **874 protected national stations** and **561 national frequency sub-bands (`taajuusjakotaulukko.json`)**, generating auditable evidence records and formal coordination/objection letters.

---

## System Architecture & Screening Flow

![Architecture Flowchart](./mermaid-diagram.svg)

Why a **Deterministic Regulatory Agent** instead of a black-box ML model? When filing formal legal objections under **ITU Radio Regulations Article 9 (No. 9.52)** or bilateral coordination requests under **ECC Recommendation T/R 25-08**, every decibel of Power Flux Density ($\text{PFD}$) and field strength ($\text{dB}\mu\text{V/m}$) must be 100% reproducible, auditable, and free of LLM math hallucinations.

---

## What Happens Inside the Program (4 Core Modules)

### 1. Interactive Single-Station Screening Engine (Tab 1)
Evaluates any space-to-Earth satellite downlink against a specific Finnish ground receiver:
- **3D Spherical Orbital Geometry:** Computes exact Earth-to-space elevation angle ($\delta$), azimuth, and slant range distance ($d$) from WGS84 station coordinates to the Geostationary Orbit ($35,786\text{ km}$). Checks line-of-sight horizon visibility ($\delta > 0^\circ$).
- **Spectral Overlap Detection:** Computes exact passband intersection in $\text{MHz}$, enforcing Traficom's domain rule (*"Every TX is received in Finland"*).
- **Power Flux Density ($\text{PFD}$) & Dynamic ITU Art. 21 Mask:** Calculates arriving $\text{PFD}$ ($\text{dBW/m}^2$) scaled to the reference bandwidth ($4\text{ kHz}$ or $1\text{ MHz}$) and compares it against selectable regulatory standards (**Regulation M4 `-144 dBW/m²`**, **Regulation M15 `-154 dBW/m²`**, **ITU Art. 21 `-126 dBW/m²/MHz`**, or the **Dynamic Elevation-Angle $\delta$ Mask** per ITU RR Article 21 Table 21-4).
- **Visual & Verifiable Outputs:** Renders an interactive **Leaflet** geometry map, a **Chart.js** spectral mask overlap visual, a **Taajuusjakotaulukko Band Ownership** card, and a copyable JSON **Evidence Record**.

### 2. National Protected Station Ingestion & Batch Comparison (Tab 2)
Ingests and screens Traficom's national inventory of **874 protected stations** (`.xlsx` / `.csv`):
- **Automated Unit Normalization & Data-Quality Quarantine:** Parses mixed frequency/bandwidth strings (`kHz`, `MHz`, `GHz`) into standardized $\text{MHz}$, and automatically quarantines incomplete Excel rows where both `TX-taajuus` and `RX-taajuus` are blank (`0 MHz`) so they never distort compliance statistics.
- **Taajuusjakotaulukko Stakeholder Lookup:** Cross-references every station against **561 compiled Finnish frequency sub-bands** (`taajuusjakotaulukko.json`) to display the affected **Band Owner** (e.g., *Telia / Elisa / DNA*, *Digita Broadcast*, *Fixed Microwave Links*).
- **Wi-Fi / Bluetooth Low-Severity Policy:** Automatically detects overlaps in license-exempt **WAS/RLAN / Bluetooth / ISM** bands (Traficom Regulation 15 / ITU RR `5.150`) and downgrades their severity to `LOW SEVERITY (Wi-Fi / BT Band)` with an explanatory regulatory note.

### 3. Automated Agent Simulator & Scheduled CRON Poller (Tab 3)
Automates the bi-weekly ITU BR IFIC circular review workflow:
- **Single-Batch or Multi-Satellite Full Sweep:** Screens individual target circulars (e.g., **BR IFIC 3076** `@ 2150 MHz`, **BR IFIC 3074** `@ 618 MHz`) or runs a **Full Sweep across all 7 listed ITU satellite filings** against all 874 Finnish stations and 561 frequency bands.
- **Excel-Ready Reporting & Objection Drafting:** Exports clean multi-column `.xlsx` and `sep=,` `.csv` issue reports and auto-drafts a formal **ITU RR Article 9.52 Technical Objection Letter** (with `CC:` to affected Finnish band owners) for human-in-the-loop sign-off.
- **CRON Job Configurator:** Allows engineers to schedule automated polling every $X$ hours or days with custom frequency scopes and evaluation standards.

### 4. Automated Terrestrial Land-Mobile Simulator — Side Quest (Tab 4)
Assesses cross-border terrestrial interference from neighbouring countries (**Estonia, Russia, Sweden, Norway**):
- **22 Finnish Border Points & 874-Station Exposure Check:** Evaluates arriving electric field strength $E$ ($\text{dB}\mu\text{V/m}$) along land and sea propagation paths using **ITU-R P.1546** ($50\%$ and $10\%$ time) and Free Space models.
- **CEPT / ECC Recommendation T/R 25-08 Triggers:** Applies coordination trigger thresholds across VHF/UHF bands—including Traficom's rule extending the **$>400\text{ MHz}$ T/R 25-08 criteria to frequencies above $470\text{ MHz}$**.
- **Bilateral Coordination Output:** Tracks IFIC Notice IDs and foreign transmitter names, exports dedicated Border Field Strength and Exposed Station `.xlsx`/`.csv` tables, and generates a formal bilateral coordination letter.

---

## Quick Start (Zero-Dependency Static Hosting)

1. Clone this repository or open `gemini-code-1794028691323.html` (or `index.html`) directly in any modern browser.
2. For **GitHub Pages**, keep `index.html`, `mermaid-diagram.svg`, and the data folder (`problem_to_solve/` or `code/` containing `default_stations.js` and `taajuusjakotaulukko_compiled.js`) in the repository root.

---

## Further Documentation
- [UI Interactive Screening Engine Guide](./problem_to_solve/UI_INTERACTIVE_SCREENING_ENGINE_GUIDE.md) — Detailed mathematical formulas, flowchart breakdown, and customer Q&A guide.
- [Traficom Expert Guidance Notes](./problem_to_solve/TRAFICOM_EXPERT_GUIDANCE.md) — Domain rules and regulatory assumptions applied in the engine.

## Built With
HTML5, JavaScript (ES6+), Tailwind CSS, Leaflet.js, OpenStreetMap / CartoDB, Chart.js, SheetJS (xlsx), Lucide Icons, JSON, SVG / Mermaid.js, PowerShell, Google Gemini / Antigravity
