# Traficom Space & Terrestrial RF Interference Screening Engine
## Technical Architecture, UI Element Dataflow & Customer Question Guide

This document breaks down the **Interactive Screening Engine** and its automated sub-systems into visual process flowcharts, mathematical explanations of UI elements, and a targeted questionnaire for Traficom spectrum managers.

---

## 1. End-to-End Processing Flowchart

The diagram below maps every input field in the UI to its internal physics calculation, regulatory evaluation, and final visual outputs.

```mermaid
flowchart TD
    subgraph UI_Inputs["1. UI Parameter Inputs"]
        subgraph Sat_Params["Space Station Filing (ITU BR IFIC)"]
            A1["Satellite Longitude (°E/°W)"]
            A2["Downlink Center Freq (MHz)"]
            A3["Signal Bandwidth (MHz)"]
            A4["Max Beam EIRP (dBW)"]
        end
        subgraph Station_Params["Finnish National Station (Inventory / Upload)"]
            B1["Station Coordinates (Lat, Lon)"]
            B2["Antenna Height / Elevation (m)"]
            B3["Rx / Tx Assigned Freq (MHz)"]
            B4["Receiver Channel BW (MHz)"]
        end
        subgraph Reg_Limits["Regulatory Standard Selector"]
            C1["Baseline: Reg M4 (-144 dBW/m²/4kHz)"]
            C2["Stricter: Reg M15 (-154 dBW/m²/4kHz)"]
            C3["ITU RR Art 21, Table 21-4 (-126 dBW/m²/MHz)"]
        end
    end

    subgraph Geometry_Engine["2. 3D Spherical Orbital Geometry Engine"]
        D1["Earth Radius + Orbital Radius (GEO = 42,164 km)"]
        D2["Calculate Geocentric Angular Separation (ψ)"]
        D3["Elevation Angle Evaluation (θ)"]
        D4{"Is Satellite Visible?<br/>(θ ≥ 0° above horizon)"}
        D5["Calculate Slant Range Distance (d in meters)"]
        D6["Line of Sight Blocked<br/>(Flagged: BELOW_HORIZON)"]
    end

    subgraph Spectral_Engine["3. Spectral Overlap Engine"]
        E1["Compute Satellite Band Edges: [f_c - BW/2, f_c + BW/2]"]
        E2["Compute Finnish Station Band Edges: [f_0 - BW/2, f_0 + BW/2]"]
        E3["Calculate Overlap: max(0, min(f_high) - max(f_low))"]
        E4{"Frequency Overlap?<br/>(Overlap > 0 MHz)"}
        E5["Out of Band / No Direct Co-channel"]
    end

    subgraph PFD_Engine["4. Power Flux-Density (PFD) Calculation"]
        F1["Free Space Spreading Loss:<br/>Loss = 10 * log10(4 * π * d²)"]
        F2["Total Arriving PFD (dBW/m²):<br/>PFD_total = EIRP - Spreading Loss"]
        F3["Bandwidth Reference Scaling Factor:<br/>10 * log10(Ref_BW / Sat_BW)"]
        F4["Normalized PFD in Reference Bandwidth:<br/>PFD_ref (e.g. per 4 kHz or per 1 MHz)"]
    end

    subgraph Reg_Engine["5. Threshold Comparison & Concern Assessment"]
        G1["Regulatory Delta Margin (dB):<br/>Margin = PFD_ref - Standard_Limit"]
        G2{"Evaluate Compliance Status"}
        G3["CRITICAL EXCEEDANCE<br/>(Margin ≥ 0 dB - Breaches Rule)"]
        G4["POTENTIAL CONCERN / NEAR-LIMIT<br/>(Margin between -6 dB and 0 dB)"]
        G5["COMPLIANT<br/>(Margin < -6 dB)"]
    end

    subgraph UI_Outputs["6. UI Outputs & Decision Support"]
        H1["KPI Cards: Elevation, Slant Range, PFD, Margin Status"]
        H2["Interactive Leaflet Map: Ground Track, Ray to Station, Footprint"]
        H3["Chart.js Spectrum Visualizer: Visual Overlap of Satellite vs Station"]
        H4["Automated Issues Report Table & Export (CSV / JSON)"]
        H5["Drafted Technical Objection Notice (RR Article 9 / ECC T/R 25-08)"]
    end

    A1 & B1 & B2 --> D1 --> D2 --> D3 --> D4
    D4 -- "Yes (θ > 0°)" --> D5
    D4 -- "No (θ ≤ 0°)" --> D6
    
    A2 & A3 & B3 & B4 --> E1 & E2 --> E3 --> E4
    E4 -- "No" --> E5
    
    A4 & D5 --> F1 --> F2
    A3 & F2 --> F3 --> F4
    
    F4 & C1 & C2 & C3 --> G1 --> G2
    G2 --> G3 & G4 & G5

    D3 & D5 & F4 & G1 --> H1
    D3 & B1 & A1 --> H2
    E1 & E2 --> H3
    G3 & G4 --> H4 & H5
```

---

## 2. Element-by-Element Breakdown

| UI Element | Source / Reference | What Happens Internally | Why It Matters to Traficom |
| :--- | :--- | :--- | :--- |
| **Satellite Orbital Longitude** | ITU BR IFIC circulars (e.g. `25.0°E` or `13.0°E`) | Converts satellite orbital position on the equator into 3D ECEF Cartesian coordinates $(X_s, Y_s, Z_s)$ at $R = 42,164\text{ km}$. | Determines whether the satellite is physically above Finland's horizon ($\theta > 0^\circ$). If below, no line-of-sight path exists. |
| **Finnish Station Coordinates & Ground Height** | 874-station Excel dataset (`Asemanro`, `Korkeus`) | Converts WGS84 latitude, longitude, and elevation into 3D station position $(X_t, Y_t, Z_t)$. | Precise ground location dictates elevation angle and geometric distance to the geostationary arc. |
| **Spectral Overlap Calculator** | Downlink frequency & Channel bandwidth vs. Finnish Rx/Tx frequency | Evaluates $[\max(f_{s,\min}, f_{t,\min}), \min(f_{s,\max}, f_{t,\max})]$. If difference $> 0$, an overlap exists. Checks both Rx and Tx per the Traficom rule *"Every TX is received in Finland"*. | Identifies co-channel or overlapping interference where satellite energy enters the receiver passband. |
| **Slant Range ($d$) & Free-Space Spreading** | 3D Euclidean distance vector | Calculates $d = \sqrt{\Delta X^2 + \Delta Y^2 + \Delta Z^2}$ (typically $38,000 - 41,000\text{ km}$ at Finnish latitudes). Spreading attenuation is $10\log_{10}(4\pi d^2) \approx 162.5 - 163.3\text{ dB}$. | Standard inverse-square energy divergence across interplanetary distance. |
| **Power Flux-Density (PFD) Scaling** | Sat EIRP (dBW) & Bandwidth | Arriving $\text{PFD}_{\text{total}} = \text{EIRP} - 10\log_{10}(4\pi d^2)$. Scaled to reference bandwidth ($4\text{ kHz}$ or $1\text{ MHz}$) using $10\log_{10}(\text{BW}_{\text{ref}} / \text{BW}_{\text{sat}})$. | Satellite filings declare peak EIRP across wide carriers (e.g. $20\text{ MHz}$ or $36\text{ MHz}$); limits are normalized per $4\text{ kHz}$ slice. |
| **Regulatory Evaluation Standard** | Dropdown: Regulation M4, M15, ITU RR Art 21 | Margin $\Delta = \text{PFD}_{\text{calc}} - \text{PFD}_{\text{limit}}$. If $\Delta \ge 0\text{ dB}$, flagged **CRITICAL EXCEEDANCE**. If $-6\text{ dB} \le \Delta < 0\text{ dB}$, flagged **POTENTIAL CONCERN / NEAR-LIMIT**. | Automatically highlights filings requiring formal objection within the ITU 4-month regulatory window. |

---

## 3. Sequence Diagram: What Happens When You Click "Run Screening"

```mermaid
sequenceDiagram
    autonumber
    actor User as Spectrum Engineer
    participant UI as Interactive UI Tab
    participant Engine as Deterministic Interference Engine
    participant Reg as Regulatory Rulebase (M4 / M15 / Art 21)
    participant Display as Visualizers (Map, Spectrum, Evidence)

    User->>UI: Select station or enter custom parameters
    User->>UI: Click "Calculate & Run Interference Screening"
    UI->>Engine: Send (SatCoords, StationCoords, Frequencies, EIRP, Bandwidths)
    
    Engine->>Engine: Compute 3D Vectors & Slant Range (d)
    Engine->>Engine: Compute Topocentric Elevation Angle (θ)
    alt Satellite below horizon (θ < 0°)
        Engine-->>UI: Return "BELOW_HORIZON" (Elevation < 0°, No LOS)
    else Satellite visible (θ ≥ 0°)
        Engine->>Engine: Compute Free Space Spreading: 10*log10(4*π*d²)
        Engine->>Engine: Compute Total PFD & Normalized PFD (per 4 kHz)
        Engine->>Engine: Evaluate Spectral Overlap (MHz)
        Engine->>Reg: Compare PFD against Selected Standard (M4: -144 dBW/m²)
        Reg-->>Engine: Return Margin (dB) and Status (EXCEEDED / CONCERN / COMPLIANT)
    end

    Engine-->>UI: Output complete numerical dataset
    UI->>Display: Update KPI Cards (Slant Range, Elevation, PFD, Margin)
    UI->>Display: Plot Spectrum Envelope on Chart.js
    UI->>Display: Draw Orbit-to-Ground Ray & Footprint on Leaflet Map
    UI->>Display: Generate Verifiable JSON Evidence Record
    Display-->>User: Visual feedback & highlighted actionable alerts
```

---

## 4. Key Questions to Ask the Customer (Traficom Specialists)

Use these targeted questions in interviews or customer alignment sessions:

### A. Regulatory Thresholds & Margins
1. **M4 vs. M15 Baseline:** "We have configured **Regulation M4 ($-144\text{ dBW/m}^2/4\text{kHz}$)** as the default baseline, with **Regulation M15 ($-154\text{ dBW/m}^2$)** selectable for sensitive mobile/terrestrial systems. Are there specific frequency ranges where M15 should be enforced automatically rather than by manual selection?"
2. **Concern Threshold Window:** "Our engine flags any filing within **$6\text{ dB}$ of the limit** as a `POTENTIAL CONCERN` even if it does not strictly exceed the hard limit. Does Traficom prioritize bilateral coordination inquiries for near-limit cases within $3\text{ dB}$, $6\text{ dB}$, or a wider margin?"

### B. Antenna Directivity & Terrain Shielding
3. **Worst-Case Peak Envelope vs. Off-Axis Discrimination:** "Per official advice, we evaluate the worst-case peak beam EIRP envelope because boresight coordinates are often omitted in early IFIC filings. When coordination advances, would you want the tool to ingest ITU-R S.465/S.580 earth station antenna patterns to model off-axis rejection?"
4. **Digital Elevation Models (DEM):** "For terrestrial microwave fixed links, does Traficom require terrain profile obstruction checks (SRTM 30m) along the path, or is geometric line-of-sight sufficient for the initial screening stage?"

### C. Terrestrial Cross-Border (ECC T/R 25-08)
5. **Border Coordination Threshold:** "For the land-mobile side quest, we calculate field strength $E$ along 22 points across the Finnish borders. Is the standard **$30\text{ dB}\mu\text{V/m}$** at 3m antenna height appropriate for all bands, or should we incorporate specific thresholds for 700 MHz (Band 28) and 3.5 GHz (Band n78)?"
6. **Sea vs. Land Propagation Weights:** "Across the Gulf of Finland from Tallinn, ducting and low sea clutter maintain high field strengths over 80 km. Do you prefer ITU-R P.1546 50% time / 50% locations or 10% time (worst-case anomalous propagation) for cross-border coordination?"
