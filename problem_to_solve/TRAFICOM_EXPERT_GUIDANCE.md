# Traficom Expert Guidance & Technical Clarifications (Since AI Hackathon 2026)

*Source: Traficom spectrum management specialists (Pekka Pussinen & Otto Julkunen)*

---

## 1. Definition of Regulatory "Concern"

- **Hard Exceedance:** Any calculated Power-Flux Density (PFD) or field strength exceeding the applicable regulatory limit is definitively a technical concern requiring coordination or objection.
- **Close-by Exceedance / Near-Limit Trigger:** A calculated PFD within a close margin of the limit (e.g. within 3 to 6 dB) or semi-high PFD near the band edge also warrants an alert.
- **Antenna Coupling:** Thresholds effectively tighten if the Finnish terrestrial station uses high-gain antennas oriented toward the satellite's orbital azimuth and elevation.
- **Publication Target:** BR IFIC circulars **3074** and **3076** are the primary evaluation editions containing clearly findable interference cases.

---

## 2. Ingestion & Protection Rule: "Every TX is Received in Finland"

- Satellite downlink (space-to-Earth) is the potential interferer.
- The delivered Excel file (`Traficom_SinceAI_Frequencies_of_interest_v05112026.xlsx`, 874 rows) represents the national assignments to be protected.
- **Both RX and TX frequencies must be screened:**
  - In broadcasting (License classes `BC` and `BT`), TX and RX frequencies are identical, but consumer receivers cover the entire footprint surrounding the transmitter.
  - In fixed point-to-point links, a row listing a TX frequency has a paired station receiver listening on that frequency at the other end of the link.
  - *Traficom Rule:* Every TX frequency in the table corresponds to an active reception footprint on Finnish territory.

---

## 3. Satellite Beam Coverage & Pointing

- The free ITU IFIC database table `s_beam` lacks explicit boresight coordinates and pointing vectors.
- *Traficom Rule:* The engine **must assume the worst case**: peak beam EIRP aimed directly at the evaluated Finnish location. Satellites can re-point, tumble, or steer beams; an upper-bound screening envelope is the standard regulatory approach.

---

## 4. Protection Criteria: ITU Radio Regulations Article 21, Table 21-4

The internationally agreed regulatory baseline is **ITU RR Article 21, Table 21-4** (PFD limits at Earth's surface from space stations).

For geostationary space stations, Table 21-4 defines elevation-dependent limits $(\theta)$ over reference bandwidths ($4\text{ kHz}$ or $1\text{ MHz}$):

$$\text{PFD}_{\text{limit}}(\theta) = 
\begin{cases}
L & \theta \le 5^\circ \\
L + 0.5 \times (\theta - 5^\circ) & 5^\circ < \theta \le 25^\circ \\
L + 10\text{ dB} & \theta > 25^\circ
\end{cases}$$

*Example (3.4 – 4.2 GHz):* $L = -152\text{ dB(W/m}^2\text{) in 4 kHz}$, rising to $-142\text{ dB(W/m}^2\text{)}$ above $25^\circ$ elevation.

---

## 5. Finnish GSO Satellite Filings (Appendix 30 / 30A / 30B)

- Official filings notified by Finland can be retrieved from the ITU Space Explorer:
  `https://www.itu.int/itu-r/space/apps/public/spaceexplorer/networks-explorer`
  *(Filter: Notifying Administration = Finland, Type = Geostationary)*

---

## 6. Terrestrial Cross-Border Coordination Above 470 MHz (CEPT T/R 25-08)

*Source: Traficom specialist (Jaakko)*

- **Context:** CEPT Recommendation T/R 25-08 formally specifies land-mobile coordination up to 470 MHz, whereas ~81% of the land-mobile records in BR IFIC 3076 operate above 470 MHz.
- **Official Rule:** For frequencies above 470 MHz, apply the **same coordination trigger criteria specified in CEPT Recommendation T/R 25-08 for the frequency range above 400 MHz** (e.g. $30\text{ dB}\mu\text{V/m}$ trigger threshold at the border line).

---

## 7. National Frequency Allocation Table (`taajuusjakotaulukko.json`) & Wi-Fi / Bluetooth Band Policy

- **Band Ownership & Stakeholder Communication:** Every screened frequency is cross-referenced against Traficom's Open Data National Frequency Allocation Table (`taajuusjakotaulukko.json`) to identify the primary Finnish service allocation, sub-band usage, and affected stakeholder group (e.g. Public Safety `VIRVE / Suomen Erillisverkot`, Commercial MNO Concessions `Telia / Elisa / DNA / Ålcom`, Broadcasting `Digita / Yle`, Defence `Puolustusvoimat`, Aviation/Maritime `Fintraffic`, Fixed Backhaul, or License-Exempt `Reg. 15`).
- **Wi-Fi / Bluetooth / WAS-RLAN / ISM Severity Downgrade:** If a frequency overlap occurs within a license-exempt Bluetooth / Wi-Fi / WAS-RLAN / ISM band (e.g., 2.4 GHz `2400–2483.5 MHz`, 5 GHz `5150–5875 MHz`, 6 GHz `5945–6425 MHz` LPI/VLP RLAN, or ISM/SRD bands under ITU RR 5.150 & Traficom Regulation 15), the conflict is annotated with an explicit regulatory note and its severity is downgraded to **LOW SEVERITY (Wi-Fi / Bluetooth Band)**.
