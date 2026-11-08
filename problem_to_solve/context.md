# Context: Finnish satellite-interference screening hackathon

## Purpose and handoff

Build an AI-assisted, transparent workflow that identifies newly published satellite filings that may affect Finnish spectrum use and prepares evidence for expert review.

The user will provide data links separately to the implementing AI agent. Do not assume those links or datasets have already been supplied, accessed, or validated. Begin by inventorying the supplied sources and their schemas, access requirements, dates, versions, units, and coverage.

This document captures the problem, available station-table headers, engineering approach, and suggested demonstrator scope. It is not a regulatory determination or a complete propagation specification.

## Problem statement

Spectrum use requires planning and compliance with the ITU Radio Regulations and national rules. Satellites can create cross-border interference concerns because their coverage is not confined to the notifying country.

Experts inspect ITU Radiocommunication Bureau International Frequency Information Circular (BR IFIC) publications, compare satellite plans with Finnish frequency allocations and assignments, assess technical concerns, and prepare correspondence to the notifying administration. The aim is to resolve potential issues before deployment and preserve applicable protection rights.

The hackathon should reduce manual effort, improve coverage, and keep every finding technically reviewable. Start with space services; keep the architecture extensible to terrestrial planning.

Relevant publication categories include API/A, CR/C, and Appendix 30, 30A, and 30B material. Their procedures differ: do not treat them as interchangeable or invent a universal objection deadline. BR IFIC is published every two weeks; the problem statement's phrase “bi-monthly” is ambiguous and should not be interpreted as every two months.

National inputs may include the Finnish frequency allocation table, regulations referred to in the challenge as M4 and M15, and additional frequencies of special interest. Resolve the exact documents, editions, applicability, and roles from the links supplied later. Do not assume that every document defines a technical protection threshold.

## Core mental model

Interference screening connects a transmitter, a propagation path, and a receiver:

1. Spectral interaction: overlapping bands, adjacent-band emissions, or other relevant unwanted emissions.
2. Geographic coupling: visibility, beam coverage, distance, and antenna directions.
3. Signal significance: arriving power, receiving-antenna coupling, and an applicable protection criterion.

Frequency overlap is a candidate alert, not proof of harmful interference. A regulatory coordination trigger, a hard limit exceedance, and actual receiver degradation are distinct findings.

Keep directions separate:

- Satellite downlink transmitter → Finnish terrestrial receiver: screen satellite TX against national RX.
- Finnish terrestrial transmitter → satellite uplink receiver: screen national TX against satellite RX.

Recommended first scope: satellite downlink → Finnish receiver screening. State explicitly which services and bands the demonstrator supports.

## Station-table headers supplied by the user

| English header | Finnish header |
|---|---|
| Station number | Asemanro |
| Class of licence | Lupalaji |
| Longitude | Longitudi |
| Latitude | Latitudi |
| Ground height (m) | Maankork (m) |
| Mast height (m) | Mastonkork (m) |
| Mast type | Mastotyyppi |
| Channel bandwidth | Kan.lev. |
| TX frequency | TX-taajuus |
| RX frequency | RX-taajuus |
| Ant height AGL (m) | Ant.kork.maa (m) |
| Ant height ASL (m) | Ant.kork.meri (m) |
| Ant gain (dB) | Ant.vahv (dB) |
| Class of transmission | Läheteluokka |
| Transmission bandwidth | Lähetteen leveys |
| Radiation direction (°) | Max.sät.suunta (°) |
| Polarisation | Polarisaatio |
| Radiation power | Sät.teho |
| Radiation power (ITU) | Sät.teho (ITU) |
| Radiation power unit | Sät.teh.yks. |
| Transmission line (dB) | Siirtolinja (dB) |
| Directivity | Suunnattavuus |

The original English text contains the typo “Radioation”; retain source column names during ingestion and map them to canonical names.

### Mandatory schema checks

- Determine whether a row represents a transmitter, receiver, paired-link endpoint, or combined licence record.
- Verify coordinate reference system, coordinate format, frequency and bandwidth units, and missing-value conventions.
- Determine which bandwidth describes the receiver passband versus the transmitted emission.
- Establish whether gain, direction, polarisation, and feeder loss apply to TX, RX, or both.
- Determine whether antenna gain is referenced to isotropic gain (dBi), dipole gain (dBd), or another convention.
- Establish whether power means transmitter output, ERP, EIRP, total power, or spectral density. Decode the ITU-formatted field using its actual specification.
- Avoid double-counting antenna gain or subtracting feeder loss twice.
- Check antenna ASL ≈ ground height + antenna AGL, allowing for reference and rounding differences. Mast height is not automatically antenna height.
- Establish azimuth convention and distinguish horizontal direction from a complete 3D antenna pattern.
- Preserve original values, units, source-row identifiers, and conversion records.

## Calculations, in priority order

### 1. Frequency overlap

For a genuinely centre-frequency-based record with bandwidth B:

    lower = centre_frequency - B / 2
    upper = centre_frequency + B / 2

Prefer explicit lower/upper limits when supplied. Do not silently apply this model to records with other frequency semantics.

    overlap_hz = max(0, min(satellite_upper, receiver_upper)
                        - max(satellite_lower, receiver_lower))

Report overlap width and separation. Adjacent-band screening requires emission-mask and receiver-selectivity information for quantitative conclusions. Any heuristic guard band must be labelled as a configurable screening assumption, not a regulatory limit.

### 2. Visibility, slant range, and antenna geometry

Given a satellite position and a terrestrial location, calculate elevation, azimuth, and slant range. Check beam coverage and use antenna gain toward the counterpart, not merely maximum antenna gain.

For non-geostationary systems, geometry varies over time. A single snapshot does not establish the worst case. Document the time sampling, orbital assumptions, and coverage model. If using a coarse Finnish geographic grid, label its spatial resolution and avoid claiming complete coverage of local maxima.

### 3. Power-flux density (PFD)

A simplified free-space expression is:

    PFD_dBW_per_m2 = EIRP_toward_location_dBW - 10*log10(4*pi*R_m**2)

R is slant range in metres. EIRP must be the value toward the evaluated point, not automatically peak beam EIRP. This expression assumes the relevant far-field, free-space model; additional propagation losses require an explicit model.

PFD derived from total EIRP is integrated over the associated signal bandwidth. Match the regulatory reference bandwidth before comparing values.

For a justified uniform-power spectral model, with a reference bandwidth within the occupied bandwidth:

    PFD_reference = PFD_signal + 10*log10(B_reference / B_signal)

Do not apply this simplification blindly to nonuniform spectra, narrow emissions, or reference bandwidths wider than the signal. Prefer filed maximum power-density data and applicable ITU calculation methods.

### 4. Regulatory margin

    margin_dB = calculated_PFD - applicable_PFD_limit

A positive result exceeds the selected threshold, provided units, reference bandwidth, geometry, and applicability match.

Store the exact rule, edition, band, service, elevation dependence, reference bandwidth, and procedural meaning. Distinguish hard limits from coordination triggers. Not every band or service is governed by the same PFD test. Missing applicability information means “assessment incomplete,” not “compliant.”

### 5. Receiver-specific interference: optional second stage

For a suitable plane-wave model:

    A_eff_m2 = G_receive_linear * wavelength_m**2 / (4*pi)
    I_dBW = PFD_dBW_per_m2 + 10*log10(A_eff_m2) - coupling_losses_dB

Use receive gain toward the satellite, compatible bandwidths, and explicitly defined polarisation and feeder losses. Apply spectral integration/filtering when required. Avoid counting losses already included elsewhere.

If receiver noise information is available:

    N_dBW = 10*log10(k_B * T_system_K * B_noise_Hz)
    I_over_N_dB = I_dBW - N_dBW

Use the applicable service-specific criterion; do not invent a universal I/N threshold or receiver sensitivity. Receiver noise and antenna parameters may be absent from the station table.

### 6. Aggregation and EPFD: later scope

Simultaneous interference powers add in linear units, not directly in dB. Non-GSO aggregate assessments can require time-dependent constellation analysis. Equivalent power-flux density (EPFD) includes receiving-antenna weighting for specific satellite protection contexts; it is not a universal replacement for terrestrial PFD screening.

## Missing satellite and receiver inputs

Depending on the selected test, obtain:

- Satellite/network identifier and notifying administration.
- Filing type, publication issue/date, and revision history.
- Frequency ranges, emission bandwidths, and transmission direction.
- Orbital characteristics or evaluated satellite positions.
- Beam coverage, directional gain, and power or power-density information.
- Applicable regulatory criteria and protection status.
- Receiver patterns, bandwidth, noise/system temperature, selectivity, and protection criteria for receiver-level analysis.

Never fabricate missing engineering data. Separate “no candidate found,” “screened below criterion,” and “insufficient information.”

## Suggested demonstrator architecture

    Supplied national data + newly published ITU material
      → source inventory and versioning
      → structured extraction with provenance
      → normalization and validation
      → satellite-TX / Finnish-RX frequency screening
      → geographic and visibility screening
      → PFD calculation where supported
      → applicable-rule comparison
      → evidence-backed expert review queue
      → correspondence draft for human approval

Use deterministic code for parsing validated numeric fields, conversions, interval matching, geometry, calculations, and comparisons. Use AI for document extraction with verification, rule retrieval, summaries, explanations, and draft correspondence. Do not let an LLM invent limits, deadlines, orbital parameters, or final regulatory conclusions.

ITU SpaceCom may assist comment-management workflows. ITU GIBC supports satellite-network technical examinations, including relevant PFD calculations. Confirm tool capabilities, supported formats, versions, and access before relying on automation. Direct calculations must document assumptions and should be benchmarked against authoritative methods where possible.

## Alert/evidence record

Each alert should contain:

- Network/satellite, filing identifier, and notifying administration.
- Publication category, issue/date, source locator, and verified procedural deadline if applicable.
- Affected Finnish station IDs or protected frequency ranges.
- Exact overlap, frequency units, and transmission direction.
- Evaluation location, satellite geometry/time, and coverage assumptions.
- Power input, directional gain, slant range, and reference bandwidth.
- Computed PFD, applicable threshold, and margin, when supported.
- Rule citation, applicability explanation, and distinction between trigger and hard limit.
- Missing inputs, uncertainty, model version, and calculation trace.
- Suggested expert action and correspondence draft, never automatically sent.

## Validation and demo success

- Test unit conversions, touching/non-overlapping intervals, bandwidth normalization, and dB/linear arithmetic.
- Check known geometry and hand-calculated free-space cases.
- Test missing fields and ambiguous power definitions explicitly.
- Separate real data from clearly labelled synthetic fixtures.
- Include a below-threshold case, a flagged case, and an insufficient-data case.
- Do not claim precision or regulatory validity beyond the model and inputs.
- Prefer a small, reproducible end-to-end workflow over an unsupported global interference prediction.

A successful output explains: “This filing overlaps these Finnish receivers; under these inputs and geometry the calculated value has this margin against this applicable criterion; these uncertainties remain; an expert should review this case.”

## How prevention follows from the analysis

The workflow supports early coordination and targeted changes such as frequency changes, power/power-density reductions, beam changes, or operating constraints. It does not itself prevent interference, confer protection rights, or replace expert regulatory correspondence.

## Starting reference links from the preceding discussion

These are orientation references, not substitutes for the data links the user will supply. Revalidate current editions and applicability before implementation.

- ITU GIBC: https://www.itu.int/en/ITU-R/software/pages/gibc.aspx
- ITU BR IFIC overview: https://www.itu.int/en/ITU-R/space/Documents/BR%20IFIC%20(Space%20services).pdf
- ITU PFD hard/trigger limits and reference bandwidth presentation: https://www.itu.int/en/ITU-R/space/WRS18space/WRS18_CRC_PFD_Hard_Trigger_Limits_and_Averaging_Bandwidth_06122018.pdf
- ITU BR IFIC / SpaceCom background: https://www.itu.int/en/ITU-R/space/workshopBangkok2010/11%20BRIFIC_CCL.pdf
- ITU EPFD background: https://www.itu.int/en/ITU-R/space/WRS16space/WRS-16_EPFD.pdf
- ITU Radio Regulations publication page: https://www.itu.int/pub/R-REG-RR
