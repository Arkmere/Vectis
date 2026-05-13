# Vectis Implementation Plan / Continuity Reference

**Project:** Vectis  
**Repository:** `Arkmere/Vectis`  
**Purpose:** Root-folder reference document to prevent project drift and preserve agreed implementation direction.  
**Current status:** Operational Alpha / Analyst Workbook v2 in active refinement.  

---

## 1. Current Baseline

Vectis is a desktop-local aviation intelligence triage and enrichment utility. It processes Eurocontrol NM/NOP movement exports and produces deterministic, auditable Excel workbooks for analyst review and VKB improvement.

Vectis is not intended to be:

- a live radar frontend;
- a web application;
- a general aviation tracker;
- an opaque AI classifier.

Vectis is intended to be:

- deterministic;
- auditable;
- analyst-oriented;
- enrichment-driven;
- VKB-integrated;
- intelligence-focused.

Current known repository state after recent work:

```text
Repository: Arkmere/Vectis
Branch: main
Recent commits:
f05cdda Update VKB callsigns and replace partial locations dataset
d82baca Use full FDMS location reference dataset
1712b99 Merge Analyst Workbook v2 / GUI refresh
```

Current confirmed capabilities:

```text
- Reads canonical Eurocontrol NM/NOP CSV/XLSX exports.
- Enforces strict NM/NOP schema.
- Uses real VKB datasets.
- Uses FDMS_CALLSIGNS_STANDARD.csv for standard tricodes.
- Uses FDMS_CALLSIGNS_NONSTANDARD CALLSIGNS.csv for nonstandard callsign roots.
- Uses FDMS_LOCATIONS.csv as full location reference.
- Uses FDMS_AIRCRAFT_TYPES.csv.
- Produces Analyst Workbook v2 with priority, hygiene, ranked candidates, diagnostics, deduplication, and summary metrics.
```

Recent confirmed VKB-enrichment improvements on the May test file:

```text
Unknown operator candidates reduced from 302 to 104.
Unknown location candidates reduced from 695 to 580.
Raw extracted rows reduced to 4017.
Unique strict extracted movements reduced to 2165.
Analyst priority unique movements around 332.
```

The system is working. The next work is not more raw extraction; it is classification accuracy, GUI reliability, semantic precision, NSCD correlation, and eventually historical intelligence.

---

## 2. Design Principles to Preserve

Do not:

- replace deterministic logic with opaque AI;
- abandon auditability;
- weaken schema discipline;
- use uncontrolled fuzzy parsing for pass logic;
- silently suppress or delete raw records;
- treat “not in VKB” as inherently suspicious;
- treat every nonblank `RM` value as a civil registration.

Do:

- preserve analyst trust;
- preserve reproducibility;
- keep reasoning explainable;
- keep workbook outputs transparent;
- distinguish intelligence signals from VKB coverage gaps;
- preserve raw audit rows even when summary/priority sheets use representative movements;
- treat uncertain interpretations as reviewable ambiguity, not forced certainty.

---

## 3. Core Architecture Direction

Vectis should continue to operate on the recursive intelligence loop:

```text
NM/NOP Export
    ↓
Deterministic Pass Engine
    ↓
Reference Correlation
    ↓
Analyst Workbook
    ↓
Unknowns / Gaps / Candidates
    ↓
VKB Enrichment
    ↓
Improved Future Detection
```

The product is both:

```text
1. an extraction/triage engine
2. a VKB quality-improvement engine
```

Future development should reinforce both roles.

---

## 4. Implementation Sequence Overview

The agreed order of work is:

```text
0. Repository hygiene and local data protection.
1. GUI progress bar, heartbeat, threaded processing, and Vectair styling.
2. ARCID/RM/airframe identifier classification.
3. NSCD contextual candidate matching.
4. VKB-gap semantics and scoring refinement.
5. Enhanced summary statistics and impact-priority reporting.
6. Continued VKB enrichment loop.
7. Historical memory and behavioural intelligence.
```

Reason for this sequence:

```text
First make the app reliable to use.
Then make the raw classification layer truthful.
Then add NSCD intelligence correlation.
Then refine scoring.
Then improve summary/reporting.
Then build historical intelligence.
```

Do not refine high-level scoring before ARCID/RM classification is fixed. Bad classifications with better-looking scores would reduce analyst trust.

---

# Phase 0 — Repository Hygiene and Local Data Protection

## Objective

Prevent local operational input/output files from being accidentally committed.

## Required `.gitignore`

The root `.gitignore` should include:

```gitignore
__pycache__/
*.py[cod]
.ruff_cache/

# Local NM/NOP inputs and generated analyst workbooks
input/*.csv
input/*.xlsx
output/*.xlsx
output/*.xlsm
```

## Acceptance

```text
- input/May test Incomplete.csv is ignored.
- output/*.xlsx remains ignored.
- git status is clean except intentional tracked source/config/VKB changes.
```

---

# Phase 1 — GUI Reliability, Progress, Heartbeat, and Vectair Styling

## Ticket

```text
VECTIS-GUI-001 — Add progress bar, heartbeat, threaded processing, and Vectair styling
```

## Objective

Improve usability and prevent the GUI from appearing hung during long Pandas/openpyxl processing.

Add:

```text
- progress bar;
- heartbeat indicator;
- current-stage status text;
- background-threaded file processing;
- visual alignment with Vectair / Flite styling.
```

## Required GUI Behaviour

Processing must not run on the Tkinter main thread.

Required implementation pattern:

```text
- process_file() starts a worker thread.
- Worker calls triage_file().
- Worker emits progress events through queue.Queue.
- Tkinter polls the queue with after().
- Only the Tkinter main thread updates widgets.
```

## Progress Stages

Use stage-based determinate progress:

| Stage | Progress |
|---|---:|
| Idle / ready | 0% |
| Reading input file | 5% |
| Validating canonical schema | 10% |
| Loading config and VKB references | 18% |
| Adding duplicate / movement identity metadata | 28% |
| Classifying callsigns | 36% |
| Running deterministic passes | 52% |
| Applying analyst scoring and representative movement logic | 65% |
| Building ranked VKB candidate sheets | 76% |
| Building diagnostics and summary sheets | 86% |
| Writing Excel workbook | 95% |
| Complete | 100% |

Display text example:

```text
Running deterministic passes... 52%
```

## Heartbeat

Add a heartbeat/status indicator that updates every 500–1000 ms while processing.

Acceptable display examples:

```text
Engine: active · 00:01:17
Heartbeat: active · 16:42:08
```

States:

```text
Engine: ready
Engine: active
Engine: complete
Engine: error
```

## Engine Progress Callback

Extend `triage_file()` to accept an optional progress callback:

```python
def triage_file(
    input_path: str | Path,
    output_dir: str | Path,
    reference_dir: str | Path | None = None,
    progress_callback=None,
) -> TriageResult:
```

Add helper:

```python
def _emit_progress(callback, stage: str, percent: int) -> None:
    if callback:
        callback(stage, percent)
```

Existing CLI usage must remain valid:

```python
triage_file(input_path, output_dir, reference_dir)
```

## Vectair / Flite Styling Direction

The GUI should align visually with Vectair and Flite:

```text
muted blue-grey header
off-white / pale grey workspace
brown/bronze accent
functional rectangular buttons
restrained borders
operational/logging feel
```

Suggested constants:

```python
VECTAIR_HEADER = "#6F8890"
VECTAIR_HEADER_DARK = "#4F6870"
VECTAIR_ACCENT = "#6B5638"
VECTAIR_BG = "#EEF2F1"
VECTAIR_PANEL = "#F7F8F7"
VECTAIR_BORDER = "#BFC8C8"
VECTAIR_TEXT = "#1D2A2E"
VECTAIR_MUTED = "#5F6D70"
VECTAIR_OK = "#4D8A57"
VECTAIR_WARN = "#B3842F"
VECTAIR_ERROR = "#A64040"
```

Header style:

```text
[Vectis logo]  VECTIS
               Aviation Intelligence Triage
```

## Acceptance

```text
- GUI remains responsive during processing.
- Progress bar updates through stages.
- Heartbeat continues updating while processing.
- Process button is disabled during active run.
- Existing browse/process/open-folder behaviour is preserved.
- Logo loads from assets/vectis.png.
- Missing-logo fallback still works.
- python -m py_compile triage_engine.py vectis_gui.py passes.
- python validate_sample.py passes.
```

---

# Phase 2 — ARCID, RM, Registration, and Airframe Identifier Classification

## Ticket

```text
VECTIS-ARCID-001 — Add contextual ARCID/RM/airframe identifier classification
```

## Objective

Fix the classification layer so the unknown-operator queue contains likely operator designators, not fragments of civil registrations, military serials, tactical identifiers, or word-like registration collisions.

## Core Problem

Some ARCID values are being misclassified. Example:

```text
HBJAZ
```

This is likely Swiss registration `HB-JAZ`, but a simplistic classifier can incorrectly extract `HBJ` as an unknown tricode.

Other problem examples:

```text
F-EVER → FEVER
M-OOSE → MOOSE
C-RAZY → CRAZY
```

NOP removes punctuation, so personalised registrations can look like words or longform callsigns.

## Core Principle

Do not classify by ARCID shape alone.

Use:

```text
ARCID + RM + ATYP + registration patterns + military serial patterns + NSCD + VKB
```

## Required Classification Hierarchy

Use this order:

```text
1. Blank / malformed
2. Exact ARCID/RM airframe identifier match
3. Civil registration pattern match
4. Military serial / tactical identifier pattern match
5. Known word-like registration match
6. Known NSCD / longform root match
7. Known interest root match
8. Strict standard tricode flight ID
9. Ambiguous registration-or-longform
10. Other unknown form
```

## Strict Tricode Rule

Only classify as `TRICODE_STYLE` if the ARCID matches:

```regex
^[A-Z]{3}[0-9][A-Z0-9]*$
```

Do not treat pure-letter strings as tricodes.

This prevents:

```text
HBJAZ → HBJ
FEVER → FEV
MOOSE → MOO
CRAZY → CRA
```

from entering the ordinary unknown-tricode queue.

---

## Civil Registration Patterns

Add:

```text
config/civil_registration_patterns.csv
```

Columns:

```text
COUNTRY,PREFIX,NORMALISED_REGEX,CONFIDENCE,NOTES
```

Seed with at least:

```csv
COUNTRY,PREFIX,NORMALISED_REGEX,CONFIDENCE,NOTES
Switzerland,HB,^HB[A-Z]{3,4}$,HIGH,Hyphen-stripped HB registrations
France,F,^F[A-Z]{4}$,MEDIUM,F-AAAA civil/state collision possible
Isle of Man,M,^M[A-Z]{4}$,MEDIUM,M-AAAA wordlike registration collision possible
Canada,C,^C[FGI][A-Z]{3}$,HIGH,C-F/G/I registrations
Australia,VH,^VH[A-Z]{3}$,HIGH,VH registrations
New Zealand,ZK,^ZK[A-Z]{3}$,HIGH,ZK registrations
Belgium,OO,^OO[A-Z]{3}$,HIGH,OO registrations
Netherlands,PH,^PH[A-Z]{3}$,HIGH,PH registrations
Denmark,OY,^OY[A-Z]{3}$,HIGH,OY registrations
Sweden,SE,^SE[A-Z]{3}$,HIGH,SE registrations
Norway,LN,^LN[A-Z]{3}$,HIGH,LN registrations
Finland,OH,^OH[A-Z]{3}$,HIGH,OH registrations
Ireland,EI,^EI[A-Z]{3}$,HIGH,EI registrations
Germany,D,^D[A-Z]{4}$,HIGH,D registrations
United Kingdom,G,^G[A-Z]{4}$,HIGH,G registrations
Austria,OE,^OE[A-Z]{3}$,HIGH,OE registrations
United States,N,^N[1-9][0-9]{0,4}[A-Z]{0,2}$,HIGH,N-number registrations
```

---

## Military Serial / Airframe Identifier Patterns

Civil registration logic is not enough. Military identifiers need their own layer.

Add:

```text
config/military_serial_patterns.csv
```

Columns:

```text
COUNTRY,SERVICE,PATTERN_NAME,NORMALISED_REGEX,CONFIDENCE,NOTES
```

Seed examples:

```csv
COUNTRY,SERVICE,PATTERN_NAME,NORMALISED_REGEX,CONFIDENCE,NOTES
United Kingdom,Joint,UK_MOD_SERIAL,^[A-Z]{2}[0-9]{3}$,HIGH,RAF/FAA/AAC style serials such as ZZ336
United States,Air Force,USAF_FISCAL_SERIAL_FULL,^[0-9]{2}[0-9]{4,5}$,MEDIUM,Fiscal-year serial without hyphen
United States,Navy/Marine Corps,USN_BUNO,^[0-9]{6}$,MEDIUM,Bureau Number candidate
France,Air Force/French State,FRENCH_FR_MARK,^FR[A-Z0-9]{3}$,MEDIUM,F-Rxxx mark normalised
France,Air Force,FRENCH_BASE_TACTICAL_CODE,^[0-9]{2,3}[A-Z]{2}$,LOW,May be tactical/base code not registration
Italy,Military,ITALIAN_MM_SERIAL,^MM[0-9]{3,6}$,HIGH,Italian military serial candidate
Australia,Military,RAAF_A_SERIAL,^A[0-9]{2,3}[0-9]{2,4}$,MEDIUM,Hyphen-stripped Axx-xxx style candidate
New Zealand,Military,RNZAF_NZ_SERIAL,^NZ[0-9]{3,5}$,HIGH,RNZAF serial candidate
Germany,Military,GERMAN_TACTICAL_NUMERIC,^[0-9]{4}$,LOW,Hyphen/plus stripped tactical code candidate; context required
```

France and the US require cautious handling:

```text
- US identifiers differ between Air Force, Navy/Marine Corps, and Army.
- US numeric RM values must not be dismissed as malformed.
- France can use civil-style marks, military marks, serials, tactical codes, or callsign disambiguators.
- French identifiers should often be labelled FRENCH_IDENTIFIER_AMBIGUOUS where context is mixed.
```

---

## RM Handling

Do not treat RM as simply “blank or registration”.

Add `RM_STATUS`:

| RM value | Status | Meaning |
|---|---|---|
| blank | `RM_BLANK` | Neutral / weak evidence only |
| `ONFILE` | `RM_ONFILE_SUPPRESSED` | Strong military/OAT indicator |
| `ON` | `RM_ONFILE_SUPPRESSED` | Strong military/OAT indicator |
| `KNOWN` | `RM_KNOWN_SUPPRESSED` | Strong military/OAT indicator |
| actual identifier | `AIRFRAME_IDENTIFIER_PRESENT` | Needs classification |
| other placeholder | `RM_PLACEHOLDER_OTHER` | Review |

Important rule:

```text
Blank RM is neutral.
ON / ONFILE / KNOWN are strong military/OAT/withheld-registration indicators.
```

## Internal Naming

Do not assume RM is always a civil registration.

Use:

```text
AIRFRAME_IDENTIFIER
```

with subclasses:

```text
CIVIL_REGISTRATION
MILITARY_SERIAL
TACTICAL_CODE
SUPPRESSED
AMBIGUOUS_IDENTIFIER
```

## ARCID/RM Equality Rule

If:

```text
normalised ARCID == normalised RM
```

and RM is not blank/ON/ONFILE/KNOWN, then this is strong evidence that ARCID is being used as an airframe identifier.

Then classify the identifier type:

```text
HBJAZ == HBJAZ → civil registration
ZZ336 == ZZ336 → UK military serial
169000 == 169000 → USN BuNo / US military serial candidate
FRBAE == FRBAE → French state/military mark candidate
MOOSE == MOOSE → possible wordlike registration / collision
```

Add:

```text
ARCID_AIRFRAME_IDENTIFIER_MATCH
```

---

## Aircraft-Type Context

Use aircraft type to support disambiguation.

Add:

```text
config/aircraft_type_context.csv
```

or use the existing local-code-to-ICAO military aircraft converter dataset.

Minimum columns:

```text
ATYP,CONTEXT,WEIGHT,NOTES
```

Contexts:

```text
MILITARY_OR_GOVERNMENT_LIKELY
GA_LIGHT
ULTRALIGHT
UNKNOWN
```

The existing local-code-to-ICAO converter should be used as:

```text
MILITARY_TYPE_INDICATOR
```

because any ICAO type appearing in that dataset is probably military/government relevant, even if the list is not exhaustive.

Use cautious wording:

```text
ATYP appears in military/government type indicator list
```

not:

```text
definitely military aircraft
```

---

## Word-Like Registrations

Add:

```text
config/wordlike_registrations.csv
```

Columns:

```text
REGISTRATION,NORMALISED,EXPECTED_ATYP,COUNTRY,CONFIDENCE,NOTES
```

Example:

```csv
REGISTRATION,NORMALISED,EXPECTED_ATYP,COUNTRY,CONFIDENCE,NOTES
F-EVER,FEVER,,France,HIGH,Known word-like registration
M-OOSE,MOOSE,,Isle of Man,HIGH,Known word-like registration
C-RAZY,CRAZY,,Canada,HIGH,Known word-like registration
```

This lets Vectis distinguish:

```text
MOOSE as known longform
```

from:

```text
M-OOSE as personalised registration
```

using RM, ATYP, and route context.

---

## ARCID Evidence Scoring

Add separate scores:

| Field | Meaning |
|---|---|
| `ARCID_REGISTRATION_SCORE` | Evidence that ARCID is an airframe identifier/registration |
| `ARCID_LONGFORM_SCORE` | Evidence that ARCID is longform/OAT/nonstandard callsign |
| `ARCID_CLASSIFICATION_CONFIDENCE` | `HIGH`, `MEDIUM`, `LOW`, `AMBIGUOUS` |
| `ARCID_COLLISION_FLAG` | Both interpretations plausible |
| `ARCID_COLLISION_REASON` | Explanation |

Registration evidence:

| Evidence | Score |
|---|---:|
| ARCID == RM, actual value | +100 |
| Known registration database match | +100 |
| Known wordlike registration match | +90 |
| High-confidence civil registration pattern | +60 |
| Military serial pattern match | +60 |
| GA/light type | +30 |
| Ultralight type | +40 |
| RM blank + ARCID matches registration pattern | +20 |
| RM is ON/ONFILE/KNOWN | -30 |

Longform/OAT evidence:

| Evidence | Score |
|---|---:|
| Known longform root | +100 |
| Known interest root | +90 |
| Word/root + numeric suffix | +60 |
| Military/government type indicator | +50 |
| Military/state route | +40 |
| RM is ON/ONFILE/KNOWN | +45 |
| NSCD candidate match | +50 |
| ARCID == RM | -50 unless conflict/review |

---

## New Output Fields

Add to audit/priority sheets:

```text
RM_STATUS
RM_SUPPRESSION_FLAG
RM_IDENTIFIER_CLASS
RM_IDENTIFIER_COUNTRY_HINT
RM_IDENTIFIER_SERVICE_HINT
RM_PATTERN_MATCH
RM_PATTERN_CONFIDENCE
RM_CONTEXT_CONFLICT_FLAG
RM_CLASSIFICATION_EXPLANATION

ARCID_AIRFRAME_IDENTIFIER_MATCH
ARCID_REGISTRATION_SCORE
ARCID_LONGFORM_SCORE
ARCID_CLASSIFICATION_CONFIDENCE
ARCID_REGISTRATION_CANDIDATE
ARCID_POSSIBLE_REG_PREFIX
ARCID_POSSIBLE_REG_COUNTRY
ARCID_LONGFORM_CANDIDATE
ARCID_AMBIGUITY_FLAG
ARCID_AMBIGUITY_REASON
ARCID_COLLISION_FLAG
ARCID_COLLISION_REASON
ARCID_CONTEXT_EXPLANATION
ARCID_REVIEW_RECOMMENDATION

ATYP_CONTEXT
ATYP_CONTEXT_SOURCE
ATYP_CONTEXT_CONFIDENCE
MILITARY_TYPE_INDICATOR
```

## New Workbook Sheet

Add:

```text
ARCID_CLASSIFICATION_DIAGNOSTICS
```

One row per unique ARCID.

Suggested columns:

```text
ARCID
NORMALISED_ARCID
COUNT_RAW_ROWS
RM
NORMALISED_RM
ATYP
ARCID_RM_MATCH
CALLSIGN_FORM
ARCID_REGISTRATION_SCORE
ARCID_LONGFORM_SCORE
ARCID_CLASSIFICATION_CONFIDENCE
ARCID_TYPE_CONTEXT
ARCID_COLLISION_FLAG
ARCID_COLLISION_REASON
WHY_CLASSIFIED_THIS_WAY
REVIEW_RECOMMENDATION
```

## Acceptance Examples

```text
HBJAZ + RM HBJAZ → REGISTRATION_CALLSIGN_CONFIRMED_BY_RM; no unknown HBJ tricode.
HBJAZ + RM ONFILE + military type → conflict/review; not ordinary unknown tricode.
MOOSE + RM ONFILE + ATYP C17 → LONGFORM_NONSTANDARD_HIGH_CONFIDENCE.
MOOSE + RM MOOSE + ATYP SKRA → REGISTRATION_WITH_LONGFORM_COLLISION or REGISTRATION_CALLSIGN_CONFIRMED_BY_RM.
FEVER + RM FEVER → REGISTRATION_CALLSIGN_CONFIRMED_BY_RM.
FEVER + RM blank + ATYP C17 → AMBIGUOUS_REGISTRATION_OR_LONGFORM.
BAW123 → TRICODE_STYLE root BAW.
RCH700 → TRICODE_STYLE root RCH, handled as known-interest if applicable.
MOOSE12 → LONGFORM_NONSTANDARD.
Pure-letter values must not create UNKNOWN_OPERATOR_TRICODE.
```

---

# Phase 3 — NSCD Contextual Candidate Matching

## Ticket

```text
VECTIS-NSCD-001 — Add NSCD contextual candidate matching for longform/nonstandard callsigns
```

## Objective

Use the nonstandard callsign database as a candidate operator intelligence table, not just a root list.

Example: `MOOSE` has multiple known NSCD users:

```text
18th Aggressor Squadron / USAF / F16
155th Airlift Squadron / Air National Guard / C17 / Tennessee
437th Airlift Wing / USAF / C17 / Formation
Air Mobility Division / USAF / C30J
2 Canadian Forces Flying Training School / RCAF / HAWK
```

Vectis should correlate ARCID, ATYP, RM, route, country, force, and NSCD notes to generate ranked candidate hypotheses.

## Required NSCD Fields

Load full NSCD fields from:

```text
FDMS_CALLSIGNS_NONSTANDARD CALLSIGNS.csv
```

Use columns where present:

```text
CALLSIGN
SSR INDICATION
ICAO 3LD
UNIT OR OPERATOR
FORCE
ACFT TYPE
COUNTRY
NOTE
ASSOCIATED FIXED CALLSIGN
```

## Candidate Matching

For each longform/nonstandard/ambiguous ARCID:

```text
normalise ARCID/root
match against NSCD CALLSIGN
```

Then score each candidate.

## NSCD Candidate Scoring

| Evidence | Score |
|---|---:|
| Exact NSCD callsign match | +50 |
| Exact ATYP match | +60 |
| Related/family ATYP match | +30 |
| RM/serial pattern compatible with candidate force/country | +25 |
| ATYP appears in military/government type list | +20 |
| Route/geography supports candidate NOTE | +25 |
| Route involves military/state location | +20 |
| RM suppressed: ON/ONFILE/KNOWN | +15 |
| Candidate note says formation and movement appears formation-like | +15 |
| Conflicting ATYP | -50 |
| Conflicting country/force context | -30 |

## Route/Geography Clue

Use location VKB fields:

```text
LOCATION SERVED
AIRPORT
COUNTRY
NOTES
ICAO REGION
```

If NSCD `NOTE` contains a geographic clue such as `TENNESSEE`, and ADEP/ADES metadata supports Tennessee, add route-context score.

This must be labelled as inference:

```text
Route context supports 155th Airlift Squadron candidate because ADEP/ADES is associated with Tennessee.
```

not:

```text
Operator confirmed as 155th AS.
```

## Output Fields

Add:

```text
NSCD_MATCH
NSCD_MATCH_COUNT
NSCD_EXACT_TYPE_MATCH_COUNT
NSCD_TOP_CANDIDATE
NSCD_TOP_FORCE
NSCD_TOP_COUNTRY
NSCD_TOP_ACFT_TYPE
NSCD_TOP_SCORE
NSCD_CONFIDENCE
NSCD_ALTERNATE_CANDIDATES
NSCD_MATCH_EXPLANATION
NSCD_ROUTE_CONTEXT_MATCH
NSCD_RM_CONTEXT_MATCH
```

## New Workbook Sheet

Add:

```text
NSCD_CANDIDATE_ANALYSIS
```

One row per candidate match:

```text
ARCID
NORMALISED_ARCID
RM
ATYP
ADEP
ADES
NSCD_CALLSIGN
CANDIDATE_UNIT_OR_OPERATOR
CANDIDATE_FORCE
CANDIDATE_ACFT_TYPE
CANDIDATE_COUNTRY
CANDIDATE_NOTE
SCORE
EVIDENCE_SUMMARY
CONFLICT_SUMMARY
```

## Acceptance Examples

```text
MOOSE + C30J → top candidate Air Mobility Division / USAF, high confidence.
MOOSE + C17 → ambiguous between C17 candidates unless route/RM context supports one.
MOOSE + C17 + Tennessee route clue → boost 155th Airlift Squadron / Tennessee ANG.
MOOSE + SKRA + RM MOOSE → likely wordlike registration; NSCD collision flagged but not high-confidence longform.
```

---

# Phase 4 — VKB-Gap Semantics and Scoring Refinement

## Ticket

```text
VECTIS-AW2-REFINE-001 — Refine VKB-gap semantics and scoring
```

## Objective

Separate intelligence significance from VKB coverage gaps.

Core rule:

```text
Missing from VKB ≠ operational interest by default.
Known-interest root ≠ unknown operator.
```

## Required Semantic Split

Replace or supplement `UNKNOWN_OPERATOR_TRICODE` with clearer categories:

| Category | Meaning |
|---|---|
| `OPERATOR_NOT_IN_LOADED_VKB` | True tricode-style operator root absent from standard VKB |
| `KNOWN_INTEREST_REFERENCE_COVERAGE_GAP` | Root known from interest config but absent from standard VKB |
| `NON_TRICODE_CALLSIGN_FORM` | Registration/longform/ambiguous value; not operator queue material |
| `REGISTRATION_FRAGMENT_FALSE_POSITIVE` | Legacy/diagnostic classification |
| `NSCD_KNOWN_NONSTANDARD` | Known nonstandard callsign family |
| `AIRFRAME_IDENTIFIER_CALLSIGN` | ARCID used as registration/serial |

## RCH Example

`RCH` can be:

```text
known operationally interesting
```

and:

```text
absent from the standard operator VKB
```

but it should not be:

```text
ordinary unknown operator
```

Correct wording:

```text
Known interest callsign RCH; standard-operator VKB coverage gap.
```

## Split Scores

Add:

```text
ANALYST_SCORE
VKB_ENRICHMENT_SCORE
```

Missing VKB data should not heavily inflate analyst score by itself.

| Signal | Analyst score | VKB enrichment score |
|---|---:|---:|
| Known interest callsign | high | low/none |
| Longform / NSCD match | high | medium |
| Military/state location | high | low |
| Sensitive region | high | low |
| Aircraft type of interest | high | low |
| Operator absent from VKB | low/none alone | high |
| Location absent from VKB | low/none alone | high |
| Known-interest coverage gap | medium | medium/high |
| Registration ambiguity | review/context | medium |

## Operational Interest Logic

Missing operator/location alone:

```text
VKB_UPDATE_REQUIRED = TRUE
OPERATIONAL_INTEREST = FALSE
```

Missing operator/location plus real signal:

```text
VKB_UPDATE_REQUIRED = TRUE
OPERATIONAL_INTEREST = TRUE
```

## Diagnostics Categories

Split diagnostics into clearer categories:

```text
TRICODE_NOT_IN_VKB
KNOWN_INTEREST_COVERAGE_GAP
NON_TRICODE_CALLSIGN_FORM
MALFORMED_OR_BLANK
REGISTRATION_OR_AIRFRAME_IDENTIFIER
NSCD_KNOWN_NONSTANDARD
```

## Acceptance

```text
- RCH no longer appears as ordinary unknown operator.
- Pure-letter registration-like values do not appear in unknown operator ranking.
- Missing VKB data alone is hygiene/enrichment, not operational interest.
- Analyst score and VKB enrichment score are separate.
- README explains intelligence signal vs VKB coverage gap.
```

---

# Phase 5 — Enhanced Summary Statistics and Impact-Priority Reporting

## Ticket

```text
VECTIS-STATS-001 — Enhance summary statistics, VKB coverage metrics, and impact-priority reporting
```

## Objective

Make the workbook more useful for both operational triage and VKB maintenance.

## Summary Blocks

Restructure `20_SUMMARY` into clear blocks:

```text
INPUT ACCOUNTING
DEDUPLICATION
ANALYST WORKLOAD
VKB MAINTENANCE WORKLOAD
VKB COVERAGE
SIGNAL DISTRIBUTION
TOP UNKNOWN OPERATORS
TOP MISSING LOCATIONS
TOP OPERATIONAL CALLSIGN ROOTS
TOP ROUTES BY OPERATIONAL INTEREST
TOP AIRCRAFT TYPES OF INTEREST
TOP MULTI-SIGNAL COMBINATIONS
KNOWN-INTEREST COVERAGE GAPS
```

## VKB Coverage Metrics

Add:

| Metric | Definition |
|---|---|
| `observed tricode roots` | Unique tricode-style roots in input |
| `known tricode roots in loaded VKB` | Observed roots present in standard callsign VKB |
| `operator VKB coverage percent` | known / observed |
| `observed location codes` | Unique ADEP/ADES/ALT1/ALT2 codes |
| `known location codes in loaded VKB` | Observed locations present in VKB |
| `location VKB coverage percent` | known / observed |
| `observed NSCD roots` | Unique longform/nonstandard roots |
| `known NSCD roots` | Observed roots present in NSCD |
| `NSCD coverage percent` | known / observed where relevant |

## Signal Distribution

Add table:

| Signal | Raw rows | Unique movements | Representative rows |
|---|---:|---:|---:|
| Known interest callsign | x | y | z |
| Longform/nonstandard | x | y | z |
| NSCD candidate match | x | y | z |
| Military/state location | x | y | z |
| Sensitive region | x | y | z |
| Aircraft type of interest | x | y | z |
| Military type indicator | x | y | z |
| Operator not in loaded VKB | x | y | z |
| Location not in loaded VKB | x | y | z |
| ZZZZ aerodrome | x | y | z |
| Airframe identifier ambiguity | x | y | z |

## VKB Impact-Priority Sheet

Add:

```text
07_VKB_IMPACT_PRIORITY
```

Combined operator/location/reference queue:

| Rank | Candidate type | Code | Unique movements affected | Raw rows affected | Max analyst score | VKB enrichment score | Suggested action |
|---:|---|---|---:|---:|---:|---:|---|
| 1 | LOCATION | LIMC | 54 | x | y | z | enrich_location |
| 2 | LOCATION | LSGG | 52 | x | y | z | enrich_location |
| 3 | OPERATOR | CSW | 17 | x | y | z | enrich_operator |
| 4 | KNOWN_INTEREST_GAP | RCH | 87 | x | y | z | add_or_mark_interest_coverage |

## Known-Interest Coverage Gaps

Add block/sheet:

```text
KNOWN_INTEREST_REFERENCE_COVERAGE_GAPS
```

Fields:

```text
ROOT
RAW_ROWS
UNIQUE_MOVEMENTS
MAX_ANALYST_SCORE
STANDARD_VKB_PRESENT
INTEREST_CONFIG_PRESENT
NSCD_PRESENT
SUGGESTED_ACTION
```

## Acceptance

```text
- Summary is clearer and grouped.
- VKB coverage percentages are visible.
- Signal distribution explains why rows are being flagged.
- Combined impact-priority list tells user what to update next.
- Known-interest coverage gaps are separated from unknown operators.
```

---

# Phase 6 — Continued VKB Enrichment Loop

## Objective

Continue using Vectis output to improve the VKB.

Current top missing locations after latest enrichment:

```text
LIMC
LSGG
LIRA
LKPR
HEGN
EPKT
ESSA
LZIB
LIRF
ETAD
```

Current top unknown operators:

```text
CSW
HFA
SWU
LUA
MEE
HBV
JSD
RMB
JSF
SQY
```

## Operating Method

Repeat:

```text
Run May file
→ Review 05_UNKNOWN_OPERATOR_RANKED
→ Review 06_UNKNOWN_LOCATION_RANKED
→ Review 07_VKB_IMPACT_PRIORITY once implemented
→ Update VKB
→ Rerun same May file
→ Compare 20_SUMMARY
```

## Metrics to Track

```text
unknown operator candidate count
unknown location candidate count
VKB hygiene-only raw rows
VKB hygiene-only unique movements
analyst priority unique movements
multi-signal unique movements
operator VKB coverage percent
location VKB coverage percent
```

## Goal

Reduce obvious high-impact VKB gaps until remaining candidates are:

```text
- obscure;
- genuinely unknown;
- low frequency;
- operationally interesting;
- worth investigation rather than simple data entry.
```

---

# Phase 7 — Historical Memory and Behavioural Intelligence

## Ticket Family

```text
VECTIS-HIST-001 onward
```

## Why This Comes Later

Do not build historical anomaly scoring until:

```text
- ARCID/RM classification is reliable;
- NSCD candidate matching exists;
- VKB gaps are semantically separated;
- movement identity/deduplication is stable.
```

## Future Capabilities

Add persistent historical store tracking:

```text
first seen
last seen
seen count
operator/root history
aircraft type history
route history
registration/airframe history
NSCD candidate history
location pair rarity
operator/type combinations
operator/route combinations
new root seen
new location seen
new aircraft type for root
route anomaly
behavioural change
```

## Identity Model

Use three levels:

| Level | Meaning |
|---|---|
| `RAW_RECORD_ID` | One row from one NM/NOP export |
| `MOVEMENT_ID` | One probable real flight/movement |
| `ENTITY_PATTERN_ID` | Recurring operator/route/type/callsign behaviour |

## Historical Outputs

Potential sheets:

```text
HISTORICAL_FIRST_SEEN
ROUTE_RARITY
OPERATOR_BEHAVIOUR_CHANGE
AIRFRAME_HISTORY
NEW_OR_CHANGED_PATTERNS
```

## Long-Term Goal

Transform Vectis from:

```text
deterministic triage and VKB enrichment workbook
```

into:

```text
aviation OSINT behavioural intelligence platform
```

---

# Immediate Next Tickets

If Codex is currently handling GUI progress/styling, continue with that first:

```text
VECTIS-GUI-001 — Add progress bar, heartbeat, threaded processing, and Vectair styling
```

Then implement:

```text
VECTIS-ARCID-001 — Add contextual ARCID/RM/airframe identifier classification
```

Then:

```text
VECTIS-NSCD-001 — Add NSCD contextual candidate matching
```

Then:

```text
VECTIS-AW2-REFINE-001 — Refine VKB-gap semantics and scoring
```

Then:

```text
VECTIS-STATS-001 — Enhance summary statistics, VKB coverage metrics, and impact-priority reporting
```

---

# Critical Acceptance Principle

The most important acceptance criterion across the next several phases is:

```text
The unknown-operator queue must contain likely operator designators, not fragments of civil registrations, military serials, tactical identifiers, or word-like registration collisions.
```

The second most important principle is:

```text
The workbook must clearly distinguish operational intelligence signals from VKB coverage gaps.
```

