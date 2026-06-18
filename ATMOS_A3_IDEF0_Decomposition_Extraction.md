# A3 IDEF0 Decomposition Extraction

## 1. Parent activity

| Field                | Extracted answer     | Source / rationale |
| -------------------- | -------------------- | ------------------ |
| Parent node          | A3                   | A0 baseline        |
| Parent function name | Quantify Uncertainty | A0 baseline        |
| Parent purpose       | Estimate uncertainty associated with observations and model-derived weather states and package results for downstream use without performing decision-making | OV-5a A3 description |

## 2. A3 child function candidates

| Child ID | Box number | Function name, verb phrase | Description | Source / rationale |
| -------- | ---------: | -------------------------- | ----------- | ------------------ |
| A31      |          1 | Estimate Observation Uncertainty | Quantify uncertainty associated with sensor observations based on sensor characteristics and conditions | OV-5a A3.1 |
| A32      |          2 | Estimate Model and State Uncertainty | Estimate uncertainty arising from model structure, inputs, and execution constraints | OV-5a A3.2 |
| A33      |          3 | Package Confidence Bounds and Risk Envelopes | Combine weather state and uncertainty into packaged confidence bounds and risk envelopes | OV-5a A3.3 |

## 3. A3 boundary ICOM allocation

### Inputs

| Parent input                                   | Target child function(s) | Notes | Source / rationale |
| ---------------------------------------------- | ------------------------ | ----- | ------------------ |
| F2 Local Micro-Weather State Estimate (IER-05) | A32, A33 | Used for model uncertainty estimation and confidence packaging | OV-5a A3 |

### Controls

| Control | Target child function(s) | Authority / source | Source / rationale |
| ------- | ------------------------ | ------------------ | ------------------ |
| Sensor uncertainty models; calibration/state-of-health policies | A31 | Sensor modeling/configuration | OV-5a A3.1 |
| Error propagation rules; model tuning parameters | A32 | Model configuration | OV-5a A3.2 |
| Confidence representation standards; formatting constraints | A33 | Data packaging rules | OV-5a A3.3 |

### Outputs

| Parent output                                  | Producing child function | Destination                                        | Source / rationale |
| ---------------------------------------------- | ------------------------ | -------------------------------------------------- | ------------------ |
| F3 Confidence Bounds & Risk Envelopes (IER-09) | A33 | A4 Federate and Maintain Federated Weather Context | OV-5a A3.3 |

### Mechanisms

| Mechanism                               | Target child function(s) | Type         | Source / rationale |
| --------------------------------------- | ------------------------ | ------------ | ------------------ |
| M1 Platforms hosting ATMOS node compute | A31, A32, A33 | System | OV-5a A3 |
| M2 ABLE-LBM / reduced models            | A32 | System/model | OV-5a A3.2 |
| ATMOS uncertainty module; sensor metadata | A31 | System/resource | OV-5a A3.1 |
| ABLE-LBM uncertainty routines          | A32 | System/resource | OV-5a A3.2 |
| ATMOS packaging/serialization; DDS data types | A33 | System/resource | OV-5a A3.3 |

## 4. A3 internal flows

| Flow ID | From child | To child | Flow name | Receiving ICOM type         | Source / rationale |
| ------- | ---------- | -------- | --------- | --------------------------- | ------------------ |
| A3-F1   | A31        | A32      | Observation Uncertainty Metrics (IER-07) | Input | OV-5a A3.1–A3.2 |
| A3-F2   | A32        | A33      | Model / State Uncertainty Estimates (IER-08) | Input | OV-5a A3.2–A3.3 |
| A3-F3   | A31, A32   | A33      | Combined uncertainty inputs | Input | OV-5a A3.3 |

## 5. Lower-level candidate details not shown as boxes

| Detail ID | Candidate lower-level action | Recommended parent child box | Source / rationale |
| --------- | ---------------------------- | ---------------------------- | ------------------ |
| D1 | Evaluate sensor characteristics | A31 | OV-5a A3.1 |
| D2 | Apply calibration/state-of-health adjustments | A31 | OV-5a A3.1 |
| D3 | Perform uncertainty propagation | A32 | OV-5a A3.2 |
| D4 | Apply model tuning parameters | A32 | OV-5a A3.2 |
| D5 | Combine uncertainty with weather state | A33 | OV-5a A3.3 |
| D6 | Format confidence outputs for dissemination | A33 | OV-5a A3.3 |

## 6. Uncertainty / risk-envelope details relevant to A3

| Detail ID | Detail | Likely IDEF0 role | Applies to child function | Source / rationale |
| --------- | ------ | ----------------- | ------------------------- | ------------------ |
| UDET-01 | Observation uncertainty metrics | Output | A31 | OV-5a A3.1 |
| UDET-02 | Model/state uncertainty estimates | Output | A32 | OV-5a A3.2 |
| UDET-03 | Confidence bounds and risk envelopes | Output | A33 | OV-5a A3.3 |
| UDET-04 | Sensor metadata | Mechanism | A31 | OV-5a A3.1 |
| UDET-05 | Error propagation rules | Control | A32 | OV-5a A3.2 |
| UDET-06 | Confidence representation standards | Control | A33 | OV-5a A3.3 |

## 7. Open issues

| Issue ID | Issue | Why it matters | Candidate resolution |
| -------- | ----- | -------------- | -------------------- |
| A3-Q1    | Whether observation uncertainty should be exposed as external output | Impacts diagram scope | Keep internal per OV-3 mapping |
| A3-Q2    | Whether uncertainty combination is explicit vs implicit step | Affects clarity of A33 inputs | Represent as combined input into A33 |
| A3-Q3    | Whether model and observation uncertainty should be separate diagrams | Affects diagram complexity | Keep within single A3 decomposition |

## 8. Recommended rendering notes

| Item                        | Recommendation | Rationale |
| --------------------------- | -------------- | --------- |
| Recommended box count       | 3 | Source defines 3 A3 activities |
| Recommended layout          | A31 → A32 → A33 | Reflects uncertainty layering |
| Flows to avoid rendering    | Any inferred branching logic | Not source-supported |
| Source-supported controls   | Sensor models; error propagation; representation standards | Explicit in OV-5a |
| Source-supported mechanisms | Uncertainty modules; ABLE-LBM routines; platform compute | Explicit in OV-5a |
