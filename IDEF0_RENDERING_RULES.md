# IDEF0 Rendering Rules for ATMOS Diagrams

## 1. Renderer role

The renderer is not the architect.

The renderer shall:

* Read the provided YAML.
* Render only what is present in the YAML.
* Preserve names exactly.
* Use strict IDEF0 semantics.
* Report missing information rather than inventing content.

The renderer shall not:

* Invent flows.
* Rename activities.
* Convert the diagram into a generic flowchart.
* Use diagonal connectors.
* Use curved freeform connectors.
* Use decision diamonds.
* Treat COWP as a node or decision-maker.

## 2. Required IDEF0 semantics

IDEF0 boxes represent functions or activities.

Function names shall be active verb phrases.

Arrow semantics:

* Inputs enter from the left side of a box.
* Controls enter from the top side of a box.
* Outputs exit from the right side of a box.
* Mechanisms connect to the bottom side of a box.
* Call arrows, if ever used, point downward from the bottom.

## 3. Connector geometry

All connectors shall be orthogonal.

Allowed:

* Horizontal line segments.
* Vertical line segments.
* Right-angle bends.
* 90-degree elbow routing.

Not allowed:

* Diagonal connectors.
* Freeform curves.
* Spline connectors.
* Connectors crossing through function boxes.
* Arrow endpoints entering inside a box.

Arrowheads:

* Use arrowheads only at the semantic receiving end.
* Boundary input arrows terminate at the left side of the target box.
* Boundary control arrows terminate at the top side of the target box.
* Boundary mechanism arrows terminate at the bottom side of the target box.
* Output arrows originate at the right side of the source box and exit to the right boundary.

## 4. Box rules

A-0 context diagram:

* Exactly one box.
* The box represents the top-level function.
* Boundary arrows show model inputs, controls, outputs, and mechanisms.

A0 and child decomposition diagrams:

* Use 3 to 6 boxes.
* Each box shall be labeled with:

  * Box number
  * Function name
  * Node ID where appropriate
* Preserve the provided activity names exactly.
* Do not add extra child functions.

## 5. Layout rules

Default A-0 layout:

* Single central box.
* Inputs on left.
* Controls on top.
* Outputs on right.
* Mechanisms on bottom.

Default A0 layout:

* Six boxes arranged left-to-right unless the YAML states otherwise.
* Recommended sequence: A1 → A2 → A3 → A4 → A5 → A6.
* Internal flows should be routed orthogonally between boxes.

## 6. Feedback routing

If feedback flows are ever provided:

* Control feedback shall route up and over.
* Input or mechanism feedback shall route down and under.
* Do not invent feedback flows.

## 7. Visual style

Use strict black-and-white technical drawing style:

* White background.
* Black box outlines.
* Black connector lines.
* Plain sans-serif font.
* No decorative icons.
* No color coding unless later requested by the architect.
* No gradients or shadows.
* Legible labels.
* Consistent margins.

## 8. File naming

Use the following file naming convention:

* ATMOS_A_MINUS_0.svg
* ATMOS_A0.svg
* ATMOS_A1.svg
* ATMOS_A2.svg

Do not overwrite an existing SVG unless explicitly instructed.

## 9. Validation behavior

Before rendering any diagram, check:

* Required YAML fields are present.
* Diagram type is supported.
* Box count is valid.
* Arrow sides match IDEF0 semantics.
* No unspecified flows are required for rendering.
* No extra flows are invented.

If required information is missing, do not render. Instead, return a missing-information checklist.
