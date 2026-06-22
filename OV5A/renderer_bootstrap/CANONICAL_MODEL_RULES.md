# ATMOS Canonical Model Rules

Architecture: ATMOS
Package/version: v1
Audience: DEVCOM
Marking: Distribution Statement C

Primary output format:
Native editable PowerPoint `.pptx`.

Renderer:
Claude Code in browser.

Architect / prompt author / QC:
ChatGPT.

Core semantic constraints:

* COWP is emergent.
* COWP is not a node.
* COWP is not a decision-maker.
* ATMOS provides descriptive weather context only.
* ATMOS does not perform route selection.
* ATMOS does not perform maneuver decisions.
* ATMOS does not exercise command authority.
* ATMOS does not make mission execution decisions.

IDEF0 rules:

* Activity boxes are rectangles.
* Function names are active verb phrases.
* Inputs enter from the left.
* Controls enter from the top.
* Outputs exit to the right.
* Mechanisms enter from the bottom.
* Boundary arrows must originate from or terminate at the diagram boundary as appropriate.
* Internal flows originate from child activity boxes and terminate at child activity box perimeters.
* All non-output arrowheads terminate at the outside perimeter of destination boxes.
* Outputs terminate at the right diagram boundary.
* No arrowhead may appear inside a box.
* No non-output arrowhead may dangle in whitespace.
* No connector may pass through an activity box.
* Connector glue is optional; correct visible endpoint geometry is mandatory.
* Use 3 to 4 child boxes per child decomposition unless explicitly approved otherwise.
* Do not render TBD placeholder boxes.
* Do not render redundant merge flows unless explicitly approved.
