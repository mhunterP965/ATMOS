# Native PPTX Rendering Rules

Use native editable PowerPoint objects only:

* rectangles for activity/function boxes
* editable text boxes for labels
* native editable straight or elbow/right-angle connectors
* native arrowheads

Do not use:

* embedded SVG
* embedded PNG/JPG
* rasterized diagrams
* Mermaid
* PlantUML
* Graphviz
* Draw.io
* generic automatic flowchart layout

Every diagram shall be one PPTX file.

Every diagram shall be one slide unless explicitly instructed otherwise.

Visual style:

* white background
* black lines
* black text
* black arrowheads
* no colors
* no gradients
* no shadows
* no icons
* no decision diamonds unless explicitly required for UML Activity diagrams later

Endpoint rules:

* Non-output arrowheads must terminate visibly at the outside perimeter of the destination activity box.
* Output arrowheads may terminate at the right diagram boundary.
* If PowerPoint connector glue places the visible arrowhead inside a box, do not use that glued endpoint.
* Use unglued elbow connectors or editable line segments when necessary.
* Correct visible endpoint geometry is mandatory.
* Connector glue is optional.

Final response after each diagram:

1. Confirm output file created.
2. Provide endpoint audit table.
3. Provide self-check table.
4. Do not promote to canonical unless the user explicitly approves.
