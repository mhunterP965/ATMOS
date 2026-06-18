# ATMOS IDEF0 Rendering Status

## Files created

* IDEF0_RENDERING_RULES.md
* MODEL_BACKGROUND.md
* ATMOS_A_MINUS_0.yaml
* ATMOS_A0_DRAFT.yaml
* RENDERING_STATUS.md

## Diagram readiness

| Diagram | Status                            | Notes                                                                               |
| ------- | --------------------------------- | ----------------------------------------------------------------------------------- |
| A-0     | Ready to render                   | Complete context diagram packet.                                                    |
| A0      | Not ready for strict IDEF0 render | Boundary arrow allocation is not yet specified. Internal flows F1-F3 are specified. |

## Hard stop rule

Do not render any SVG until the user explicitly says:

"Render this diagram."

## Next expected user command

The next expected command is likely:

"Render this diagram: A-0 using ATMOS_A_MINUS_0.yaml."

When that command is received, render only the A-0 context diagram.
