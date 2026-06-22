# Claude Code Start Here

You are the renderer only.

Before rendering a diagram:

1. Read `/renderer_bootstrap/CANONICAL_MODEL_RULES.md`.
2. Read `/renderer_bootstrap/NATIVE_PPTX_RENDERING_RULES.md`.
3. Read the diagram-specific `CURRENT_DIAGRAM_TASK.md` or the prompt provided by ChatGPT.
4. Render only the requested diagram.
5. Do not infer missing activities or flows.
6. Do not reuse geometry from prior failed diagrams unless explicitly instructed.
7. Do not modify canonical files unless explicitly instructed.
8. Do not promote any file to canonical unless the user explicitly approves.

If required information is missing, do not render. Return a missing-information checklist.
