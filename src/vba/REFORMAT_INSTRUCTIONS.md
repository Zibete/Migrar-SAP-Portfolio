REFORMAT INSTRUCTIONS

1) Strict rules
- DO NOT change any logic, comparisons, condition order, or operation order.
- DO NOT rename public macros, UserForm names, control names, named ranges, ListObject names, or any strings used as IDs.
- You MAY rename local variables and procedure parameters (only within a single procedure).
- Add per-file header (purpose, dependencies, inputs/outputs, risks).
- Keep existing `Option Explicit` if present; do NOT add it automatically.
- Split long lines using VBA line continuation `_` when applicable.

2) Suggested workflow
- Export all VBA files into `src/vba/*`.
- Run the scan tool (not included) or ask me to perform a read-only scan.
- I'll create explicit commits: (1) headers+format (2) local renames (3) README+report.

3) Naming / formatting
- Use 4 spaces for indentation consistently.
- Keep existing public procedure names untouched.
- Prefer descriptive local variable names when renaming (e.g., `iRow` -> `rowIndex`) but only inside procedure.

4) What I'll deliver
- Three separate commits with clear messages.
- `REPORT.md` listing files touched and exact cosmetic changes.
