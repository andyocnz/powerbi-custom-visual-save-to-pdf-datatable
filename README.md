# Modern DataTable Visual

Buildable Power BI custom visual project with:
- Pagination (`Prev`, `Next`)
- `Load all` / `Restore paging`
- `Print / Save PDF` (opens print dialog; use Save as PDF)

## Project folder
- `moderndatatablebuild/`

## Build output
- `moderndatatablebuild/dist/moderndatatablebuild9AA7B8C7C17448ED8262BA952CFCBA05.1.0.0.0.pbiviz`

## Commands
From `moderndatatablebuild/`:

```bash
npm install
npm run start
npm run package
```

## Notes
- This is TypeScript custom visual code (not R visual).
- PDF is implemented via print workflow for host compatibility.
- Some `pbiviz` warnings about optional/recommended features are expected and do not block packaging.
