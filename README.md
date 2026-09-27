# ChileEaster26 Geological Field Journal Generator

This project builds a formatted `.docx` geological field journal from structured JSON data and local photo assets.

## Canonical project files

- `/home/runner/work/ChileEaster26/ChileEaster26/index.js` — build entrypoint
- `/home/runner/work/ChileEaster26/ChileEaster26/travelData.json` — canonical trip data source
- `/home/runner/work/ChileEaster26/ChileEaster26/photos/` — photo assets referenced by `travelData.json`
- `/home/runner/work/ChileEaster26/ChileEaster26/src/` — modular generator and data pipeline code
- `/home/runner/work/ChileEaster26/ChileEaster26/scripts/` — validation and smoke-test scripts

Legacy variants and historical snapshots are archived under `/home/runner/work/ChileEaster26/ChileEaster26/misc/archive/`.

## Setup

```bash
npm install
```

## Canonical workflow

1. Edit `/home/runner/work/ChileEaster26/ChileEaster26/travelData.json`
2. Validate the data model:
   ```bash
   npm run validate
   ```
3. Build the journal:
   ```bash
   npm run build
   ```

## Commands

- `npm run validate` — strict schema validation for `travelData.json`
- `npm run build` — generate `Geological_Field_Journal_2026.docx`
- `npm run smoke-test` — run a build to a temp file and assert output exists
- `npm test` — run validation + smoke test

## Data model requirements

Each `days[]` item must contain:

- `day` (string)
- `title` (string)
- `description` (string)
- `images` (array of objects)

Each image object supports:

- `url` (required string)
- `caption` (optional string)

Legacy `image` fields are not supported.

## Optional output auto-open

By default, the build does not auto-open the generated document.

To auto-open on local desktop environments:

```bash
AUTO_OPEN=true npm run build
```
