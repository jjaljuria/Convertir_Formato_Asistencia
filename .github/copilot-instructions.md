# Copilot instructions for Convertir Formato de Asistencia

## Build, test, and validation

- Package manager: this repo declares `pnpm` in `package.json` (`packageManager: "pnpm@10.26.2"`). Prefer `pnpm` for install/test commands.
- Install dependencies:
  - `pnpm install`
- Run the full test suite:
  - `pnpm test`
  - Equivalent direct runner: `pnpm vitest run`
- Run a single test file:
  - `pnpm vitest run index.test.js`
- Node version expectation: project docs and CI target Node 20+; GitHub Actions runs on Node 20.x and 22.x.
- There is no dedicated lint script in `package.json`; validation is currently via Vitest tests.

## High-level architecture

- This repository is a small Node.js CLI utility that converts attendance Excel files into a cleaner Excel report.
- The main implementation lives in `index.js` and exports `generarReporte`, which is the core conversion pipeline used by the CLI and by programmatic callers.
- Flow of the conversion:
  1. Validate input/output paths and ensure the source file exists.
  2. Read the first worksheet from the Excel file using `xlsx`.
  3. Convert Excel rows into JSON records.
  4. Normalize the input columns and group records by user/date.
  5. Build a calendar across the date range, then determine per-user attendance for each day.
  6. Split results by department and create one worksheet per department with `ExcelJS`.
  7. Write the final `.xlsx` output file.
- The CLI entrypoint also provides `--help` and `--version` behavior in the same file, and the binary name is `cofa` via `package.json` `bin`.
- Tests live in `index.test.js` and validate behavior through mocked `xlsx` and `exceljs` dependencies; the suite is designed around the current conversion logic rather than browser/UI assumptions.

## Key conventions in this repo

- The project is written in JavaScript with ESM (`"type": "module"`).
- Naming conventions are mostly Spanish and domain-specific, for example: `generarReporte`, `generarCalendario`, `obtenerTrabajadores`, `generarMapaAsistencia`.
- The workbook output uses department-based worksheets, not a single flat sheet. Each worksheet is created from `departamentos` discovered in the input data.
- Default input column names are Spanish and expected to match the attendance export format described in `README.md`:
  - `ID de usuario`
  - `Nombre`
  - `Departamento`
  - `Fecha/Hora`
- Default output column names are also defined in code and can be overridden via the optional `usuarioConfig` argument.
- File/date handling is important: records are grouped by `user + formatted date`, and time extraction uses the first/last timestamp of the day as entry/exit.
- The project treats weekend dates as a special state (`TRABAJÓ FIN DE SEMANA` or `FIN DE SEMANA`) and missing entries as `FALTÓ`.
- This repo is CLI/data-processing oriented, so changes usually touch `index.js` and the related test file rather than larger application frameworks or build tooling.

## Repository-specific notes

- `README.md` is the authoritative user-facing description of the tool and its CLI contract.
- CI workflow is defined at `.github/workflows/ci.yml`; it installs dependencies via pnpm and runs the test suite against Node 20.x and 22.x.
- There are no extra AI or agent instruction files in the repository root, so this file should be treated as the main repository-local guidance for Copilot sessions.
