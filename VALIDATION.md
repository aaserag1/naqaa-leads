# Naqaa | نقاء — Validation — 2026-09-14

## Executed locally

- Python 3.12; pinned dependencies installed and import-tested.
- Automated test command: python -m pytest -q tests
- Actual final output: 85 passed, 132 subtests passed in 13.44s
- Real Chromium browser: 151.0.7922.34. File and logo upload, metadata/header detection, filtering, dark mode, Excel/PDF/source-ZIP downloads were exercised.
- Browser JavaScript errors: 0. Failed requests: 0.
- Mobile viewport 390 × 844: fresh session, automatic sidebar and no document-wide horizontal overflow; main table scrolls horizontally within its own container.
- Synthetic browser fixture: 12 leads; 10 valid telephone links, 4 WhatsApp links. No calls or messages were initiated.
- Downloaded PDF: 3 A4-landscape pages, 14 clickable contact annotations. Arabic shaping, branding and table pagination visually inspected.
- Downloaded Excel: RTL sheets, cached summary values verified against the processed dataset, no Excel error cells. Application-owned formulas preserved; uploaded formula-like content is literal text.
- Originals archive: uploaded CSV bytes compared exactly with the ZIP member.
- Sources and exports contain synthetic data only.
- Approved platform identity: Naqaa | نقاء. Branding regressions verify page identity, custom company names and logos in Excel/PDF, neutral fallback report titles, and consistent deployment naming.

## Not executed or not provided

- Repository target: aaserag1/naqaa-leads. This report records local release validation; check GitHub for the current commit and workflow results. No hosted application deployment or live app URL is claimed.
- Docker image builds were not executed locally. GitHub Actions results are tracked separately under the repository Actions tab; they are not implied by the local test result.
- OIDC allowlisting and fail-closed behavior were tested with synthetic claims/configuration; no real identity-provider login was configured.
- Python 3.10 is a CI compatibility target, not a locally executed test environment.
- No load/SLA, penetration-test certification, real-lead migration, billing, multi-tenant database or advertising-platform API integration is claimed.
- Free-host memory/cold-start limits still apply. The app does not promise an always-on production SLA.

## Repository transport validation

The release tests were rerun in a clean staged tree containing assets/Cairo.ttf.b64 and no binary Cairo.ttf. The in-memory decoder and isolated PDF worker passed, including rejection of modified fonts. SHA-256 of the decoded font: 667c987182391c91f4e57a2f455b1794fb5e3ee6ca4ef3383e86bb690fa9c964. This is a lossless transport change, not a different font or a runtime network dependency.

## Security checks included

Malformed/oversized inputs, ZIP-bomb checks, XML entities, formula injection, ambiguous/corrupt phone values, explicit country prefixes, currency separation, negated readiness, duplicated contacts, original-value preservation, escaping HTML, blocking file/network PDF resources, logo validation, bounded PDF rendering, session separation and invalidation of stale exports.
