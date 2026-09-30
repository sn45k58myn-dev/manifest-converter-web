# Manifest Converter Web: IT handover

## Purpose

This repository contains a static browser edition of Manifest Converter for
deployment on GitHub Pages. Staff select one 12-column manifest CSV and receive
a formatted `.xlsx` workbook download.

## Data flow

1. The user selects a CSV through the browser file picker.
2. JavaScript reads the file locally in that browser tab.
3. The app creates the workbook in browser memory.
4. The browser downloads the finished workbook.

The application has no backend endpoint. It does not submit CSV content to
GitHub or another service. Production dependencies are bundled into the static
site and are not loaded from third-party CDNs at runtime.

Preferences are stored in browser local storage. They contain only display and
sorting options, not manifest data.

## GitHub permissions

The deployment workflow uses the minimum standard GitHub Pages permissions:

- `contents: read` to check out the repository;
- `pages: write` to publish the static site;
- `id-token: write` for the GitHub Pages deployment identity.

The workflow runs linting, converter tests and the production build before
publishing. GitHub Actions are pinned to the commit versions shown in
`.github/workflows/deploy.yml`.

## Deployment

1. Create an approved GitHub repository.
2. Add the project files to the repository root and push to `main`.
3. In **Settings > Pages**, select **GitHub Actions** as the source.
4. Confirm the `Deploy Manifest Converter to GitHub Pages` workflow succeeds.
5. Review the Pages URL and its access policy before issuing it to staff.

No secrets or environment variables are required.

## Acceptance checks

- Select the supplied sample CSV.
- Confirm the page reports 5 orders and 2 vehicles.
- Enable Items and Delivery borders.
- Convert and download the workbook.
- Confirm sheets `Vehicle A` and `Vehicle B` exist.
- Confirm Vehicle A starts with bulky orders `1060000002` and `1060000004`.
- Confirm order `1060000003` contains the E-prefix cleanup and the
  Trade/Redelivery note.
- Confirm the Items column, Code128 barcodes and horizontal delivery separators
  are present.
- Confirm print setup is A4 landscape and one page wide.

## Limitations

GitHub Pages cannot continuously watch an Incoming folder or automatically
write to the dated Converted hierarchy. The Windows edition remains the
supported option for unattended folder automation.
