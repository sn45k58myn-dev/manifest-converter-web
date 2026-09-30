# Manifest Converter Web

A browser edition of the proven Manifest Converter, prepared for GitHub Pages.

The page converts one manifest CSV at a time and downloads the formatted Excel
workbook. Processing happens in the browser. The CSV is not sent to a server by
this application.

## Preserved output behaviour

- Six standard columns, with the optional `Items` column.
- Code128 order barcodes.
- E-prefix removal and Trade/Redelivery note.
- `STH&BULK` note.
- Bulky-first sorting per vehicle using editable keywords.
- Standalone `Retail` in vehicle names displayed as `Hertz`.
- Optional medium dark-grey horizontal delivery separators.
- Hidden worksheet gridlines.
- A4 landscape print setup, fitted to one page wide.
- One worksheet per vehicle.
- Preferences saved in the current browser.
- `Developed by Chris Ryan` credit.

The browser version cannot watch an Incoming folder or save automatically into
the dated Converted hierarchy. Keep the Windows edition for unattended folder
automation.

## Run locally

Install Node.js 20.19 or newer, then run:

```powershell
npm ci
npm run dev
```

The development address appears in the terminal.

## Verify and build

```powershell
npm run lint
npm test
npm run build
```

The production website is created in `dist`.

## Publish with GitHub Pages

1. Create a GitHub repository and add the project files to its root.
2. Push the files to the `main` branch.
3. Open the repository's **Settings**, then **Pages**.
4. Set **Source** to **GitHub Actions**.
5. Open the **Actions** tab and confirm the deployment succeeds.

The included `.github/workflows/deploy.yml` checks, tests, builds and publishes
the site automatically after each push to `main`.

The packaged handover also includes a `sample` folder with a synthetic manifest
and its expected workbook output.

## Data and access notes for IT

- The web application has no backend and makes no request containing CSV data.
- The production JavaScript dependencies are included in the built site rather
  than loaded from a third-party CDN at runtime.
- Browser preferences are stored in local storage on the user's device.
- GitHub Pages access and repository visibility should be approved by the
  organisation before workplace use.
