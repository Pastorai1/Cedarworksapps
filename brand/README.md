# Cedarworks Apps — Brand assets

Business card designs for Cedarworks Apps. Brand palette: deep green
(`#143d28`), cream (`#f4ebd5`), orange (`#d98b3c`), tan (`#ece1c7`); type is
Fraunces (display serif) + Inter (sans).

## Print-ready files

| File | Description |
|------|-------------|
| `cedarworks-card-360print.png` | **Ordered design.** Single-sided green card sized to the 360onlineprint template — 3.7008 in × 2.126 in (94 × 54 mm), ~384 DPI. |
| `cedarworks-card-360print.pdf` | Same card as vector PDF (exact size). |
| `card-1side-green.png` | Single-sided green card, 3.75 × 2.25 in (0.125 in bleed → trims to 3.5 × 2 in). |
| `card-1side-tan.png` | Single-sided tan (light) alternate, same size. |
| `card-two-sided-front.png` | Two-sided layout — FRONT (brand: mark, wordmark, tagline). |
| `card-two-sided-back.png` | Two-sided layout — BACK (contact details). |

Contact details on the cards: James Chambers · james@cedarworksapps.com ·
417-370-2084 · cedarworksapps.com

## Editable sources (`sources/`)

Plain HTML/CSS — open in a browser to preview; edit text/colors directly.

| File | Renders |
|------|---------|
| `card-green-360print.html` | The ordered single-sided green card (360onlineprint size). |
| `card-1side-both-colors.html` | Single-sided green + tan side by side. |
| `card-two-sided.html` | Two-sided front + back. |

### Regenerating PNG/PDF from a source

The images were produced by rendering these HTML files headlessly (Chromium /
Playwright) — screenshot the card element for PNG, or "Print → Save as PDF"
with margins off and the page size set to the card size.
