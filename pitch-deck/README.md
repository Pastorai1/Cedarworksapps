# Cedarworks Apps: Investor Pitch Deck

An 18-slide pre-seed deck ($250K) for investor programs such as Initio Capital.

| File | Use |
|------|-----|
| `Cedarworks-Apps-Pitch-Deck.pdf` | Upload this to investor portals (e.g. Initio's "Upload deck to autofill") and send it to investors. |
| `Cedarworks-Apps-Pitch-Deck.pptx` | Editable PowerPoint version, with speaker notes on every slide. |
| `build.js` | Script that generates the `.pptx` (pptxgenjs). |

## Slides

1. Title · 2. Executive Summary · 3. Problem Statement · 4–7. Solution & Product
(overview, PastorAI, Ready Room, AppForge) · 8. Traction & Metrics · 9. Market
Opportunity · 10. Why Now & Customer Segments · 11. Business Model ·
12. Go-to-Market Strategy · 13. Competitive Landscape · 14. Use of Funds ·
15. Roadmap & Milestones · 16. Financial Projections · 17. Team & Founders ·
18. The Ask

Each slide's small orange label uses the standard category name (e.g. "Problem
Statement", "Traction & Metrics") as plain, unspaced text, so AI deck scorers
such as Initio's can match every section.

## Update before sending

- **Traction (slide 8):** pre-revenue, ~30 trial members. Update with paying subscribers and MRR as they come in.
- **Unit economics (slide 11):** AI/hosting cost per user and 5% monthly churn are estimates. Replace them with real numbers once you have them.
- **Competitors (slide 13):** check the named examples and how they're described before sending.
- **Roadmap targets (slide 15):** 75 → 1,000 paying subscribers are targets, not forecasts. Change them if your plan differs.
- **Use of funds (slide 14):** the 40/30/10/10/10 split is a proposal.

- **Financial projections (slide 16):** a 5-year model built from the roadmap targets. Adjust it if your plan differs.
- **Problem stat (slide 3):** the Barna Group 2022 figure (42% of pastors considered quitting). Confirm it before sending.
- **Product screens (slides 5–7):** `pastorai-dashboard.png`, `readyroom-app.png` and `appforge-app.png` are recreations of the live apps, made from the founder's screenshots with personal emails removed. They're built from the `*-app.html` / `pastorai-dashboard.html` files. To use real screenshots, save them over those PNGs and rebuild.

## Rebuilding

```bash
npm install pptxgenjs react react-dom react-icons sharp
node build.js Cedarworks-Apps-Pitch-Deck.pptx
```
