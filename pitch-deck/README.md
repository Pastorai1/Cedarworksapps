# Cedarworks Apps: Investor Pitch Deck

A 16-slide pre-seed deck ($250K) for investor programs such as Initio Capital.

| File | Use |
|------|-----|
| `Cedarworks-Apps-Pitch-Deck.pdf` | Upload this to investor portals (e.g. Initio's "Upload deck to autofill") and send it to investors. |
| `Cedarworks-Apps-Pitch-Deck.pptx` | Editable PowerPoint version, with speaker notes on every slide. |
| `build.js` | Script that generates the `.pptx` (pptxgenjs). |

## Slides

1. Title · 2. Executive Summary · 3. Problem Statement · 4–7. Solution & Product
(overview, PastorAI, Ready Room, AppForge) · 8. Traction & Metrics · 9. Market
Opportunity · 10. Business Model · 11. Go-to-Market Strategy · 12. Competitive
Landscape · 13. Use of Funds · 14. Roadmap & Milestones · 15. Team · 16. The Ask

Each slide's small orange label uses the standard category name (e.g. "Problem
Statement", "Traction & Metrics") as plain, unspaced text, so AI deck scorers
such as Initio's can match every section.

## Update before sending

- **Traction (slide 8):** pre-revenue, ~30 trial members. Update with paying subscribers and MRR as they come in.
- **Unit economics (slide 10):** AI/hosting cost per user and 5% monthly churn are estimates. Replace them with real numbers once you have them.
- **Competitors (slide 12):** check the named examples and how they're described before sending.
- **Roadmap targets (slide 14):** 75 → 1,000 paying subscribers are targets, not forecasts. Change them if your plan differs.
- **Use of funds (slide 13):** the 40/30/10/10/10 split is a proposal.

## Rebuilding

```bash
npm install pptxgenjs react react-dom react-icons sharp
node build.js Cedarworks-Apps-Pitch-Deck.pptx
```
