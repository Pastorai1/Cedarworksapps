# Cedarworks Apps: Investor Pitch Deck

A 15-slide pre-seed deck ($250K) for investor programs such as Initio Capital.

| File | Use |
|------|-----|
| `Cedarworks-Apps-Pitch-Deck.pdf` | Upload this to investor portals (e.g. Initio's "Upload deck to autofill") and send it to investors. |
| `Cedarworks-Apps-Pitch-Deck.pptx` | Editable PowerPoint version, with speaker notes on every slide. |
| `build.js` | Script that generates the `.pptx` (pptxgenjs). |

## Slides

1. Title and the ask · 2. Problem · 3. Solution · 4. PastorAI · 5. Ready Room ·
6. AppForge · 7. Traction · 8. Market · 9. Business model · 10. Go-to-market ·
11. Competition · 12. Use of funds · 13. 18-month roadmap · 14. Team · 15. The ask

## Update before sending

- **Traction (slides 4, 5, 7):** trial counts (~15 per app). Add paid conversions, MRR, ratings and testimonials when you have them.
- **Roadmap targets (slide 13):** 75 → 1,000 paying subscribers are targets, not forecasts. Change them if your plan differs.
- **Use of funds (slide 12):** the 40/30/10/10/10 split is a proposal.
- **Team (slide 14):** add your background (career, ministry and sales experience).

## Rebuilding

```bash
npm install pptxgenjs react react-dom react-icons sharp
node build.js Cedarworks-Apps-Pitch-Deck.pptx
```
