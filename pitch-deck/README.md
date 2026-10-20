# Cedarworks Apps: Investor Pitch Deck

A 27-slide pre-seed deck ($450K) for investor programs such as Initio Capital.

| File | Use |
|------|-----|
| `Cedarworks-Apps-Pitch-Deck.pdf` | Upload this to investor portals (e.g. Initio's "Upload deck to autofill" and data room) and send it to investors. |
| `Cedarworks-Apps-Pitch-Deck.pptx` | Editable PowerPoint version, with speaker notes on every slide. |
| `initio-profile-answers.md` | Ready-to-paste answers for the Initio Company Profile form. |
| `build.js` | Script that generates the `.pptx` (pptxgenjs). |
| `*-wide.png`, `appforge-*.png` | Product screens used on the slides, plus the `.html` files they're rendered from. |

## Slides

1. Title · 2. Executive Summary · 3. Problem Statement · 4. Solution overview ·
5–6. PastorAI (features, inside the app) · 7–8. Ready Room (features, inside
the app) · 9. AppForge overview · 10. AppForge step 1: Know (Brain, Chief of
Staff) · 11. Step 2: Find (Top 100 Apps, Viability Scorer) · 12. Step 3: Build ·
13. Step 4: Market (funnel tools) · 14. Why AppForge matters · 15. Traction &
Metrics · 16. Market Opportunity · 17. Why Now & Customer Segments ·
18. Business Model · 19. Unit Economics · 20. Go-to-Market · 21. Competitive
Landscape · 22. Use of Funds · 23. Roadmap · 24. Financial Projections ·
25. Team & Founders · 26. The Ask · 27. Key Terms

The small orange label on each slide uses the standard category name (e.g.
"Problem Statement", "Traction & Metrics") so AI deck scorers such as
Initio's can match every section. The smallest body text is about 12pt, and
business terms are explained on the slides and in the Key Terms appendix.

## Update before sending

- **Traction (slide 15):** pre-revenue, ~30 trial users. Update with paying subscribers and MRR as they come in.
- **Unit economics (slide 19):** AI/hosting cost per user and 5% monthly churn are estimates. Replace them with real numbers once you have them.
- **Competitors (slide 21):** check the named examples and how they're described.
- **Use of funds (slide 22) and financial projections (slide 24):** a planning model. Adjust it if your plan changes.
- **Product screens (slides 6, 8, 10, 11):** recreations of the live apps, made from the founder's screenshots with personal emails removed. To use real screenshots, save them over the PNGs and rebuild.

## Rebuilding

```bash
npm install pptxgenjs react react-dom react-icons sharp
node build.js Cedarworks-Apps-Pitch-Deck.pptx
```
