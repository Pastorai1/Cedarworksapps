# Cedarworks Apps: Investor Pitch Deck

An 18-slide main deck plus an 11-slide appendix: a $450K pre-seed pitch for investor programs such as Initio Capital.

| File | Use |
|------|-----|
| `Cedarworks-Apps-Pitch-Deck.pdf` | Upload this to investor portals (e.g. Initio's "Upload deck to autofill" and data room) and send it to investors. |
| `Cedarworks-Apps-Pitch-Deck.pptx` | Editable PowerPoint version, with speaker notes on every slide. |
| `initio-profile-answers.md` | Ready-to-paste answers for the Initio Company Profile form. |
| `build.js` | Script that generates the `.pptx` (pptxgenjs). |
| `*-wide.png`, `appforge-*.png` | Product screens used on the slides, plus the `.html` files they're rendered from. |

## Slides

**Main deck (1–18):** 1. Title · 2. Executive Summary · 3. Problem Statement ·
4. Solution overview · 5. PastorAI · 6. Ready Room · 7. AppForge ·
8. Traction & Metrics · 9. Focus plan (prove two apps, then unlock apps 3–10) ·
10. Market Opportunity · 11. Competition: PastorAI vs. faith-tech ·
12. Competition: Ready Room vs. AI sales coaching · 13. Business Model ·
14. Go-to-Market · 15. Team & Founders · 16. Financial Projections (3-year) ·
17. Use of Funds & Milestones · 18. The Ask

**Appendix (19–29):** divider · PastorAI screen · Ready Room screen · AppForge
steps 1–4 (Know, Find, Build, Market) · Why AppForge matters · Why now &
customer segments · Unit economics · Key terms

Every scored category sits in the first 18 slides, because Initio's scorer
appears to stop reading after about 18–20 slides. The small orange label on
each slide uses the standard category name (e.g. "Problem Statement",
"Traction & Metrics") so AI deck scorers can match every section.

## Update before sending

- **Traction (slide 8):** pre-revenue, ~30 trial users. Update with paying subscribers and MRR as they come in.
- **Unit economics (slide 28):** AI/hosting cost per user and 5% monthly churn are estimates. Replace them with real numbers once you have them.
- **Competitors (slides 11–12):** check the named competitors (Logos, Pulpit AI, Gloo, Second Nature, Hyperbound, Gong) and how they're described.
- **Financial projections (slide 16) and use of funds (slide 17):** a planning model. Adjust it if your plan changes.
- **Product screens (slides 20–23):** recreations of the live apps, made from the founder's screenshots with personal emails removed. To use real screenshots, save them over the PNGs and rebuild.

## Rebuilding

```bash
npm install pptxgenjs react react-dom react-icons sharp
node build.js Cedarworks-Apps-Pitch-Deck.pptx
```

The committed deck lists only the office phone. To add the founder's cell number
for a copy you send directly to investors (kept out of this public repo):

```bash
CELL_PHONE=xxx-xxx-xxxx node build.js Cedarworks-Apps-Pitch-Deck.pptx
```
