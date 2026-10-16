# India EV Two-Wheeler Aftermarket: Market Sizing

A consulting-style market-sizing case: how big is the servicing and battery-replacement market for India's
electric two-wheelers (E2Ws), how fast will it grow to FY31, and where could an independent service business play?

**Deliverables**
- **Model:** [`model/ev2w_service_market.xlsx`](model/ev2w_service_market.xlsx), formula-driven, with sourced inputs, labelled assumptions, three scenarios and a sensitivity grid.
- **Deck:** [`deck/ev2w_aftermarket_deck.pdf`](deck/ev2w_aftermarket_deck.pdf) (PowerPoint: [`.pptx`](deck/ev2w_aftermarket_deck.pptx)), 12 slides with the answer first.
- **Memo:** [`reports/case_memo.md`](reports/case_memo.md), a one-page summary.

## Key results
_Fill in after reviewing: FY26 and FY31 size, scenario range, what drives it, where to play._

## Approach
Bottom-up: E2W sales by year (Vahan-based sources) → fleet in use after scrappage → vehicles past their free first year ×
annual service and wear-parts spend, plus battery replacements as each year's sales reach replacement age. Spend per vehicle
is cross-checked against reported e-scooter maintenance costs (about half a petrol scooter's).

## Model structure
| Sheet | Contents |
| --- | --- |
| Cover | Question, answer (live formulas), colour code, sources |
| Inputs | Sourced sales data (blue) and assumptions (blue on yellow), by scenario |
| Model | Sales, fleet, paying fleet, service, battery and total market by year for bear, base and bull |
| Summary | Key outputs and two charts |
| Sensitivity | FY31 market across sales growth (10-30%) and spend per vehicle (Rs 2,000-4,000) |

## Quality checks
`src/market_model.py` re-implements the model in Python, and `tests/test_model.py` checks the workbook's calculated
values against it (every year and scenario), so a broken Excel formula fails the GitHub Actions run.
`src/build_model.py` rebuilds the workbook from code; `src/export_data.py` and `deck/build_deck.js` rebuild the deck.

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
pytest -q
```

## Sources
Argus Media and PIB (Vahan data, FY22-FY23), Ministry of Heavy Industries via Times Drive (FY24-FY25), IESA and
JMK Research via EV Infrastructure News (FY26), EVreporter (CY2025), MeraEV (maintenance costs), e-vehicleinfo and
BikeLeague (battery costs), IESA via IANS (state shares). Full list on the Cover sheet and the last slide.
