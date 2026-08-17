# Australian IO Impact Model — build plan v2 (multipliers as input)

Supersedes the 13 August plan, which assumed multipliers were derived inside the model. Companion files: `IO_Impact_Model_v0-2.xlsx` (working model with the ABS source tables embedded) and `IO_model_reference_lists.xlsx` (spine, margin map, file checklist).

---

## What changed from v1

| | v1 assumption | v2 (this plan) |
|---|---|---|
| Multipliers | Derived in-model from ABS Table 5 / Table 7 | **Supplied as an input.** Pasted into `RAW_Multipliers`, never computed |
| Leontief inverse | Table 7 ÷ 100, held in the workbook | Not needed |
| Type II closure | Sherman-Morrison rank-one update | Not needed — your set already carries consumption-induced and total columns |
| ABS files needed | ~20 tables | **The margin and tax tables you already have, plus nothing else** — all now embedded in the workbook |
| Biggest risk | Getting the matrix algebra right | **Provenance and integrity of the supplied set.** The QA gates now carry the model |
| Regionalisation | Build it (FLQ / RPC) | Already done by your provider — the model's job is to prove it's real |

The engine layer effectively disappears. In exchange, the model has no internal way to tell whether the multipliers pasted into it are correct, complete, or from the right year. That is what the QA tab is for, and it is now the most important part of the build rather than an afterthought.

---

## Two findings from building it

**1. Your multiplier set reconciles exactly to Table 5.** I derived the simple output multiplier from the flow tables in the master model and compared it to the supplied set: all 114 industries within 0.001 for Australia, NSW and SA. Against Table 8 it does not reconcile at all. So your provider used the direct-allocation (domestic) table, which is the correct choice, and the set is internally coherent. That is a good result — it means the data is sound and the problem is purely one of loading and vintage.

The reconciliation needs the right denominator. Your state blocks carry primary inputs as P3a/P3b/P3c/P3d and P4a/P4b where the national block uses P3/P4, so a hardcoded row reference understates output and inflates every derived multiplier. Sum the P-rows dynamically.

**2. A bill of quantities cannot be allocated to a GFCF column.** This one surprised me, and it contradicts Worked Example B in the methodology document.

In the ABS framework, the gross fixed capital formation columns contain *finished capital assets* — construction services, machinery, software. They do not contain steel or concrete. Nationally, the whole General Government GFCF column shows $20m of iron and steel and $17m of cement against $35,115m of heavy and civil engineering construction. A road authority buys a bridge; the contractor buys the steel.

So there are two valid ways to model a capital project, and the model now supports both:

- **Top-down (recommended).** One line: the contract value against 3101 Heavy and civil engineering construction, column group GFCF_GG. The strip uses the real GFCF_GG rates (97.6% domestic for 3101) and the multiplier generates the steel, concrete, plant and design automatically. Fewer assumptions, no double-count risk.
- **Bottom-up.** The contractor's cost plan, allocated to the **Intermediate** column group, because the contractor is a business buying inputs. Use this only when you have a real bill of quantities and the project's cost structure genuinely differs from the industry average. The contractor's own wages and margin line must then be flagged `Direct only`, or the multiplier will rebuild the supply chain you have already itemised.

The first version of the model silently returned zero for the steel and concrete lines, because there was no ABS cell to strip against. There is now a QA gate for exactly that.

---

## The model as built — v0.2

`IO_Impact_Model_v0-2.xlsx`. 25 tabs, 28,406 formulas, recalculates clean with zero errors. Every number in the calculation chain is a formula reading a source table — nothing downstream of the RAW tabs is a pasted result.

**The chain, in order.** Each tab reads only the ones above it.

| # | Tab | What it holds |
|---|---|---|
| — | `README` | Versions, the calculation chain, the placeholder register, control totals, how to run |
| — | `Settings` | Region, start year, study type, headline basis, price years, deflator |
| — | `Lists` | The 115-code IOIG(2022) spine, the bridge to your 114-code set, column groups, margin destinations with their ABS totals |
| 1 | `RAW_T5_Aus` | **ABS Table 5, 2023-24, full 115 × 122.** Domestic production only |
| 1 | `RAW_T8_Aus` | **ABS Table 8, 2023-24, full 115 × 122.** Imports embedded |
| 2 | `RAW_T5_State` `RAW_T8_State` | **PLACEHOLDER** state flow tables, red-bannered |
| 3 | `RAW_T23_Wholesale` `RAW_T24_Retail` `RAW_T35_NetTaxes` | **ABS Tables 23, 24 and 35 in full** |
| 3 | `RAW_T25_34_Margins` | ABS Tables 25–34 stacked, one 115-row block per margin type |
| 4 | `CALC_StripData` | Aggregates the RAW tabs to 8 purchasing-column groups. Every cell a formula |
| 5 | `CALC_PP` | Purchasers price = the SUMIF of all fourteen blocks by column group |
| 6 | `CALC_Rates` | The shares, shown as percentages, with a Check column that must read 100% |
| 7 | `RAW_Multipliers` | **Paste target.** Aus, NSW, Vic, SA loaded; the other five deliberately empty |
| 8 | `ENG_SelectedMult` | The selected region's multipliers, 5 measures × 4 effects |
| 9 | `IN_Shock` | 40 spending lines × 8 years, with the column-group guidance at the top |
| 10 | `CALC_Strip` | Line × year, split into domestic / imports / taxes / 11 margins, with a per-row check |
| 11 | `CALC_Margins` `CALC_Vector` | Margins reallocated to earning industries; the direct domestic vector |
| 12 | `CALC_Impacts` `OUT_Summary` `OUT_Detail` | Vector × multipliers, and the results |
| 13 | `QA_Checks` `Assumptions` | 20 gates with one overall status; the register to attach to reports |

**How the strip data is derived, in the workbook's own formulas:**

```
Domestic  =IF(Settings!$B$5="Aus",SUM(RAW_T5_Aus!$C6:$DM6),SUM(RAW_T5_State!$C6:$DM6))
Imports   =IF(Settings!$B$5="Aus",SUM(RAW_T8_Aus!$C6:$DM6)-SUM(RAW_T5_Aus!$C6:$DM6),
                                  SUM(RAW_T8_State!$C6:$DM6)-SUM(RAW_T5_State!$C6:$DM6))
NetTaxes  =SUM(RAW_T35_NetTaxes!$C6:$DM6)
Margin_3301 =SUM(RAW_T23_Wholesale!$C6:$DM6)
Margin_4801 =SUM(RAW_T25_34_Margins!$D351:$DN351)+SUM(RAW_T25_34_Margins!$D466:$DN466)
```

The `Intermediate` group sums the 115 industry columns; each Q column is a direct cell reference. The region switch sits only on Domestic and Imports, because the ABS publishes no state margin or tax matrices.

**Three new gates** beyond v0.1: a margin control total (`CALC_StripData` must reproduce the ABS Tables 23–34 total of $421,410m, and does, to the dollar), a placeholder-in-use flag that changes the overall status banner to *RUNNING ON PLACEHOLDER DATA*, and a note that state runs use national margin and tax rates.

**Switching region to NSW** produces exactly what it should: the model moves to the placeholder flow tables, the banner changes, the conversion ratio falls from 96.4% to 93.9% (a state leaks more to imports), and every Type I total drops. The mechanism works; the state numbers are not yet real.

### Placeholder register — everything in the model that is not real data

| What | Placeholder used | How it is flagged | To fix |
|---|---|---|---|
| State Table 5 and Table 8 | NSW 2021-22 block from the master model, expanded to the 115 spine with 6700 = 0. Stands in for **every** non-Australia region | Red banner on both tabs, `Settings!B16`, QA gate 18, `Assumptions` | Load the real nine-region set |
| QLD, WA, Tas, NT, ACT multipliers | **None — left empty on purpose.** The model refuses to produce a result rather than quietly returning national numbers | QA gate 2 fails | Load the real set |
| 6700 Imputed rent | Borrows 6701's multipliers | `Lists` column D highlighted, QA gate 14 | Vendor supplies a 6700 row |
| State margin and tax rates | National ABS rates | QA gate 20 | Nothing to fix — the ABS publishes national only. Disclose it |
| Vintage | Multipliers 2021-22 against ABS 2023-24 strip data | QA gate 19 | Decision needed |
| Export column import share | Zero, because Table 8 and Table 5 treat re-exports differently | QA gate 13 | Separate import-content source |

### The strip engine

For a line of `$PP` on product **p** bought by column group **k**:

```
Domestic  = Table 5 [p,k]                    ' direct allocation — domestic only
Imports   = Table 8 [p,k] - Table 5 [p,k]    ' the competing-import content
NetTaxes  = Table 35 [p,k]
Margin_t  = Tables 23-34 [p,k]               ' collapsed to the 11 earning industries
PP_abs    = Domestic + Imports + NetTaxes + SUM(Margin_t)
```

Every component is then `$PP × component / PP_abs`. The four shares sum to exactly 1, which is QA gate 10.

Deriving imports as Table 8 less Table 5 is what removes the need for ABS Tables 2 and 3, and it means that **once you load state Table 5 and Table 8, state-specific import shares come free**. That is a real methodological gain over using national import rates for a state study.

Rates are held at product × 8 column groups rather than product × 122 individual columns. The variation that matters is between groups — retail margin on food is $140bn against households and under $7bn across all industry columns combined — and the grouped table is 13,000 cells instead of 200,000. Cell-level rates can be added later for a study that needs them.

Sample rates it produces, all straight from the ABS data with no assumptions:

| Product | Bought by | Domestic | Imports | Tax | Retail | Wholesale | Road |
|---|---|---:|---:|---:|---:|---:|---:|
| 1101 Meat | Households | 52.9% | 3.3% | 0.3% | 33.8% | 8.6% | 1.1% |
| 1305 Clothing | Households | 0.8% | 40.0% | 9.0% | 33.6% | 15.9% | 0.2% |
| 1701 Petroleum | Households | 13.5% | 38.6% | 29.2% | 9.1% | 8.6% | 1.0% |
| 2101 Iron and steel | Industry | 43.7% | 37.2% | 0.2% | 0.0% | 14.2% | 4.1% |
| 6901 Professional services | Industry | 93.4% | 6.6% | 0.1% | 0.0% | 0.0% | 0.0% |

Compare the clothing row with the "retail 60%, imports 80%" placeholder in the methodology document, and the petroleum row with the manual fuel excise assumption. The excise falls out of the data.

### The QA gates

Seventeen checks, one overall status. The ones that matter most given multipliers are an input:

| # | Gate | Catches |
|---|---|---|
| 2 | Multipliers loaded for the selected region | Running a study on an empty region |
| 3 | Selected region is not a copy of Australia | The five placeholder jurisdictions |
| 4 | State multipliers smaller than national | Failed or fake regionalisation |
| 5, 6 | Simple = initial + production induced; production induced = first round + industrial support | Column misalignment on paste |
| 7 | Type II not smaller than Type I | Block swap on paste |
| 10 | Components sum back to the purchasers price | A broken rate table |
| 11 | Direct vector reconciles to the strip | Margins lost or double counted |
| 15 | No included line falls on an empty ABS cell | The bill-of-quantities-into-GFCF trap |
| 16 | No negative domestic or import share | ABS cells with net disposals or import adjustments |
| 17 | Shock price year matches multiplier price year | The vintage mismatch |

On the built-in example the model returns 0 fail, 2 review — the two reviews being the 6700 bridge and the vintage mismatch, both of which are real open items rather than bugs.

---

## What I need from you

Short list now, because deriving multipliers is off the table and the Table 5 / Table 8 pair replaces ABS Tables 2 and 3.

**Essential**

1. **The complete multiplier set, all nine regions.** Whatever format it comes in — the master model only has four genuinely distinct sets (Aus, NSW, Vic, SA), and QLD, WA, Tas, NT and ACT are copies of the national rows. I have left those regions empty in `RAW_Multipliers` rather than shipping placeholders that look like data.

2. **The complete Table 5 and Table 8 set, all nine regions.** The master model has three genuinely distinct blocks (Aus, NSW, SA); Vic, QLD, WA, Tas, NT and ACT are byte-identical copies of NSW. I need these for two things: the reconciliation gate, and state-specific import shares in the strip engine.

3. **A decision on vintage.** The supplied multipliers are labelled 2021-22; the strip rates I have built are ABS 2023-24. Either
   - ask your provider for a 2023-24 set (cleanest, and the flow tables would come with it), or
   - tell me to rebuild the strip rates on the 2021-22 ABS release to match, or
   - keep both and use the deflator cell on `Settings`, documented as a known limitation.

**Useful**

4. **The ABS concordance workbook** — `Industry and Product Concordance Tables 2023-24.xlsx`, from the ABS Input-Output tables *Methodology* page, not Data downloads. It contains IOPG↔Household Expenditure Classification, IOPC↔HFCE and IOPC↔CAPEX. These are what turn a client's expenditure categories into IOIG lines, and building bridging matrices by hand without them is guesswork.

5. **Confirmation on employment units** — whether the "Employed" block is persons or FTE, and in what price year the jobs-per-dollar coefficients sit. The model currently labels it "persons".

6. **A 6700 row**, if your provider can supply one. The 2023-24 ABS classification splits ownership of dwellings into 6700 imputed rent and 6701 actual rent; your set has only 6701. The model bridges 6700 to 6701 and flags it, which is fine for most studies but wrong for any housing or household-consumption work.

**Not needed any more** — ABS Tables 2, 3, 4, 6, 7, 20 and the whole derivation apparatus from the v1 plan.

---

## Roadmap to v1.0

| Stage | Work | Done when |
|---|---|---|
| 1 | Load the full nine-region multiplier set | QA gates 2, 3 and 4 pass for every region |
| 2 | Load the nine-region Table 5 and Table 8 | Reconciliation of derived vs supplied first-round effect passes per region |
| 3 | Resolve the vintage decision and rebuild strip rates if needed | QA gate 17 passes |
| 4 | Region-specific import shares in the strip engine | State conversion ratios differ from national, and are lower |
| 5 | Fix the exports column | Table 5 / Table 8 re-export treatment reconciled, or a separate import-content source |
| 6 | Bridging matrix tab driven by the ABS concordances | A client expenditure category list maps to IOIG lines without hand-coding |
| 7 | Contribution-study mode | `Settings` study type actually drives behaviour: measured direct effects override modelled ones |
| 8 | Reporting pack | Charts, a standard results table, and the limitations boilerplate |
| 9 | Template lock-down | RAW tabs protected, per-study save-as workflow, change log discipline |

Stage 1 to 3 are the blocking ones. Everything after that is refinement.

---

## Things to keep from the v1 review

These still stand and are worth folding into the methodology document:

- **Table 5 versus Table 8.** Both are industry-by-industry at basic prices; the difference is import allocation. Table 5 is domestic-only. Your provider used it correctly; the document describes it backwards.
- **Delete the illustrative margin and import rate tables.** The model now computes them from the ABS data.
- **The three ABS claims to remove** — the ABS has not published multipliers since 1998-99, does not publish state tables, and IOIG(2022) has 115 industries not 114.
- **The IOIG codes in the document are wrong in about fifteen places.** The master model's own code list is right; the document is not. Corrected list is in `IO_model_reference_lists.xlsx`.
- **Worked Example B needs rewriting** given the GFCF finding above.
- **GVA has three definitions** — factor cost, basic prices, market prices. The model reports basic prices as the headline, consistent with ABS industry GVA.
- **The transport margin treatment changes in the 2024-25 ABS release** (due March 2027), worth $46.8bn in 2023-24. Margins on agricultural, mining and manufactured products will fall materially. Nothing in the model hardcodes a margin rate, so the update is a re-paste, but expect a genuine break in the series.
- **The ABS view on multipliers** — that they assume no supply constraints, fixed prices and fixed input ratios, and are likely to significantly overstate impacts — belongs in the limitations section of every report, quoted rather than paraphrased.
