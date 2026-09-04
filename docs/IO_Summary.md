<!-- Title block (renders as the header table in the Word template)
     Title:    Input–output modelling overview
     Subtitle: Short internal note on our input–output modelling capability
     Date:     4 September 2026 -->

# Input–output modelling

*Introduction.* Input–output (IO) modelling measures the economic footprint of a project,
industry or organisation—the activity, value added and employment associated with it, including
through its supply chain and the spending of its workforce. It is the natural companion to our
CGE modelling: where Tasman Global answers what difference a change makes to the economy, IO
answers how much economic activity is connected to it. This brief explains what IO modelling is,
how to choose between IO and CGE, what each set of results actually represents, and what our IO
capability offers.

---

## What input–output modelling is

An input–output table is a picture of an economy in a single year. It records what every industry
buys from every other industry, what it pays in wages, profits and taxes, and where its output
goes—to other industries, to households, to government, to investment or to export. Every row and
column must balance, so the table is a complete and internally consistent account of production
and spending.

Because the table shows how much of each input an industry needs to produce a dollar of output, it
can be used to trace a chain of demand. If a project buys construction services, the construction
industry must buy steel, transport, professional services and energy; those industries in turn buy
from their own suppliers; and the workers employed along the chain spend part of their wages in
the wider economy. Adding those rounds together gives a **multiplier**—the total activity
associated with each dollar of initial expenditure.

Results are reported in three layers:

- **Direct effect.** The project's own expenditure, value added and employment.
- **Production-induced effect.** Activity in the supply chain—the industries that supply the
  project, the industries that supply them, and so on. Also described as indirect or
  supply-chain effects. Where it is useful, this layer can be split further into the *first-round*
  effect (the project's immediate suppliers) and the *industrial support* effect (everything
  beyond them).
- **Consumption-induced effect.** Activity generated when employees—of the project and of its
  supply chain—spend their wages in the wider economy. Also described as induced effects.

Some studies and guidelines use the labels *Type I* and *Type II* multipliers. Type I covers the
direct and production-induced effects; Type II adds the consumption-induced effects. We hold the
Type 1A, 1B, 2A and 2B forms and report the three layers separately, so results can be presented
on whichever basis the client's audience expects.

**What it is not.** An input–output model assumes that industries use inputs in fixed proportions,
that prices and wages do not change, and that the labour, capital and materials required are
available without being drawn away from other uses. It has no time dimension, no capacity limits
and no financing constraint. These assumptions are what make it fast, transparent and highly
detailed by industry—and they are also why its results describe the scale of activity associated
with an entity rather than the net gain to the economy from it. The ABS puts it plainly: input–
output multipliers assume no supply constraints, fixed prices and fixed input ratios, and are
likely to significantly overstate impacts. We quote that caveat rather than paraphrase it.

---

## Model overview

*Table 1  An overview of our input–output model*

| | Our input–output model |
|---|---|
| What it is | A generated, fully auditable Excel model that converts a client's spending profile at purchasers' prices into direct, flow-on and total impacts |
| Geography | Australia and all eight states and territories. Every line of expenditure carries its own region, so a study spanning jurisdictions is modelled in one run. Sub-state and local-area tables can be built where a study needs one |
| Industries | 114 industries on the ABS IOIG(2022) spine, reported individually or grouped to the 19 ANZSIC divisions. The ABS spine carries 115; ownership of dwellings is held as a single industry rather than split between actual and imputed rent |
| Measures | Output; wages and salaries; value added at basic prices; gross operating surplus and mixed income; employment in full-time equivalents |
| Effect layers | Direct (initial), production-induced and consumption-induced, reported separately and as a total. Production-induced splits into first-round and industrial-support effects |
| Multiplier types | Type 1A, 1B, 2A and 2B |
| Price basis | Spending is entered at purchasers' prices and converted to domestic basic prices: twelve margin types and net taxes on products are stripped and reallocated separately |
| Vintage | Flow tables and multipliers 2022-23; ABS margin and tax matrices 2023-24 |
| Time | A spending profile of any number of years, with each year reported separately. The economy itself does not respond over time |
| Assurance | A battery of automated checks runs on every build and blocks a run that fails one. The assumptions and caveats that must accompany the results are generated alongside them |

*Source: ACIL Allen*

---

## How it works

- **Start from an input–output table.** A table for the relevant economy and year, showing the
  purchases and sales of every industry and the composition of final demand.
- **Convert the table into multipliers.** The pattern of purchases is used to calculate how much
  total production, value added, income and employment each dollar of demand in each industry
  ultimately requires, once the supply chain and household spending are followed through.
- **Convert the client's expenditure into demand on industries.** Spending is recorded as the
  client incurs it—at the price actually paid—and then stripped to the domestic production it
  represents. This step is set out below.
- **Apply the multipliers.** Each dollar of resulting demand is multiplied by the coefficients for
  the industry that supplies it, in the region where the supply occurs.
- **Report the result by layer, industry and region.** Direct, production-induced and
  consumption-induced effects are reported for output, wages, value added and employment, broken
  down by industry and by the region in which the activity occurs.

### Converting spending into demand on industries

A dollar paid by a purchaser is not a dollar of demand for the industry that made the product. The
price includes wholesale and retail margins, the freight to move the goods and the taxes levied on
them, and part of what remains buys an import. Each of those goes to a different place in the
economy and some of it leaves the model entirely. This step is the most common reason two IO
studies of the same project produce different answers, and it is where most of our modelling
effort sits.

For every line of expenditure we:

- **Remove the taxes on products.** GST, duty and other product taxes are transfers, not
  production, and are excluded from the amount that is multiplied.
- **Separate the margins and reallocate them to the industries that earn them.** Twelve margin
  types are identified individually—wholesale, retail, restaurants/hotels/clubs, road, rail,
  pipeline, water, air, port handling, marine insurance, gas and electricity—and each is assigned
  to the industry that earned it rather than the industry that made the product. A retail purchase
  generates demand for retail trade and for transport as well as for the manufacturer.
- **Split what remains between domestic and imported supply**, using the ratio for that product in
  the region where it is bought. Imported content is excluded, because the production it supports
  happens overseas.

The rates are taken from the ABS supply and use tables and reconciled to the ABS published control
totals before use. They also depend on *who* is buying, not only what: nationally, around
$140 billion of retail margin sits against household consumption and under $7 billion against all
115 industries combined. We strip each line using the column the purchaser actually occupies.

---

## Choosing between CGE and input–output

The two approaches answer different questions, and the choice follows from which question the
client is actually asking.

**Input–output answers:** *how much economic activity is associated with this?* It suits footprint
and contribution studies—an industry's or a facility's presence in a region, the reach of its
supply chain, the jobs it supports.

**CGE answers:** *what difference does this make to the economy?* It suits questions where
resources are scarce, prices respond, financing matters, or the answer changes over time.

*Table 2  Comparing input–output and CGE modelling*

| | Input–output | CGE (Tasman Global) |
|---|---|---|
| What it measures | Gross activity associated with an entity or expenditure | Net change relative to a baseline of what would otherwise happen |
| Resources | Assumed available; no competition for workers or capital | Scarce and contested; crowding out is calculated |
| Prices and wages | Fixed | Adjust until markets balance |
| Substitution | None—input proportions are fixed | Firms and households substitute as relative prices change |
| Time | Expenditure can be phased year by year and reported that way, but the economy does not respond over time—nothing accumulates and nothing lags | Year by year, with investment, debt and population accumulating |
| Financing | Not represented | Government budgets, savings, borrowing and ownership are tracked |
| Geography | Australia and the eight states and territories, plus sub-state regions where built | 153 regions, routinely split to sub-state and local areas |
| Industry detail | Very fine—114 industries, grouped to 19 ANZSIC divisions for reporting | Coarser, up to 76 industries, set by the study aggregation |
| Effort | Days | Weeks |
| Typical use | Economic footprint, contribution and supply-chain studies | Policy change, large projects, net benefit, distributional and dynamic questions |

*Source: ACIL Allen*

**How to decide.** The following usually settle it:

- **Size relative to the economy.** If the change is large enough that it would compete for
  workers, land or materials—particularly in a small region or a tight labour market—IO will
  overstate it and CGE is the appropriate tool.
- **The question behind the question.** "How important is this industry to our region?" is a
  footprint question. "Should the government fund this?" or "what happens if this closes?" is a
  net-impact question.
- **Whether prices, timing or financing are part of the story.** If the change works through
  energy prices, wages, the exchange rate or taxes; if the client needs to know when effects
  arrive and how long they last; or if the money must come from taxes, borrowing or foregone
  spending elsewhere, IO cannot represent it and will count a benefit without its cost.
- **Industry granularity and budget.** If the value is in fine industry detail delivered quickly,
  IO does that well and CGE does not.

**Used together.** The two are complementary and are often best presented side by side: IO to
describe the footprint and supply-chain reach that stakeholders recognise, CGE to establish the
net economic effect that withstands scrutiny. Presented this way, each is doing the job it is
suited to—provided the two sets of numbers are clearly labelled and never added together.

---

## What the results actually represent

The most common error in impact work is treating IO and CGE results as though they measure the
same thing. They do not, and the difference explains why IO numbers are almost always larger.

**Input–output results describe activity supported.** They answer: if this expenditure occurs,
and the resources it needs are available, how much production, value added and employment is
associated with it across the economy? They are a **gross** measure. They do not net off the
activity that those workers and materials would otherwise have been engaged in, so they are not a
measure of what the economy would lose without the project, nor of net benefit.

The measures are worth being precise about:

- **Output** is gross turnover summed along the supply chain, so the same value is counted at each
  stage. It is the largest number available and the least meaningful as a measure of economic
  contribution. It is best used to describe the scale of the supply chain, not the value created.
- **Value added (GVA)** is compensation of employees, gross operating surplus and mixed income,
  and other taxes less subsidies on production. This is value added at **basic prices**—the ABS
  headline definition and ours. It excludes taxes on products such as GST, which belong to value
  added at market prices, and it is the appropriate headline measure because it does not double
  count.
- **Wages and salaries** is the labour income component, reported separately. It is what drives
  the consumption-induced layer, so reporting it makes that layer auditable rather than assumed.
- **Gross operating surplus and mixed income** is the returns-to-capital component, useful where
  the question concerns profits or the ownership of the activity.
- **Employment** is jobs supported, reported in full-time equivalents. These are jobs associated
  with the activity, not necessarily new jobs created.

**CGE results describe a difference.** They are deviations from a baseline: the change in GDP,
national income or employment relative to what would have happened anyway, after resources have
been reallocated and prices have adjusted. They are a **net** measure, they can be negative, and
they are usually much smaller than the corresponding IO figures.

This distinction matters commercially. Multiplier-based results have been criticised—including by
the ABS and the Productivity Commission—when they are presented as net economic benefits rather
than as a description of activity. Describing IO results accurately is not a weakness in a
proposal; it is what makes them defensible when a reviewer or a competing consultant tests them.

---

## Our input–output capability

Applying published national multipliers to a regional question overstates the regional result. A
region buys far more of its inputs from outside its own borders than the nation does, so much of
the flow-on activity occurs elsewhere—and a national multiplier cannot show where. It is also a
long time since the ABS published multipliers: the last national set was for 1998-99, and
the ABS does not publish state or territory input–output tables at all. Working from regionalised
tables of a current vintage is what allows the results below.

- **Current, regionalised tables, independently reconciled.** We work from regionalised tables and
  multipliers for Australia and all eight states and territories, benchmarked to ABS data, on a
  2022-23 vintage rather than the vintage of the last published ABS multiplier set. We do not take
  them on trust: the output multipliers are re-derived from the underlying flow tables and
  compared, and on the current vintage they match for all 114 industries in all nine jurisdictions.
  We record the provenance and vintage of every input.
- **Any region, including below state level.** Tables can be estimated for smaller regions within
  states, so a client's region is modelled directly rather than inferred from a state average.
- **National and state results from one run, with the leakage visible.** Every line of expenditure
  carries its own region, so a study spanning several jurisdictions is modelled in a single pass,
  each line resolved against the tables of the state in which the spending occurs. The same
  expenditure can also be run against the national tables. Reporting both is how the leakage is
  made explicit: the national result is larger than the sum of the state results, and the
  difference is flow-on activity that leaves each state and lands elsewhere in Australia. State
  results are not additive to a national total, and we label them so they are never summed.
- **Flow-on effects allocated to industries.** We hold the industry-by-industry matrices behind the
  multipliers, not just the headline coefficients, so flow-on activity is reported as a genuine
  breakdown—which industries the supply-chain dollars land in—for all 114 industries or grouped to
  the 19 ANZSIC divisions. The check that makes this legitimate is that the industry split sums
  back to the same multiplier the headline result uses; that reconciliation is re-run on every data
  drop and currently holds for every industry in all nine jurisdictions.
- **The full effect decomposition.** Because we hold the components rather than a single headline
  number, results can be presented as direct, first-round, industrial-support and
  consumption-induced effects, on a Type I or Type II basis, for any of the five measures—without
  re-running the study when an audience wants a different cut.
- **An auditable model, with gates that block a bad run.** The workbook is generated by script and
  never hand-edited, so every result can be traced to the code and the source data that produced
  it. Source data is held verbatim and proven cell-for-cell against the original files; every
  number downstream of it is a visible formula rather than a pasted value. A set of automated
  checks tests the things that go wrong silently—expenditure landing on a product that does not
  exist in the region's table, margins lost or double counted between steps, shares that fail to
  sum to one—and reports a failure rather than a plausible-looking number.

Our input–output and CGE capabilities sit with the same team, so a client can move from a
footprint study to a net-impact study—or commission both—with a consistent treatment of the
industries and regions involved.

### What we disclose with every study

Stating the limits is part of the product, and each of the following is generated with the results
rather than remembered at drafting time.

- **These are gross activity figures, not net benefit.** Repeated wherever a total is presented.
- **State results do not sum to a national result.** Tracing the interstate flows themselves,
  industry by industry, requires a multi-regional model; where a client needs that, we scope it as
  such.
- **Margin and tax rates are national.** The ABS publishes no state margin or tax matrices, so a
  margin is assumed to be earned in the region where the purchase is made. This will overstate
  local wholesale and retail activity where goods are distributed from another state.
- **The vintage is mixed.** Flow tables and multipliers are 2022-23; the ABS margin and tax
  matrices are 2023-24. The gap is disclosed rather than smoothed over.

---

## When to use input–output

*Table 3  Typical input–output applications*

| Type of work | Examples |
|---|---|
| Economic footprint and contribution studies | The contribution of an industry, company, facility, port, university or event to a region, state or the nation |
| Supply-chain and local content analysis | Where a project's expenditure flows, which industries supply it, and how much is retained locally |
| Project expenditure assessment | Construction and operating expenditure translated into activity, value added and employment by industry and region |
| Regional and community impact | The economic role of a major employer in a town or region, and what is associated with its presence |
| Scoping and screening | A fast first estimate to size an impact and decide whether fuller CGE analysis is warranted |
| Reporting and disclosure | Recurring economic contribution reporting for annual reports, submissions and stakeholder communication |

*Source: ACIL Allen*

---

## What clients receive

Clients receive a concise report describing the expenditure or activity assessed, the modelling
approach and assumptions, and the results with their interpretation. Depending on scope, results
can be provided in the following layers:

- **Headline results (core results).** Output, wages and salaries, value added at basic prices,
  gross operating surplus and mixed income, and employment in full-time equivalents—reported as
  direct effects and total effects for the region analysed, year by year where the expenditure is
  phased.
- **Effect layers.** Direct, production-induced and consumption-induced effects reported
  separately, so the client can see how much of the total comes from the project itself, how much
  from its supply chain, and how much from the spending of wages. Results can be presented
  excluding or including the consumption-induced layer, depending on what the client's audience
  expects.
- **Industry results (supplementary spreadsheet).** The full breakdown by industry, including the
  allocation of flow-on effects across industries rather than a single aggregate, at 114-industry
  detail or grouped to the 19 ANZSIC divisions.
- **Regional components (where relevant).** Results for each region in which the client spends,
  and for Australia as a whole, showing how much of the flow-on activity is retained in the
  client's region and how much leaves it.
- **The assumptions and caveats.** Issued with the results, in the form a reviewer will ask for.

---

## What we need from a client

Input–output analysis is usually quick to scope. We need enough detail to allocate expenditure to
the right industries, the right purchasers and the right places.

- **Expenditure**, ideally by category and by year, split between capital and operating spending,
  and recorded at the price actually paid. We do the conversion to basic prices; the client does
  not need to strip anything out first.
- **Where it is spent**—the region or state, and the split between local, interstate and imported
  purchases where known.
- **Who is buying, not only what is bought.** The same product bought by a contractor as an input
  and bought by an agency as a finished asset sits in different columns of the table and carries
  different margins. A construction bill of quantities belongs with business purchases of inputs,
  because the contractor is a business buying materials—not with capital formation, where the
  table records finished assets. Nationally, the entire general government capital column shows
  around $20 million of iron and steel against $35 billion of civil construction, so materials
  entered there return almost nothing. We check this on every line.
- **Anything already itemised elsewhere.** If a supplier's own expenditure has been counted
  separately, we flag that line so it is included as direct activity but not multiplied a second
  time.
- **Direct employment**, by region. We report on a full-time equivalent basis, so headcount
  figures should be identified as such.
- **Wages and salaries**, which drive the consumption-induced effects.
- **The region for which results are required**, agreed at the outset, because it determines which
  tables are used and how much of the flow-on activity is captured.

Where expenditure detail is limited, we can allocate it using industry-typical cost structures,
and we will state that assumption in the reporting.
