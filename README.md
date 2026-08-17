# data/

Source files only. Nothing in here is ever edited — see rule 1 in CLAUDE.md.

## abs/
ABS Input-Output Tables, as downloaded from
https://www.abs.gov.au/statistics/economy/national-accounts/australian-national-accounts-input-output-tables

Needed: Table 5, Table 8, Tables 23-34 (margins), Table 35 (net taxes).
Keep the ABS filenames (520905500105.xlsx etc.) so the loader can identify them,
or rename to "Table 5.xlsx" style.

Optional: the single combined workbook 5209055001DO001_202324.xlsx contains
every table. Also worth having: the Industry and Product Concordance Tables,
from the Methodology page rather than Data downloads — they map household
expenditure and capex categories to IOIG and are what bridging matrices are
built from.

## supplied/
Your provider's files: Table 5 and Table 8 for all nine regions, and the
multiplier set for all nine regions. Record where each came from and its
vintage in CLAUDE.md when you add them.
