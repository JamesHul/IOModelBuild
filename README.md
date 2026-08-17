# io-model

Generator for the Australian Input-Output impact model.

## Folder structure

```
io-model/
├── CLAUDE.md              Project rules and domain knowledge. Claude Code reads this automatically
├── README.md              This file
├── data/
│   ├── abs/               ABS source files, exactly as downloaded. Never edited
│   │   ├── 520905500105.xlsx      Table 5
│   │   ├── 520905500108.xlsx      Table 8
│   │   ├── 520905500123.xlsx      Table 23 wholesale margin
│   │   ├── ...                    Tables 24-34
│   │   └── 520905500135.xlsx      Table 35 net taxes
│   └── supplied/          Your provider's files, exactly as received. Never edited
│       ├── multipliers_*.xlsx     All nine regions
│       └── table5_table8_*.xlsx   All nine regions
├── scripts/
│   ├── inspect_inputs.py  Run first. Reports the shape of every file in data/
│   ├── load_sources.py    The only file that knows your file layouts. Writes build/sources.pkl
│   ├── check_sources.py   Integrity gates on the loaded data. Run before every build
│   └── build_model.py     sources.pkl -> output/IO_Impact_Model_vX.xlsx
├── docs/
│   ├── IO_Model_Build_Plan_v2.md      Methodology, architecture, roadmap
│   └── IO_model_reference_lists.xlsx  IOIG spine, margin map, IOIG code corrections
├── build/                 Intermediate artefacts. Gitignored
└── output/                Generated workbooks
```

## Setup

```bash
cd io-model
git init
python3 -m venv .venv && source .venv/bin/activate
pip install openpyxl numpy
```

Then in Claude Code, install the xlsx skill so it can verify builds:

```
/plugin marketplace add anthropics/skills
/plugin install document-skills@anthropic-agent-skills
```

That gives Claude a `recalc.py` that opens a workbook headless, forces a full recalculation and reports every error cell. A build that has not been recalculated is not verified.

If you want the same check yourself, install LibreOffice. Otherwise open the output in Excel and press Ctrl+Alt+F9.

## Workflow

```bash
python scripts/inspect_inputs.py            # what have we got?
python scripts/load_sources.py              # -> build/sources.pkl
python scripts/check_sources.py             # must pass before building
python scripts/build_model.py               # -> output/IO_Impact_Model_v0-3.xlsx
```

Then recalculate and scan for errors. Commit the script and the workbook together.

## First session in Claude Code

A good opening prompt:

> Read CLAUDE.md. The nine-region multiplier set and Table 5 / Table 8 set are now in data/supplied/.
> Run scripts/inspect_inputs.py, then write the loaders in load_sources.py for the real files.
> Then run check_sources.py and tell me what it finds before we build anything.

Expect the checks to surface something. On the previous data drop they found six of eight state flow-table blocks were identical to NSW and five of eight multiplier sets were identical to Australia.

## The one thing to remember

The workbook is an artefact. The script is the asset. If something needs to change in the model, change the script and regenerate.
