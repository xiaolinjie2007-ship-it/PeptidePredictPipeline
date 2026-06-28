# PeptidePredictPipeline

Automated bioactive peptide prediction pipeline integrating enzymatic digestion, antioxidant activity prediction, allergenicity prediction, toxicity prediction, BBB penetration prediction, and bioactivity scoring.

> 多肽生物活性预测自动化流水线。整合多种网络预测工具，实现从蛋白质酶解到活性/毒性/致敏性/抗氧化/血脑屏障穿透的全流程预测。

---

## Project Structure

```
PeptidePredictPipeline/
│
├── digestion/                    # Module 1: Enzymatic Digestion
│   ├── run_digestion.py          # BIOPEP online digestion (25 enzyme schemes)
│   ├── merge_xlsx.py             # Merge detail + summary Excel
│   ├── check_xlsx.py             # Check digestion output structure
│   ├── check_merged.py           # Check merged file content
│   └── inspect_excel.py          # Check input protein file format
│
├── prediction/                   # Module 2: Multi-step Prediction Pipeline
│   ├── pipeline.py               # [Main Entry] 4-step combined prediction
│   ├── antioxidant.py            # AnOxPePred antioxidant (FRS + CHEL)
│   ├── antioxidant_zisu.py       # Perilla peptide antioxidant (dedicated)
│   ├── allergen.py               # AllerTOP allergenicity prediction
│   ├── allergen_resume.py        # Resume interrupted allergenicity
│   ├── toxicity.py               # ToxinPred3 single toxicity prediction
│   ├── toxicity_batch.py         # ToxinPred3 batch toxicity prediction
│   ├── toxicity_verify.py        # Toxicity prediction verification
│   ├── bbb.py                    # B3Pred BBB batch prediction
│   └── bbb_single.py             # B3Pred BBB single prediction (recommended)
│
├── bioactivity/                  # Module 3: Bioactivity Scoring
│   ├── __init__.py
│   └── ranker_pipeline.py        # PeptideRanker auto-submission pipeline
│
├── data_crawl/                   # Module 4: Data Collection
│   ├── __init__.py
│   └── fetch_satpdb_cpp.py       # Crawl SATPdb cell-penetrating peptides
│
├── data/                         # Input data (Excel / CSV)
│   └── samples/                  # Sample data
│
├── outputs/                      # Prediction output directory
├── config.example.py             # Configuration template
├── requirements.txt              # Python dependencies
├── LICENSE                       # MIT License
├── README.md                     # This file
└── .gitignore                    # Git ignore rules
```

---

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Full Pipeline (4-Step Combined)

```bash
cd prediction
python pipeline.py
```

Automatically executes: Activity screening → Antioxidant prediction → Allergenicity prediction → Toxicity prediction

### 3. Individual Digestion

```bash
cd digestion
python run_digestion.py
```

### 4. Individual Predictions

```bash
# Antioxidant
cd prediction
python antioxidant.py

# Allergenicity
python allergen.py

# Batch toxicity
python toxicity_batch.py

# BBB penetration (recommended: single mode)
python bbb_single.py input.csv
```

---

## Module Details

### Module 1: Digestion (`digestion/`)

| Script | Function | Input | Output |
|--------|----------|-------|--------|
| `run_digestion.py` | BIOPEP digestion with 25 enzyme schemes | Protein sequence Excel | `digest_detail_*.xlsx` + `digest_summary_*.xlsx` |
| `merge_xlsx.py` | Merge detail and summary into one Excel | Two xlsx files | `enzyme_digestion_results.xlsx` |
| `check_xlsx.py` | Debug: check digestion output structure | `digest_detail_*.xlsx` | Console output |
| `check_merged.py` | Debug: check merged result | `enzyme_digestion_results.xlsx` | Console output |
| `inspect_excel.py` | Debug: check input file format | Protein sequence Excel | Console output |

### Module 2: Prediction (`prediction/`)

**Combined Pipeline** `pipeline.py` runs 4 steps sequentially:

```
Step 1: Activity screening (bioscore >= 0.5)
    → active_ge_05.xlsx
Step 2: AnOxPePred antioxidant prediction
    → antioxidant_pred.xlsx + antioxidant_ge_05.xlsx
Step 3: AllerTOP allergenicity prediction (browser automation)
    → allergen_pred.xlsx + non_allergen.xlsx
Step 4: ToxinPred3 toxicity prediction
    → toxicity_pred.xlsx + non_toxic.xlsx
```

| Script | Function | Usage |
|--------|----------|-------|
| `pipeline.py` | 4-step combined entry | `python pipeline.py` |
| `antioxidant.py` | DTU AnOxPePred antioxidant (FRS + CHEL) | `python antioxidant.py` |
| `antioxidant_zisu.py` | Perilla peptide dedicated antioxidant | `python antioxidant_zisu.py` |
| `allergen.py` | AllerTOP browser automation | `python allergen.py` |
| `allergen_resume.py` | Resume interrupted prediction | `python allergen_resume.py` |
| `toxicity.py` | Single peptide toxicity (CLI) | `python toxicity.py AAAA...` |
| `toxicity_batch.py` | Batch toxicity from Excel | `python toxicity_batch.py` |
| `toxicity_verify.py` | Verify results vs cache | `python toxicity_verify.py` |
| `bbb.py` | B3Pred BBB batch (CSV) | `python bbb.py` |
| `bbb_single.py` | B3Pred single (recommended) | `python bbb_single.py input.csv` |

### Module 3: Bioactivity (`bioactivity/`)

| Script | Function | Description |
|--------|----------|-------------|
| `ranker_pipeline.py` | PeptideRanker bioactivity scoring | Browser submit + QQ mail receive, resume supported |

### Module 4: Data Collection (`data_crawl/`)

| Script | Function | Description |
|--------|----------|-------------|
| `fetch_satpdb_cpp.py` | Crawl SATPdb CPP database | 17 pages, 831 records, CSV + Excel |

---

## Prediction Tools Overview

| Type | Tool | Script |
|------|------|--------|
| 🧬 Bioactivity (0-1) | PeptideRanker | `ranker_pipeline.py` |
| 🛡️ Antioxidant (FRS/CHEL) | AnOxPePred (DTU) | `antioxidant.py` |
| 🤧 Allergenicity | AllerTOP | `allergen.py` |
| ☠️ Toxicity | ToxinPred3 | `toxicity_batch.py` |
| 🧠 BBB Penetration | B3Pred | `bbb_single.py` |
| 🔪 Digestion | BIOPEP | `run_digestion.py` |

---

## Dependencies

```txt
openpyxl>=3.1.0
requests>=2.31.0
nodriver>=0.38.0
beautifulsoup4>=4.12.0
lxml>=5.0.0
Pillow>=10.0.0
```

## License

MIT License
