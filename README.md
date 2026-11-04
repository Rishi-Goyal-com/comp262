# COMP 262 Project, Phase 1: Lexicon Sentiment Analysis (Team 2)

Sentiment analysis of Amazon **All Beauty** reviews with two lexicon models, **VADER** and **TextBlob**.

Team 2: Manav Chaudhary, Harsh Eary, Steven Hao Gong, Rishi Goyal, Parth Kishorbhai Kevadiya, Jasdish Singh

## Requirements
- Python 3.10 or newer
- Internet connection on the first run (downloads the dataset and NLTK word lists)

## Setup (one time)
Open a terminal in this folder and install the packages:

```bash
pip install pandas numpy matplotlib seaborn nltk scikit-learn vaderSentiment textblob wordcloud
```

Optional: use a virtual environment so the packages stay separate.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas numpy matplotlib seaborn nltk scikit-learn vaderSentiment textblob wordcloud
```

On macOS or Linux, activate with `source .venv/bin/activate` instead.

## Run

```bash
python team2_phase1_lexicon_sentiment.py
```

The script takes under a minute. It uses a fixed random seed (42), so the numbers are the same on every run and match the report.

## What the script does
| Step | What happens |
|---|---|
| Load data | Downloads `All_Beauty_5.json.gz` into `data/` if missing |
| 1. Exploration | Counts, distributions, review lengths, duplicates |
| 2. Pre-processing | Labels from ratings, column choice, outliers |
| 3. Lexicon study | Compares VADER, TextBlob and SentiWordNet on test sentences |
| 4. Model cleaning | Different pre-processing for each model, tested on held-out reviews |
| 5. Sample | Picks 1,000 random reviews for testing |
| 6. Models | Scores the sample with VADER and TextBlob |
| 7. Validation | Accuracy, precision, recall, F1, confusion matrices |

## Outputs
- `figures/`: 16 charts (PNG)
- `results/`: result tables (CSV)
- `results/09_comparison_table.csv`: main results table
- `results/run_log.txt`: full copy of the console output

## Expected results
| Model | Accuracy | Macro-F1 |
|---|---|---|
| VADER | 0.891 | 0.469 |
| TextBlob | 0.866 | 0.404 |

## Troubleshooting
- **`ModuleNotFoundError`**: run the `pip install` command again in the same environment you use to run the script.
- **Download fails**: check the internet connection, or place `All_Beauty_5.json.gz` in the `data/` folder manually. Source: https://nijianmo.github.io/amazon/index.html
- **Unicode error in the console on Windows**: run `set PYTHONIOENCODING=utf-8` first, then run the script again.

## Data source
Ni, J., Li, J., & McAuley, J. (2019). Justifying recommendations using distantly-labeled reviews and fine-grained aspects. EMNLP-IJCNLP 2019.
