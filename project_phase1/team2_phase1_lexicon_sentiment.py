# -*- coding: utf-8 -*-
"""
COMP 262 - NLP and Recommender Systems
Project Phase 1: Lexicon-based sentiment analysis of Amazon "All Beauty" reviews
Team 2

Dataset
-------
Amazon Review Data (2018), "All Beauty" 5-core subset (All_Beauty_5.json.gz).
Jianmo Ni, Jiacheng Li, Julian McAuley. "Justifying recommendations using
distantly-labeled reviews and fine-grained aspects." EMNLP 2019.
https://nijianmo.github.io/amazon/index.html

What this script does (mapped to the Phase 1 deliverables)
----------------------------------------------------------
 Step 1  Data exploration (counts, averages, distributions, lengths, duplicates)
 Step 2  Labelling (4-5 Positive, 3 Neutral, 1-2 Negative), column selection, outliers
 Step 3  Lexicon study: VADER vs TextBlob vs SentiWordNet on probe sentences
 Step 4  Model-specific pre-processing, with an ablation that justifies each choice
 Step 5  Random sample of 1000 reviews (fixed seed)
 Step 6  Two lexicon models: VADER and TextBlob
 Step 7  Validation: accuracy / precision / recall / F1, confusion matrices,
         comparison table, error analysis

How to run
----------
    pip install pandas numpy matplotlib seaborn nltk scikit-learn vaderSentiment textblob wordcloud
    python team2_phase1_lexicon_sentiment.py

The dataset is downloaded automatically into ./data if it is not already there.
All figures go to ./figures and all tables to ./results.
"""

import json
import os
import re
import html
import random
import string
import urllib.request
from collections import Counter

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # write figures to disk; no window needed
import matplotlib.pyplot as plt
import seaborn as sns
import nltk
from nltk.corpus import stopwords, wordnet as wn, sentiwordnet as swn
from nltk.stem import WordNetLemmatizer
from sklearn.metrics import (accuracy_score, precision_recall_fscore_support,
                             classification_report, confusion_matrix, cohen_kappa_score)
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from textblob import TextBlob
from wordcloud import WordCloud


# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
RANDOM_SEED = 42                      # every random step uses this seed
SAMPLE_SIZE = 1000                    # step 5: number of reviews to model
LABELS = ["Negative", "Neutral", "Positive"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
FIG_DIR = os.path.join(BASE_DIR, "figures")
RES_DIR = os.path.join(BASE_DIR, "results")
DATA_FILE = os.path.join(DATA_DIR, "All_Beauty_5.json.gz")
# the original jmcauley.ucsd.edu link has an expired certificate, so we use the lab's mirror
DATA_URL = ("https://mcauleylab.ucsd.edu/public_datasets/data/amazon_v2/"
            "categoryFilesSmall/All_Beauty_5.json.gz")

for d in (DATA_DIR, FIG_DIR, RES_DIR):
    os.makedirs(d, exist_ok=True)

random.seed(RANDOM_SEED)
np.random.seed(RANDOM_SEED)
sns.set_theme(style="whitegrid", palette="deep")
plt.rcParams.update({"figure.dpi": 110, "savefig.dpi": 150, "savefig.bbox": "tight"})

# colours used for the three sentiment classes in every chart
CLASS_COLORS = {"Negative": "#d1495b", "Neutral": "#edae49", "Positive": "#00798c"}

# NLTK resources needed for SentiWordNet, stop words and lemmatisation
for res in ["stopwords", "wordnet", "omw-1.4", "sentiwordnet",
            "averaged_perceptron_tagger_eng", "punkt_tab"]:
    nltk.download(res, quiet=True)

# everything printed is also collected here and written to results/run_log.txt
_LOG = []


def log(*args):
    """Print to the console and keep a copy for the run log."""
    text = " ".join(str(a) for a in args)
    print(text)
    _LOG.append(text)


def section(title):
    """Print a visible section header."""
    log("\n" + "=" * 78)
    log(title)
    log("=" * 78)


def save_fig(name):
    """Save the current matplotlib figure into ./figures and close it."""
    path = os.path.join(FIG_DIR, name)
    plt.tight_layout()
    plt.savefig(path)
    plt.close()
    log(f"  [figure] figures/{name}")


def save_table(df, name, index=False):
    """Save a DataFrame as CSV into ./results."""
    df.to_csv(os.path.join(RES_DIR, name), index=index)
    log(f"  [table]  results/{name}")


# ----------------------------------------------------------------------------
# Load data
# ----------------------------------------------------------------------------
section("LOAD DATA")
if not os.path.exists(DATA_FILE):
    log("Downloading", DATA_URL)
    urllib.request.urlretrieve(DATA_URL, DATA_FILE)

# one JSON object per line
raw = pd.read_json(DATA_FILE, lines=True)
log(f"Raw shape: {raw.shape[0]:,} rows x {raw.shape[1]} columns")
log("Columns:", list(raw.columns))

# derived columns used throughout the exploration
raw["reviewText"] = raw["reviewText"].fillna("")
raw["summary"] = raw["summary"].fillna("")
raw["review_date"] = pd.to_datetime(raw["unixReviewTime"], unit="s")
raw["year"] = raw["review_date"].dt.year
# "vote" is a string like "1,341"; missing means nobody voted the review helpful
raw["votes"] = pd.to_numeric(raw["vote"].astype(str).str.replace(",", ""), errors="coerce").fillna(0).astype(int)
raw["has_image"] = raw["image"].notna()
raw["n_words"] = raw["reviewText"].str.split().str.len()
raw["n_chars"] = raw["reviewText"].str.len()
raw["summary_words"] = raw["summary"].str.split().str.len()


# ============================================================================
# STEP 1 - DATA EXPLORATION
# ============================================================================
section("STEP 1 - DATA EXPLORATION")

# --- 1.0 Data dictionary: type, missing values and meaning of every column ---
meaning = {
    "overall": "Star rating given by the reviewer (1-5)",
    "verified": "True if Amazon verified the purchase",
    "reviewTime": "Review date as text (MM DD, YYYY)",
    "reviewerID": "Anonymised reviewer id",
    "asin": "Amazon product id (one id per size/colour variant)",
    "style": "Dict of variant attributes (Size, Color, ...)",
    "reviewerName": "Display name of the reviewer",
    "reviewText": "Body of the review",
    "summary": "Review headline / title",
    "unixReviewTime": "Review date as a Unix timestamp",
    "vote": "Number of 'helpful' votes (missing = 0)",
    "image": "URLs of images posted with the review",
}
orig = pd.read_json(DATA_FILE, lines=True)   # untouched copy for true missing counts
dictionary = pd.DataFrame({
    "column": orig.columns,
    "dtype": [str(orig[c].dtype) for c in orig.columns],
    "missing": [int(orig[c].isna().sum()) for c in orig.columns],
    "missing_%": [round(100 * orig[c].isna().mean(), 2) for c in orig.columns],
    "unique": [int(orig[c].astype(str).nunique()) for c in orig.columns],
    "meaning": [meaning.get(c, "") for c in orig.columns],
})
log(dictionary.to_string(index=False))
save_table(dictionary, "01_data_dictionary.csv")

# --- 1a. Counts and averages ---
n_rows = len(raw)
n_products = raw["asin"].nunique()
n_users = raw["reviewerID"].nunique()
log(f"\nReviews: {n_rows:,}   Products: {n_products}   Reviewers: {n_users}")
log(f"Average rating: {raw['overall'].mean():.3f}  (median {raw['overall'].median():.0f}, "
    f"std {raw['overall'].std():.3f})")
log(f"Verified purchases: {raw['verified'].mean():.1%}")
log(f"Reviews with helpful votes: {(raw['votes'] > 0).mean():.1%}   with images: {raw['has_image'].mean():.1%}")
log(f"Date range: {raw['review_date'].min().date()} to {raw['review_date'].max().date()}")
log("Rating counts:", raw["overall"].value_counts().sort_index().to_dict())

fig, ax = plt.subplots(figsize=(7, 4))
counts = raw["overall"].value_counts().sort_index()
bars = ax.bar(counts.index.astype(int).astype(str), counts.values, color="#00798c")
for b, v in zip(bars, counts.values):
    ax.text(b.get_x() + b.get_width() / 2, v, f"{v:,}\n({v / n_rows:.1%})", ha="center", va="bottom", fontsize=9)
ax.set(title="Rating distribution (raw data)", xlabel="Star rating", ylabel="Number of reviews")
ax.set_ylim(0, counts.max() * 1.18)
save_fig("01_rating_distribution.png")

# --- 1g. Duplicates (explored early because it changes every other number) ---
# Amazon shows one review on every size/colour variant of a product, so the same
# review (same reviewer, same text, same timestamp) appears under several asins.
dup_key = ["reviewerID", "reviewText", "unixReviewTime"]
exact_dups = raw.duplicated(subset=[c for c in raw.columns if c not in ("style", "image")]).sum()
key_dups = raw.duplicated(subset=dup_key).sum()
copies = raw.groupby(dup_key)["asin"].nunique()
log(f"\nFully identical rows (all fields except style/image): {exact_dups}")
log(f"Same reviewer + same text + same time (copied across variants): {key_dups:,} "
    f"({key_dups / n_rows:.1%} of rows)")
log(f"Unique reviews after removing copies: {len(copies):,}")
log("How many asins each unique review is copied to:", copies.value_counts().sort_index().to_dict())
rating_conflicts = raw.groupby(dup_key)["overall"].nunique().gt(1).sum()
log(f"Copies with conflicting ratings: {rating_conflicts}  (copies are true duplicates)")

fig, ax = plt.subplots(figsize=(7, 4))
cv = copies.value_counts().sort_index()
ax.bar(cv.index.astype(str), cv.values, color="#d1495b")
for x, v in zip(cv.index.astype(str), cv.values):
    ax.text(x, v, f"{v}", ha="center", va="bottom", fontsize=9)
ax.set(title=f"Each unique review appears under 1-5 product variants\n"
             f"{n_rows:,} rows -> {len(copies):,} unique reviews",
       xlabel="Number of asins the same review is listed under", ylabel="Unique reviews")
save_fig("02_duplicate_copies.png")

# all remaining exploration is done on BOTH views where it matters, but the
# de-duplicated data is the one used for modelling
dedup = raw.drop_duplicates(subset=dup_key).copy()

# --- 1b/1c. Reviews across products and per product ---
per_prod_raw = raw["asin"].value_counts()
per_prod = dedup["asin"].value_counts()
log(f"\nReviews per product (raw):   mean {per_prod_raw.mean():.1f}, median {per_prod_raw.median():.0f}, "
    f"max {per_prod_raw.max()}")
log(f"Reviews per product (dedup): mean {per_prod.mean():.1f}, median {per_prod.median():.0f}, "
    f"max {per_prod.max()}  over {len(per_prod)} products")
top5_share = per_prod.head(5).sum() / per_prod.sum()
log(f"Top 5 products hold {top5_share:.1%} of unique reviews; top 2 hold "
    f"{per_prod.head(2).sum() / per_prod.sum():.1%}")
save_table(per_prod.rename_axis("asin").reset_index(name="unique_reviews").head(20),
           "02_top20_products.csv")

fig, axes = plt.subplots(1, 2, figsize=(12, 4.2))
# 1c: histogram of number of reviews per product (log x because of the long tail)
bins = np.logspace(0, np.log10(per_prod_raw.max() + 1), 20)
axes[0].hist(per_prod_raw, bins=bins, alpha=0.55, label="raw rows", color="#edae49")
axes[0].hist(per_prod, bins=bins, alpha=0.75, label="unique reviews", color="#00798c")
axes[0].set_xscale("log")
axes[0].set(title="Number of reviews per product", xlabel="Reviews per product (log scale)",
            ylabel="Number of products")
axes[0].legend()
# 1b: how reviews are spread across products (cumulative share / Pareto curve)
cum = per_prod.cumsum() / per_prod.sum()
axes[1].plot(np.arange(1, len(cum) + 1), cum.values, marker="o", ms=3, color="#00798c")
axes[1].axhline(0.8, ls="--", c="grey", lw=1)
k80 = int(np.argmax(cum.values >= 0.8)) + 1
axes[1].annotate(f"{k80} of {len(cum)} products\n= 80% of reviews", (k80, 0.8),
                 xytext=(k80 + 8, 0.55), arrowprops=dict(arrowstyle="->", color="black"))
axes[1].set(title="Distribution of reviews across products", xlabel="Products ranked by review count",
            ylabel="Cumulative share of reviews")
save_fig("03_reviews_per_product.png")

# average rating per product vs its popularity
prod_stats = dedup.groupby("asin").agg(n=("overall", "size"), avg=("overall", "mean"))
fig, ax = plt.subplots(figsize=(7, 4))
ax.scatter(prod_stats["n"], prod_stats["avg"], s=25, alpha=0.7, color="#00798c")
ax.set_xscale("log")
ax.set(title="Average rating vs number of reviews per product",
       xlabel="Unique reviews for the product (log)", ylabel="Average rating")
save_fig("04_product_rating_vs_popularity.png")

# --- 1d. Reviews per user ---
per_user_raw = raw["reviewerID"].value_counts()
per_user = dedup["reviewerID"].value_counts()
log(f"\nReviews per user (raw):   min {per_user_raw.min()}, mean {per_user_raw.mean():.2f}, "
    f"max {per_user_raw.max()}  <- the 5-core guarantee")
log(f"Reviews per user (dedup): min {per_user.min()}, mean {per_user.mean():.2f}, max {per_user.max()}; "
    f"{(per_user == 1).mean():.1%} of users wrote only 1 unique review")

fig, axes = plt.subplots(1, 2, figsize=(12, 4))
for ax, s, title, c in [(axes[0], per_user_raw, "raw rows (5-core)", "#edae49"),
                        (axes[1], per_user, "unique reviews", "#00798c")]:
    vc = s.value_counts().sort_index()
    ax.bar(vc.index.astype(str), vc.values, color=c)
    ax.set(title=f"Reviews per user - {title}", xlabel="Reviews written", ylabel="Number of users")
save_fig("05_reviews_per_user.png")

# --- 1e/1f. Review lengths and outliers ---
w = dedup["n_words"]
q1, q3 = w.quantile([0.25, 0.75])
iqr = q3 - q1
upper = q3 + 1.5 * iqr
log(f"\nReview length (words, dedup): mean {w.mean():.1f}, median {w.median():.0f}, "
    f"Q1 {q1:.0f}, Q3 {q3:.0f}, max {w.max()}")
log(f"IQR upper fence = {upper:.0f} words -> {(w > upper).sum()} long outliers "
    f"({(w > upper).mean():.1%}); {(w <= 3).sum()} reviews of 3 words or fewer; "
    f"{(w == 0).sum()} empty")
length_by_rating = dedup.groupby("overall")["n_words"].agg(["count", "mean", "median", "max"]).round(1)
log("Length by rating:\n" + length_by_rating.to_string())
save_table(length_by_rating, "03_length_by_rating.csv", index=True)

fig, axes = plt.subplots(1, 3, figsize=(16, 4.2))
axes[0].hist(w, bins=60, color="#00798c")
axes[0].axvline(w.median(), c="k", ls="--", lw=1, label=f"median {w.median():.0f}")
axes[0].axvline(upper, c="#d1495b", ls="--", lw=1, label=f"IQR fence {upper:.0f}")
axes[0].set(title="Review length distribution", xlabel="Words in review", ylabel="Reviews")
axes[0].legend()
axes[1].boxplot(w, orientation="horizontal", widths=0.6)
axes[1].set(title="Review length box plot (outliers = dots)", xlabel="Words in review", yticks=[])
sns.boxplot(data=dedup, x="overall", y="n_words", ax=axes[2], color="#edae49", showfliers=False)
axes[2].set(title="Review length by rating (outliers hidden)", xlabel="Star rating", ylabel="Words")
save_fig("06_review_length.png")

# summary (headline) length
s_w = dedup["summary_words"]
log(f"Summary length: mean {s_w.mean():.1f} words, median {s_w.median():.0f}, max {s_w.max()}")

# --- Other attributes and relationships ---
# verified purchase vs rating
ver = pd.crosstab(dedup["verified"], dedup["overall"], normalize="index").round(3)
log("\nRating mix by verified status (row %):\n" + ver.to_string())
log(f"Avg rating verified {dedup.loc[dedup.verified, 'overall'].mean():.2f} vs "
    f"not verified {dedup.loc[~dedup.verified, 'overall'].mean():.2f}")

# helpful votes vs rating
votes_by_rating = dedup.groupby("overall")["votes"].agg(
    share_voted=lambda v: (v > 0).mean(), mean_votes="mean").round(3)
log("Helpful votes by rating:\n" + votes_by_rating.to_string())

# images
log(f"Reviews with images: {dedup['has_image'].sum()} "
    f"(avg rating {dedup.loc[dedup.has_image, 'overall'].mean():.2f})")

# style attributes (what variant was bought)
style_keys = Counter(k.strip(":") for s in raw["style"].dropna() for k in s.keys())
log("Style attribute keys (raw rows):", dict(style_keys.most_common()))

fig, axes = plt.subplots(2, 2, figsize=(13, 8.5))
# reviews and average rating over time
yearly = dedup.groupby("year").agg(n=("overall", "size"), avg=("overall", "mean"))
axes[0, 0].bar(yearly.index, yearly["n"], color="#00798c")
axes[0, 0].set(title="Unique reviews per year", xlabel="Year", ylabel="Reviews")
ax2 = axes[0, 0].twinx()
ax2.plot(yearly.index, yearly["avg"], c="#d1495b", marker="o")
ax2.set_ylabel("Average rating", color="#d1495b")
ax2.set_ylim(1, 5.2)
ax2.grid(False)
# verified vs rating
pd.crosstab(dedup["overall"], dedup["verified"]).plot(kind="bar", ax=axes[0, 1],
                                                     color=["#edae49", "#00798c"])
axes[0, 1].set(title="Ratings by verified purchase", xlabel="Star rating", ylabel="Reviews")
axes[0, 1].legend(title="verified")
axes[0, 1].tick_params(axis="x", rotation=0)
# helpful votes vs rating
votes_by_rating["share_voted"].plot(kind="bar", ax=axes[1, 0], color="#edae49")
axes[1, 0].set(title="Share of reviews with >= 1 helpful vote", xlabel="Star rating", ylabel="Share")
axes[1, 0].tick_params(axis="x", rotation=0)
# style attributes
sk = pd.Series(dict(style_keys.most_common(8)))
axes[1, 1].barh(sk.index[::-1], sk.values[::-1], color="#00798c")
axes[1, 1].set(title="Variant attributes in 'style' (raw rows)", xlabel="Rows")
save_fig("07_other_attributes.png")

# --- Summaries: Amazon auto-generates "Five Stars"/"One Star" titles ---
auto_summary = re.compile(r"^(one|two|three|four|five) stars?$", re.I)
dedup["auto_summary"] = dedup["summary"].str.strip().str.match(auto_summary)
log(f"\nAuto-generated summaries ('Five Stars' etc.): {dedup['auto_summary'].sum()} "
    f"({dedup['auto_summary'].mean():.1%}) - these restate the rating, so they leak the label")
log("Most common summaries:", dedup["summary"].value_counts().head(8).to_dict())

# --- Most frequent words per rating class (light cleaning, stop words removed) ---
STOP = set(stopwords.words("english"))
def top_words(texts, n=15):
    """Return the n most common non-stop-word tokens in a list of texts."""
    toks = []
    for t in texts:
        toks += [x for x in re.findall(r"[a-z']+", t.lower()) if x not in STOP and len(x) > 2]
    return Counter(toks).most_common(n)

def to_label(r):
    """Map a star rating to a sentiment label (step 2a)."""
    return "Positive" if r >= 4 else ("Neutral" if r == 3 else "Negative")

dedup["label"] = dedup["overall"].apply(to_label)
fig, axes = plt.subplots(1, 3, figsize=(16, 4.8))
for ax, lab in zip(axes, LABELS):
    tw = top_words(dedup.loc[dedup.label == lab, "reviewText"])
    words, freq = zip(*tw)
    ax.barh(words[::-1], freq[::-1], color=CLASS_COLORS[lab])
    ax.set(title=f"Top words - {lab} reviews", xlabel="Frequency")
save_fig("08_top_words_by_class.png")

fig, axes = plt.subplots(1, 3, figsize=(18, 5))
for ax, lab in zip(axes, LABELS):
    text = " ".join(dedup.loc[dedup.label == lab, "reviewText"])
    wc = WordCloud(width=700, height=450, background_color="white", stopwords=STOP, collocations=False,
                   colormap={"Negative": "Reds", "Neutral": "YlOrBr", "Positive": "GnBu"}[lab],
                   random_state=RANDOM_SEED).generate(text)
    ax.imshow(wc, interpolation="bilinear")
    ax.axis("off")
    ax.set_title(f"{lab} reviews", fontsize=14)
save_fig("09_wordclouds.png")

eda_summary = {
    "raw_rows": int(n_rows), "products": int(n_products), "reviewers": int(n_users),
    "avg_rating": round(float(raw["overall"].mean()), 3),
    "duplicate_copies": int(key_dups), "unique_reviews": int(len(dedup)),
    "median_words": float(w.median()), "iqr_upper_fence_words": float(upper),
    "long_outliers": int((w > upper).sum()), "empty_reviews": int((w == 0).sum()),
    "auto_summaries": int(dedup["auto_summary"].sum()),
}
with open(os.path.join(RES_DIR, "eda_summary.json"), "w") as f:
    json.dump(eda_summary, f, indent=2)


# ============================================================================
# STEP 2 - LABELLING, COLUMN SELECTION, OUTLIERS
# ============================================================================
section("STEP 2 - LABELLING, COLUMN SELECTION, OUTLIERS")

# 2a. label from the rating: 4-5 Positive, 3 Neutral, 1-2 Negative (done above)
label_counts = dedup["label"].value_counts().reindex(LABELS)
log("Label distribution (unique reviews):")
for lab, c in label_counts.items():
    log(f"  {lab:<9}{c:>6}  ({c / len(dedup):.1%})")

fig, ax = plt.subplots(figsize=(6.5, 4))
ax.bar(LABELS, label_counts.values, color=[CLASS_COLORS[l] for l in LABELS])
for i, v in enumerate(label_counts.values):
    ax.text(i, v, f"{v:,} ({v / len(dedup):.1%})", ha="center", va="bottom")
ax.set(title="Sentiment labels derived from rating (unique reviews)", ylabel="Reviews")
ax.set_ylim(0, label_counts.max() * 1.15)
save_fig("10_label_distribution.png")

# 2b. Columns used by the sentiment analyser:
#   reviewText -> main opinion text
#   summary    -> short, opinion-dense headline; added in front of the body,
#                 EXCEPT auto-generated "Five Stars" style titles (they restate the
#                 rating, i.e. label leakage, and carry no lexicon sentiment anyway)
#   overall    -> only used to build the ground-truth label
#   reviewerID, reviewText, unixReviewTime -> only used as the de-duplication key
# Not used: verified, vote, image, style, reviewerName, reviewTime, asin
#   (metadata, no opinion words; a lexicon model cannot use them)
dedup["summary_used"] = np.where(dedup["auto_summary"], "", dedup["summary"].str.strip())

def combine(summary, body):
    """Join summary and body; add a full stop so VADER treats them as separate clauses."""
    summary = summary.strip()
    if summary and summary[-1] not in ".!?":
        summary += "."
    return (summary + " " + body.strip()).strip()

dedup["text"] = [combine(s, b) for s, b in zip(dedup["summary_used"], dedup["reviewText"])]

# 2c. Outliers
#   - rating: all values are in {1..5}, no invalid ratings
#   - empty text after combining: removed (nothing to score)
#   - very short reviews ("Love it!"): KEPT - short but clearly opinionated
#   - very long reviews (> IQR fence): KEPT - VADER/TextBlob average over the text;
#     we check accuracy by length bucket in step 7 to confirm this is safe
#   - non-English / symbol-only text: checked below
log(f"\nInvalid ratings: {(~dedup['overall'].isin([1, 2, 3, 4, 5])).sum()}")
empty = dedup["text"].str.strip().eq("")
log(f"Empty text after combining summary + body: {empty.sum()} -> removed")
non_ascii_share = dedup["text"].apply(lambda t: sum(ord(ch) > 127 for ch in t) / max(len(t), 1))
log(f"Reviews with >20% non-ASCII characters: {(non_ascii_share > 0.2).sum()}")
has_html = dedup["text"].str.contains(r"&[a-z]+;|<[a-z/]+", regex=True)
has_url = dedup["text"].str.contains(r"https?://|www\.", regex=True)
log(f"Reviews with HTML entities/tags: {has_html.sum()}, with URLs: {has_url.sum()}")
log(f"Long-review outliers kept: {(dedup['n_words'] > upper).sum()}")

clean = dedup.loc[~empty].reset_index(drop=True)
log(f"Clean modelling pool: {len(clean):,} reviews")


# ============================================================================
# STEP 3 - LEXICON STUDY (VADER vs TextBlob vs SentiWordNet)
# ============================================================================
section("STEP 3 - LEXICON STUDY")

vader = SentimentIntensityAnalyzer()
lemmatizer = WordNetLemmatizer()


def penn_to_wn(tag):
    """Map a Penn Treebank POS tag to a WordNet POS (None = not a sentiment-bearing word)."""
    if tag.startswith("J"):
        return wn.ADJ
    if tag.startswith("V"):
        return wn.VERB
    if tag.startswith("N"):
        return wn.NOUN
    if tag.startswith("R"):
        return wn.ADV
    return None


def sentiwordnet_score(text):
    """
    Simple SentiWordNet scorer: POS-tag, lemmatise, take the most frequent sense
    of every word and average (pos - neg). SentiWordNet has no notion of
    negation, intensifiers, punctuation or emoji, and needs word-sense
    disambiguation - this is what the probe sentences below expose.
    """
    scores = []
    for word, tag in nltk.pos_tag(nltk.word_tokenize(text)):
        pos = penn_to_wn(tag)
        if pos is None:
            continue
        synsets = list(swn.senti_synsets(lemmatizer.lemmatize(word.lower(), pos), pos))
        if synsets:
            s = synsets[0]
            scores.append(s.pos_score() - s.neg_score())
    return float(np.mean(scores)) if scores else 0.0


# probe sentences built around things that matter in beauty reviews
probes = [
    "This moisturizer is good.",
    "This moisturizer is not good.",
    "This moisturizer is GOOD!!!",
    "This moisturizer is very good.",
    "The scent is nice, but it broke me out terribly.",
    "Love it :)",
    "Meh. It's ok I guess.",
    "Smells amazing, my skin feels so soft.",
    "Waste of money, it burned my scalp.",
    "Not bad at all.",
]
probe_rows = []
for p in probes:
    probe_rows.append({
        "sentence": p,
        "VADER_compound": round(vader.polarity_scores(p)["compound"], 3),
        "TextBlob_polarity": round(TextBlob(p).sentiment.polarity, 3),
        "TextBlob_subjectivity": round(TextBlob(p).sentiment.subjectivity, 3),
        "SentiWordNet": round(sentiwordnet_score(p), 3),
    })
probe_df = pd.DataFrame(probe_rows)
log(probe_df.to_string(index=False))
save_table(probe_df, "04_lexicon_probe_sentences.csv")

lexicon_compare = pd.DataFrame([
    ["Lexicon size", "~7,500 entries incl. slang, emoticons, emoji", "~2,900 adjectives (pattern)", "~117,000 synsets (WordNet)"],
    ["Built for", "Social media / short informal text", "General text", "General text, per word sense"],
    ["Negation", "Yes (3-word window)", "Yes ('not' flips)", "No"],
    ["Intensifiers / CAPS / !!!", "Yes / Yes / Yes", "Yes / No / No", "No / No / No"],
    ["Emoticons, emoji", "Yes", "Some emoticons", "No"],
    ["'but' contrast", "Yes (weights 2nd clause)", "No", "No"],
    ["Needs POS tagging / WSD", "No", "No", "Yes / Yes (hard)"],
    ["Output", "pos/neu/neg + compound [-1,1]", "polarity [-1,1] + subjectivity [0,1]", "pos/neg/obj per synset"],
    ["Chosen", "YES", "YES", "No"],
], columns=["Property", "VADER", "TextBlob", "SentiWordNet"])
save_table(lexicon_compare, "05_lexicon_comparison.csv")
log("\nChosen lexicons: VADER and TextBlob. SentiWordNet scores 'not good' the same as 'good', "
    "ignores emphasis and emoji, and depends on error-prone sense selection.")


# ============================================================================
# STEP 4 - PRE-PROCESSING PER MODEL (with an ablation to justify each choice)
# ============================================================================
section("STEP 4 - MODEL-SPECIFIC PRE-PROCESSING")

URL_RE = re.compile(r"https?://\S+|www\.\S+")
TAG_RE = re.compile(r"<[^>]+>")
REPEAT_RE = re.compile(r"(.)\1{3,}")   # 4+ repeats of a char -> 2 (soooo -> soo)
CONTRACTIONS = {"n't": " not", "'re": " are", "'ve": " have", "'ll": " will", "'m": " am", "'d": " would"}
NEGATIONS = {"no", "not", "nor", "never", "none", "nothing", "nobody", "neither", "nowhere", "without",
             "don't", "doesn't", "didn't", "isn't", "wasn't", "aren't", "weren't", "won't", "wouldn't",
             "couldn't", "shouldn't", "can't", "cannot", "but", "very", "too", "so", "really"}


def base_clean(text):
    """Shared, safe cleaning: unescape HTML, drop tags and URLs, squeeze whitespace."""
    text = html.unescape(text)
    text = TAG_RE.sub(" ", text)
    text = URL_RE.sub(" ", text)
    return re.sub(r"\s+", " ", text).strip()


def preprocess_vader(text):
    """
    VADER: keep case (CAPS = emphasis), punctuation (!!! = emphasis), emoticons,
    stop words ('not', 'but', 'very' are rules in VADER). Only remove noise and
    shorten stretched words so 'sooooo' matches the lexicon's 'so'.
    """
    text = base_clean(text)
    return REPEAT_RE.sub(r"\1\1", text)


def preprocess_lower_nopunct(text):
    """Ablation only: lower-case, expand contractions, strip punctuation, keep every word."""
    text = base_clean(text).lower()
    for c, full in CONTRACTIONS.items():
        text = text.replace(c, full)
    return text.translate(str.maketrans("", "", string.punctuation))


def preprocess_textblob(text):
    """
    TextBlob: its lexicon is lower-case and has no CAPS rules, so we lower-case.
    Contractions are expanded ("isn't" -> "is not") so the negation is a separate
    word. TextBlob only flips polarity when the negation sits right before the
    sentiment word, so removing stop words (but KEEPING negations and
    intensifiers like 'not', 'very', 'too') pulls "not ... good" together.
    Lemmatising maps 'smells'/'scents' onto lexicon entries. Punctuation is
    removed. The dev-set ablation below confirms this beats light cleaning.
    """
    toks = [lemmatizer.lemmatize(t) for t in preprocess_lower_nopunct(text).split()
            if t not in STOP or t in NEGATIONS]
    return REPEAT_RE.sub(r"\1\1", " ".join(toks))


def preprocess_aggressive(text):
    """
    Ablation only: 'classic' pipeline - lower-case, strip punctuation, remove ALL
    stop words (including 'not'), lemmatise.
    """
    toks = [lemmatizer.lemmatize(t) for t in preprocess_lower_nopunct(text).split() if t not in STOP]
    return " ".join(toks)


def vader_score(text):
    return vader.polarity_scores(text)["compound"]


def textblob_score(text):
    return TextBlob(text).sentiment.polarity


def to_sentiment(score, pos_thr, neg_thr):
    """Convert a continuous lexicon score to Positive / Neutral / Negative."""
    if score >= pos_thr:
        return "Positive"
    if score <= neg_thr:
        return "Negative"
    return "Neutral"


# Default thresholds from the packages' documentation:
#   VADER compound >= 0.05 positive, <= -0.05 negative (Hutto & Gilbert)
#   TextBlob polarity > 0 positive, < 0 negative, == 0 neutral
DEFAULT_THR = {"VADER": (0.05, -0.05), "TextBlob": (1e-9, -1e-9)}

example = clean.loc[clean["n_words"].between(15, 30)].iloc[3, :]["text"]
log("Example review:\n  original   :", example)
log("  VADER      :", preprocess_vader(example))
log("  TextBlob   :", preprocess_textblob(example))
log("  aggressive :", preprocess_aggressive(example))


# ============================================================================
# STEP 5 - RANDOM SAMPLE OF 1000 REVIEWS
# ============================================================================
section("STEP 5 - RANDOM SAMPLE OF 1000 REVIEWS")

sample = clean.sample(n=SAMPLE_SIZE, random_state=RANDOM_SEED).reset_index(drop=True)
# the reviews NOT sampled form a small development set: we use it (and only it)
# to choose pre-processing and to tune thresholds, so the 1000-review test sample
# never influences any modelling decision
dev = clean.drop(clean.sample(n=SAMPLE_SIZE, random_state=RANDOM_SEED).index).reset_index(drop=True)
log(f"Test sample: {len(sample)} reviews   Development set (rest): {len(dev)} reviews")
log("Sample labels:", sample["label"].value_counts().reindex(LABELS).to_dict())
log("Dev labels:   ", dev["label"].value_counts().reindex(LABELS).to_dict())
save_table(sample[["reviewerID", "asin", "overall", "label", "summary", "reviewText", "text"]],
           "06_random_sample_1000.csv")


# --- Pre-processing ablation on the development set (justifies step 4) ---
def macro_f1(y_true, y_pred):
    return precision_recall_fscore_support(y_true, y_pred, labels=LABELS, average="macro",
                                           zero_division=0)[2]

pipelines = {
    "P0 raw text": lambda t: t,
    "P1 light clean (keep case, punct, stop words)": preprocess_vader,
    "P2 lower + contractions + no punct": preprocess_lower_nopunct,
    "P3 P2 + drop stop words except negations + lemma": preprocess_textblob,
    "P4 P2 + drop ALL stop words + lemma": preprocess_aggressive,
}
CHOSEN = {"VADER": "P1 light clean (keep case, punct, stop words)",
          "TextBlob": "P3 P2 + drop stop words except negations + lemma"}
ablation = []
for model, scorer in [("VADER", vader_score), ("TextBlob", textblob_score)]:
    for name, fn in pipelines.items():
        scores = dev["text"].apply(lambda t: scorer(fn(t)))
        pred = scores.apply(lambda s: to_sentiment(s, *DEFAULT_THR[model]))
        ablation.append({"model": model, "pre-processing": name,
                         "accuracy": accuracy_score(dev["label"], pred),
                         "macro_F1": macro_f1(dev["label"], pred),
                         "chosen": "yes" if CHOSEN[model] == name else ""})
ablation_df = pd.DataFrame(ablation).round(3)
log("\nPre-processing ablation (development set, default thresholds):")
log(ablation_df.to_string(index=False))
save_table(ablation_df, "07_preprocessing_ablation_dev.csv")

fig, ax = plt.subplots(figsize=(10, 4.2))
piv = ablation_df.pivot(index="pre-processing", columns="model", values="macro_F1").reindex(list(pipelines))
piv.plot(kind="barh", ax=ax, color=["#edae49", "#00798c"])
ax.invert_yaxis()
ax.set(title="Pre-processing ablation (development set): macro-F1", xlabel="Macro-F1", ylabel="")
save_fig("11_preprocessing_ablation.png")


# --- Threshold tuning on the development set ---
# Neutral (3-star) reviews are usually mildly positive in wording, so the default
# near-zero neutral band is too narrow. We grid-search the thresholds on the dev set.
def tune_thresholds(scores, y, pos_grid, neg_grid):
    best = (-1, None)
    for p in pos_grid:
        for n in neg_grid:
            if n >= p:
                continue
            f = macro_f1(y, [to_sentiment(s, p, n) for s in scores])
            if f > best[0]:
                best = (f, (round(float(p), 3), round(float(n), 3)))
    return best

dev["vader_score"] = dev["text"].apply(lambda t: vader_score(preprocess_vader(t)))
dev["textblob_score"] = dev["text"].apply(lambda t: textblob_score(preprocess_textblob(t)))
f_v, TUNED_V = tune_thresholds(dev["vader_score"], dev["label"],
                               np.arange(0.05, 0.96, 0.05), np.arange(-0.6, 0.51, 0.05))
f_t, TUNED_T = tune_thresholds(dev["textblob_score"], dev["label"],
                               np.arange(0.0, 0.61, 0.025), np.arange(-0.3, 0.31, 0.025))
TUNED_THR = {"VADER": TUNED_V, "TextBlob": TUNED_T}
log(f"\nTuned thresholds on dev: VADER pos>={TUNED_V[0]}, neg<={TUNED_V[1]} (dev macro-F1 {f_v:.3f}); "
    f"TextBlob pos>={TUNED_T[0]}, neg<={TUNED_T[1]} (dev macro-F1 {f_t:.3f})")


# ============================================================================
# STEP 6 - LEXICON MODELS ON THE 1000-REVIEW SAMPLE
# ============================================================================
section("STEP 6 - MODELLING (VADER and TextBlob)")

sample["vader_input"] = sample["text"].apply(preprocess_vader)
sample["textblob_input"] = sample["text"].apply(preprocess_textblob)
sample["vader_score"] = sample["vader_input"].apply(vader_score)
sample["textblob_score"] = sample["textblob_input"].apply(textblob_score)
sample["textblob_subjectivity"] = sample["textblob_input"].apply(lambda t: TextBlob(t).sentiment.subjectivity)

for model, col in [("VADER", "vader"), ("TextBlob", "textblob")]:
    sample[f"{col}_pred"] = sample[f"{col}_score"].apply(lambda s: to_sentiment(s, *DEFAULT_THR[model]))
    sample[f"{col}_pred_tuned"] = sample[f"{col}_score"].apply(lambda s: to_sentiment(s, *TUNED_THR[model]))

log(sample[["label", "vader_score", "vader_pred", "textblob_score", "textblob_pred"]].head(10).to_string())
save_table(sample[["overall", "label", "text", "vader_score", "vader_pred", "vader_pred_tuned",
                   "textblob_score", "textblob_subjectivity", "textblob_pred", "textblob_pred_tuned"]],
           "08_sample_predictions.csv")

# score distributions per true label
fig, axes = plt.subplots(1, 2, figsize=(13, 4.3))
for ax, col, title, thr in [(axes[0], "vader_score", "VADER compound", DEFAULT_THR["VADER"]),
                            (axes[1], "textblob_score", "TextBlob polarity", DEFAULT_THR["TextBlob"])]:
    sns.violinplot(data=sample, x="label", y=col, order=LABELS, ax=ax, cut=0,
                   hue="label", palette=CLASS_COLORS, legend=False)
    ax.axhline(thr[0], c="k", ls="--", lw=0.8)
    ax.axhline(thr[1], c="k", ls="--", lw=0.8)
    ax.set(title=f"{title} by true label (dashed = default thresholds)", xlabel="True label", ylabel=title)
save_fig("12_score_distribution_by_label.png")


# ============================================================================
# STEP 7 - VALIDATION AND COMPARISON
# ============================================================================
section("STEP 7 - VALIDATION AND COMPARISON")

def metrics_row(name, y, p):
    """One row of the comparison table."""
    acc = accuracy_score(y, p)
    mp, mr, mf, _ = precision_recall_fscore_support(y, p, labels=LABELS, average="macro", zero_division=0)
    wp, wr, wf, _ = precision_recall_fscore_support(y, p, labels=LABELS, average="weighted", zero_division=0)
    per = precision_recall_fscore_support(y, p, labels=LABELS, zero_division=0)
    row = {"model": name, "accuracy": acc, "macro_precision": mp, "macro_recall": mr,
           "macro_F1": mf, "weighted_F1": wf}
    for i, lab in enumerate(LABELS):
        row[f"F1_{lab}"] = per[2][i]
    return row

y = sample["label"]
rows = [
    metrics_row("Majority baseline (always Positive)", y, ["Positive"] * len(y)),
    metrics_row("VADER (default thresholds)", y, sample["vader_pred"]),
    metrics_row("TextBlob (default thresholds)", y, sample["textblob_pred"]),
    metrics_row(f"VADER (tuned {TUNED_V})", y, sample["vader_pred_tuned"]),
    metrics_row(f"TextBlob (tuned {TUNED_T})", y, sample["textblob_pred_tuned"]),
]
comparison = pd.DataFrame(rows).round(3)
log("Comparison table (1000-review test sample):")
log(comparison.to_string(index=False))
save_table(comparison, "09_comparison_table.csv")

for model, col in [("VADER", "vader_pred"), ("TextBlob", "textblob_pred")]:
    log(f"\nClassification report - {model} (default thresholds):")
    log(classification_report(y, sample[col], labels=LABELS, digits=3, zero_division=0))

# confusion matrices: default and tuned, counts and row-normalised recall
fig, axes = plt.subplots(2, 2, figsize=(11, 9.5))
for r, (suffix, tag) in enumerate([("", "default thresholds"), ("_tuned", "tuned thresholds")]):
    for c, (model, col) in enumerate([("VADER", "vader_pred"), ("TextBlob", "textblob_pred")]):
        cm = confusion_matrix(y, sample[col + suffix], labels=LABELS)
        ax = axes[r, c]
        annot = np.array([[f"{v}\n({v / max(cm[i].sum(), 1):.0%})" for v in row] for i, row in enumerate(cm)])
        sns.heatmap(cm / cm.sum(axis=1, keepdims=True), annot=annot, fmt="", cmap="Blues", cbar=False,
                    xticklabels=LABELS, yticklabels=LABELS, ax=ax, vmin=0, vmax=1)
        acc = accuracy_score(y, sample[col + suffix])
        ax.set(title=f"{model} - {tag}\naccuracy {acc:.1%}, macro-F1 {macro_f1(y, sample[col + suffix]):.3f}",
               xlabel="Predicted", ylabel="True")
plt.tight_layout()
save_fig("13_confusion_matrices.png")

# side-by-side bar chart of the main metrics
fig, ax = plt.subplots(figsize=(11, 4.5))
plot_df = comparison.set_index("model")[["accuracy", "macro_precision", "macro_recall", "macro_F1",
                                         "F1_Negative", "F1_Neutral", "F1_Positive"]].T
plot_df.plot(kind="bar", ax=ax, width=0.82,
             color=["#bbbbbb", "#edae49", "#00798c", "#f4d58d", "#7fc8d0"])
ax.set(title="Model comparison on the 1000-review sample", ylabel="Score", ylim=(0, 1.05))
ax.tick_params(axis="x", rotation=0)
ax.legend(fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3)
save_fig("14_model_comparison.png")

# agreement between the two models
agree = (sample["vader_pred"] == sample["textblob_pred"]).mean()
kappa = cohen_kappa_score(sample["vader_pred"], sample["textblob_pred"])
both_right = ((sample["vader_pred"] == y) & (sample["textblob_pred"] == y)).mean()
both_wrong = ((sample["vader_pred"] != y) & (sample["textblob_pred"] != y)).mean()
corr = sample[["vader_score", "textblob_score"]].corr().iloc[0, 1]
log(f"\nVADER vs TextBlob: agree on {agree:.1%} of reviews, Cohen's kappa {kappa:.3f}, "
    f"score correlation r = {corr:.3f}")
log(f"Both correct {both_right:.1%}, both wrong {both_wrong:.1%}")

fig, ax = plt.subplots(figsize=(6.5, 5.5))
for lab in LABELS:
    m = sample["label"] == lab
    ax.scatter(sample.loc[m, "vader_score"], sample.loc[m, "textblob_score"], s=14, alpha=0.6,
               label=lab, color=CLASS_COLORS[lab])
ax.set(title=f"VADER vs TextBlob scores (r = {corr:.2f})", xlabel="VADER compound",
       ylabel="TextBlob polarity")
ax.legend(title="True label")
save_fig("15_vader_vs_textblob.png")

# accuracy by review length (are long-review outliers a problem?)
sample["length_bucket"] = pd.cut(sample["n_words"], [-1, 10, 30, 70, 150, 10_000],
                                 labels=["1-10", "11-30", "31-70", "71-150", "150+"])
by_len = sample.groupby("length_bucket", observed=True).apply(
    lambda g: pd.Series({"reviews": len(g),
                         "VADER_acc": accuracy_score(g["label"], g["vader_pred"]),
                         "TextBlob_acc": accuracy_score(g["label"], g["textblob_pred"])}),
    include_groups=False).round(3)
log("\nAccuracy by review length (words):\n" + by_len.to_string())
save_table(by_len, "10_accuracy_by_length.csv", index=True)

fig, ax = plt.subplots(figsize=(8, 4))
by_len[["VADER_acc", "TextBlob_acc"]].plot(kind="bar", ax=ax, color=["#edae49", "#00798c"])
for i, n in enumerate(by_len["reviews"]):
    ax.text(i, 0.02, f"n={int(n)}", ha="center", fontsize=8, color="white")
ax.set(title="Accuracy by review length", xlabel="Words in review", ylabel="Accuracy", ylim=(0, 1))
ax.tick_params(axis="x", rotation=0)
save_fig("16_accuracy_by_length.png")

# robustness: same models on the whole clean pool (all unique reviews)
clean["vader_pred"] = clean["text"].apply(lambda t: to_sentiment(vader_score(preprocess_vader(t)),
                                                                 *DEFAULT_THR["VADER"]))
clean["textblob_pred"] = clean["text"].apply(lambda t: to_sentiment(textblob_score(preprocess_textblob(t)),
                                                                    *DEFAULT_THR["TextBlob"]))
full = pd.DataFrame([metrics_row("VADER (all unique reviews)", clean["label"], clean["vader_pred"]),
                     metrics_row("TextBlob (all unique reviews)", clean["label"], clean["textblob_pred"])]).round(3)
log(f"\nRobustness check on all {len(clean):,} unique reviews:")
log(full.to_string(index=False))
save_table(full, "11_full_pool_metrics.csv")

# error analysis: the most confident mistakes of each model
errors = []
for model, col in [("VADER", "vader"), ("TextBlob", "textblob")]:
    wrong = sample[sample[f"{col}_pred"] != sample["label"]].copy()
    wrong["abs_score"] = wrong[f"{col}_score"].abs()
    for _, r in wrong.sort_values("abs_score", ascending=False).head(6).iterrows():
        errors.append({"model": model, "true": r["label"], "pred": r[f"{col}_pred"],
                       "score": round(r[f"{col}_score"], 3), "text": r["text"][:220]})
errors_df = pd.DataFrame(errors)
log("\nMost confident errors:")
for _, e in errors_df.iterrows():
    log(f"  [{e['model']}] true={e['true']:<8} pred={e['pred']:<8} score={e['score']:+.3f} | {e['text'][:120]}")
save_table(errors_df, "12_error_examples.csv")

with open(os.path.join(RES_DIR, "run_log.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(_LOG))
log("\nDone. Figures in ./figures, tables in ./results, full log in results/run_log.txt")
