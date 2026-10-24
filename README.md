# 📊 Social Media Performance Report Analysis

**Digital Marketing Track — Task: Analyze a Social Media Performance Report**
Author: **Sadhanala Nischay Prasad** · Tools: **Excel / Google Sheets** (+ Python for charts)

---

## 🎯 Objective
Learn basic social media analytics by analysing a sample dataset of 40 posts and identifying the best and worst-performing posts, then turning the numbers into insights and recommendations.

## 📁 Repository Structure
```
social-media-performance-analysis/
├── README.md
├── data/
│   └── social_media_data.csv                 # Raw sample dataset (40 posts)
├── Social_Media_Performance_Analysis.xlsx    # Main analysis workbook (formulas + charts)
├── Social_Media_Performance_Report.pdf       # Final task report
├── charts/                                   # Chart images used in the report
│   ├── er_by_post_type.png
│   ├── er_by_platform.png
│   ├── er_by_time_slot.png
│   ├── er_by_topic.png
│   └── reach_vs_ctr.png
└── scripts/                                  # Optional: Python used to build data/charts
    ├── gen_data.py
    ├── analyze.py
    ├── build_xlsx.py
    └── report.py
```

## 🗂️ Dataset
A **synthetic sample dataset** of 40 posts (Jul–Sep 2026) across **Instagram, Facebook, LinkedIn and Twitter/X**.

| Column | Description |
|---|---|
| Post_ID, Date, Day | Post identifier and publish date |
| Platform | Instagram / Facebook / LinkedIn / Twitter/X |
| Post_Type | Reel/Video, Carousel, Image, Text |
| Topic | Product Launch, Tips & Tutorial, Meme/Trend, etc. |
| Time_Slot | Morning, Afternoon, Evening, Night |
| Reach | Unique people who saw the post |
| Impressions | Total number of times the post was shown |
| Likes, Comments, Shares, Saves | Engagement actions |
| Clicks | Link / profile clicks |

## 📐 Metrics & Formulas
| Metric | Formula | Excel example |
|---|---|---|
| Total Engagement | Likes + Comments + Shares + Saves | `=SUM(J2:M2)` |
| Engagement Rate | Total Engagement ÷ Reach | `=IFERROR(O2/H2,0)` |
| CTR | Clicks ÷ Impressions | `=IFERROR(N2/I2,0)` |
| Rank (by ER) | Rank with tie-break | `=RANK(P2,$P$2:$P$41,0)+COUNTIF($P$2:P2,P2)-1` |
| Group averages | Avg per category | `=AVERAGEIF('Raw Data'!$E$2:$E$41,A17,'Raw Data'!$P$2:$P$41)` |
| Top/Bottom lookup | Pull post by rank | `=INDEX(range,MATCH(rank,'Raw Data'!$R$2:$R$41,0))` |

> Posts are ranked by **engagement rate**, not raw likes, because high-reach posts naturally get more likes. The rate shows how well the content actually performed.

## 📘 Workbook Sheets
1. **Raw Data** – dataset + calculated columns, colour scale on engagement rate, filters.
2. **Summary** – overall KPIs and breakdowns by post type, platform, time slot and topic, with charts.
3. **Top & Bottom Posts** – top 5 and bottom 5 posts, pulled automatically with `INDEX/MATCH`.
4. **Insights & Recs** – key findings and the three recommendations.

Everything is formula-driven — change any number in *Raw Data* and every table, ranking and chart updates. Opens in both Excel and Google Sheets.

## 📈 Overall Performance
| Posts | Total Reach | Total Impressions | Total Engagement | Total Clicks | Avg Engagement Rate | Avg CTR |
|---|---|---|---|---|---|---|
| 40 | 263,696 | 413,428 | 17,994 | 7,918 | 5.93% | 1.97% |

## 🏆 Top-Performing Posts
| Post | Platform | Type | Topic | Time | Eng. Rate |
|---|---|---|---|---|---|
| P004 | Instagram | Carousel | Meme/Trend | Evening | 13.62% |
| P030 | Twitter/X | Reel/Video | Meme/Trend | Evening | 12.80% |
| P039 | Facebook | Reel/Video | Behind the Scenes | Evening | 10.67% |
| P002 | LinkedIn | Reel/Video | Meme/Trend | Morning | 9.92% |
| P005 | Twitter/X | Reel/Video | Festival Greeting | Evening | 9.82% |

## 📉 Low-Performing Posts
| Post | Platform | Type | Topic | Time | Eng. Rate |
|---|---|---|---|---|---|
| P040 | LinkedIn | Text | Offer/Discount | Morning | 1.47% |
| P008 | Instagram | Text | Offer/Discount | Night | 1.80% |
| P031 | Facebook | Text | Behind the Scenes | Afternoon | 1.82% |
| P013 | Twitter/X | Text | Tips & Tutorial | Morning | 2.00% |
| P026 | LinkedIn | Text | Product Launch | Morning | 2.14% |

## 💡 Engagement Insights
- **Format matters most:** Reels/Videos average **9.0%** engagement vs **2.2%** for Text — about 4x.
- **Carousels get saved:** ~85 saves per post, the highest of any format.
- **Evening wins:** 6–9 PM posts average **8.5%** engagement vs **4.4%** in the afternoon.
- **Reach vs clicks:** Instagram has the biggest reach (~10.9K/post) but one of the lowest CTRs; LinkedIn has the smallest reach but the best CTR (**2.8%**).
- **Trends travel:** Meme/Trend posts average **10.9%**; Industry News is among the weakest topics.
- **Pattern:** all top 5 posts are Reels/Carousels, all bottom 5 are plain Text posts.

![Engagement rate by post type](charts/er_by_post_type.png)
![Engagement rate by time slot](charts/er_by_time_slot.png)
![Reach vs CTR by platform](charts/reach_vs_ctr.png)

## ✅ Three Recommendations
1. **Shift the content mix to Reels and Carousels** (60–70% of posts) and convert text posts into short videos or carousels.
2. **Post key content between 6 and 9 PM** and avoid the 1–3 PM slot for launches and offers.
3. **Match each platform to a goal:** Instagram for reach and awareness; LinkedIn and Facebook for traffic, with a clear CTA and link in every post.

## ⚠️ Limitations
The data is synthetic and small (40 posts); some groups contain only 2–5 posts. The patterns are for learning — a real account should test these recommendations over 4–8 weeks.

## ▶️ How to Use
- **Excel:** open `Social_Media_Performance_Analysis.xlsx`.
- **Google Sheets:** File → Import → Upload the `.xlsx` (or the CSV).
- **Rebuild with Python (optional):** `cd scripts && python gen_data.py && python build_xlsx.py && python report.py` (needs `pandas`, `openpyxl`, `matplotlib`, `reportlab`).

## 🧠 What I Learned
- The difference between **reach** and **impressions**, and why engagement *rate* beats raw likes.
- Using `SUM`, `AVERAGEIF`, `COUNTIF`, `RANK`, `INDEX/MATCH` and conditional formatting for analysis.
- Turning numbers into clear insights and action points.
