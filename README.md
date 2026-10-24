<div align="center">

# 💬 WhatsApp Chat Analyzer
### *Turning Conversations into Insights with Microsoft Excel*


> *"What if your WhatsApp chats could tell you who matters most?"*

**95 Contacts | 110 Days | 13,830 Messages | 2 Excel Files | 11 Sheets | 100% Excel**

---

</div>

## 🧠 What Is This?

This project takes **raw WhatsApp chat export data** from 95 contacts over a period of **January 19 – May 8, 2026** and transforms it into a fully interactive Excel-based analytics system.

No Python. No SQL. No Power BI.
Just **pure Excel magic** — from messy raw files to clean dashboards.

---

## 📊 Dataset at a Glance

| 📌 Metric | 📈 Value |
|-----------|---------|
| 👥 Total Contacts | 95 |
| 👨 Male Contacts | 85 (89%) |
| 👩 Female Contacts | 10 (11%) |
| 📅 Date Range | Jan 19, 2026 — May 8, 2026 |
| 🗓️ Total Days Tracked | 110 days |
| 💬 Total Messages | 13,830 |
| 🏆 Most Active Contact | 12M (2,420 messages!) |
| 🔥 Very High Activity | 6 contacts |
| 📊 High Activity | 4 contacts |
| 🟡 Medium Activity | 15 contacts |
| 🔵 Low Activity | 70 contacts |

---

## 📁 Project Structure

```
📦 WhatsApp Chat Analysis
│
├── 📂 excel_project/                  ← 95 individual contact files
│   ├── 1M.xlsx                        ← Contact 1, Male
│   ├── 2M.xlsx                        ← Contact 2, Male
│   ├── 3F.xlsx                        ← Contact 3, Female
│   └── ... (95 total files)
│
├── 📊 Contact_Summary_Table.xlsx      ← Main summary & analysis
│   ├── 📋 Summary                     ← 95 rows × 29 columns
│   ├── 🔬 Analysis                    ← Text, Logical & Aggregation functions
│   ├── 🖥️ Dashboard                   ← Interactive XLOOKUP lookup
│   ├── 🏆 Rank_Analysis               ← Contacts ranked by messages
│   └── 📊 Pivot_Analysis              ← Pivot Tables + Slicers + Charts
│
└── 📊 Master_All_Messages.xlsx        ← Daily combined data
    ├── 📅 All_Messages                ← 10,450+ daily rows
    ├── 📆 Date_Analysis               ← Month, Weekday, DATEDIF
    ├── 📉 Gap_Analysis                ← Days between messages
    ├── 📅 Weekday_Analysis            ← Mon–Sun pattern
    ├── 🗓️ Month_Analysis              ← Jan–May breakdown
    └── 🗂️ Contact_ID                  ← Dropdown source list
```

---

## 🏷️ File Naming Convention

Every individual file follows this pattern:

```
[Contact Number][Gender].xlsx

Examples:
  12M.xlsx  →  Contact 12, Male
  3F.xlsx   →  Contact 3, Female
  41F.xlsx  →  Contact 41, Female
```

---

## 🛠️ Excel Features & Where They're Used

### ⚡ Power Query
| Feature | Where Used |
|---------|-----------|
| Combine Files from Folder | Merging all 95 files |
| Custom Column | Extracting Contact_ID & Gender from filename |
| Pivot Column | Converting tall → wide format |
| Filter Rows | Removing null/header rows |
| Data Type Changes | Text, Number, Date conversions |
| Close & Load | Loading to Excel Table |

---

### 🔍 Lookup Functions
| Function | Purpose |
|----------|---------|
| `XLOOKUP` | Dashboard — type Contact ID → get all stats |
| `INDEX MATCH` | Rank-based contact lookup |
| `LARGE + INDEX` | Top contacts by total messages |

---

### ➕ Aggregation Functions
| Function | Example Use |
|----------|------------|
| `SUMIFS` | Total messages by gender |
| `COUNTIFS` | Active days per weekday/month |
| `AVERAGEIFS` | Average messages — Male vs Female |
| `MAXIFS` | Most active Male / Female contact |
| `MINIFS` | Least active contact |
| `SUMPRODUCT` | Multi-condition month/weekday analysis |

---

### 🧠 Logical Functions
| Function | Use Case |
|----------|---------|
| `IFS` | Activity Level: Very High / High / Medium / Low |
| `AND` | Communication Type tagging |
| `IF` | Is_Active flag (message day vs silent day) |
| `IFERROR` | Clean error handling in Gap Analysis |

---

### 📝 Text Functions
| Function | Use Case |
|----------|---------|
| `LEFT` | Extract contact number from ID |
| `RIGHT` | Extract gender letter (M/F) |
| `LEN` | Length of Contact_ID |
| `TEXTJOIN` | Build label: "Contact 10 (Male)" |
| `SUBSTITUTE` | M → Male, F → Female |
| `TRIM` | Clean extra whitespace |

---

### 📅 Date Functions
| Function | Use Case |
|----------|---------|
| `MONTH` | Monthly trend analysis |
| `YEAR` | Year extraction |
| `WEEKDAY` | Day number (1=Mon, 7=Sun) |
| `TEXT` | Format date as "Monday", "January" etc. |
| `EOMONTH` | Month-end reference dates |
| `DATEDIF` | Days since first message |

---

### 📊 Pivot Tables & Charts
| Feature | Details |
|---------|---------|
| Pivot Table 1 | Gender × Total/My/Your Messages |
| Pivot Table 2 | Activity Level × Count & Sum |
| Slicers | Gender filter — connected to both Pivots |
| Pie Chart | Male vs Female message distribution |
| Bar Chart | Top 10 most active contacts |
| Column Chart | Contacts by Activity Level |

---

### 🎨 Conditional Formatting
| Rule | Column | Type |
|------|--------|------|
| Top 10% | Total Messages | Green Fill |
| Data Bars | Total Messages | Blue Bars |
| Zero highlight | Call Sent | Red Fill |
| Color Scale | Days with Messages | Green→Red |
| Color Scale | Your Reactions | Green→Red |
| Icon Sets | Call Received (Answered) | 3 Arrows |
| Highlight Rules | Activity_Level | Color Coded |
| Data Bars | Voice Note Sent | Orange |

---

### ✅ Data Validation
| Location | Type | Source |
|----------|------|--------|
| Dashboard C2 | Dropdown | Contact_ID list |
| Summary Gender col | Restriction | M, F only |
| Weekday_Analysis G1 | Dropdown | Contact_List sheet |
| Month_Analysis B1 | Dropdown | Contact_List sheet |

---

### ⚡ Dynamic Arrays
| Function | Use Case |
|----------|---------|
| `FILTER` | Extract active dates per contact |
| `SORT + CHOOSE` | Rank analysis sorted by messages |
| `UNIQUE` | Unique Contact ID list |
| `SEQUENCE` | Auto rank numbering 1–95 |

---

## 📋 Sheets — Detailed Breakdown

### 📊 Contact_Summary_Table.xlsx

#### Summary Sheet
- **95 rows** — one per contact
- **29 columns** — Contact_ID, Gender, all metrics
- Activity_Level column: Very High / High / Medium / Low
- Conditional Formatting applied on multiple columns

#### Analysis Sheet
Demonstrates all text & logical functions:
- `Contact_Number` — LEFT() extracted
- `Gender_Code` — RIGHT() extracted  
- `Gender_Full` — SUBSTITUTE() (M→Male, F→Female)
- `Contact_Label` — TEXTJOIN() ("Contact 10 (Male)")
- `Activity_Level` — IFS() based on Total Messages
- `Communication_Type` — AND/IFS (Multi-Channel / Caller / Voice Note User / Text Only)
- Gender Analysis section — SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS, MINIFS

#### Dashboard Sheet
- Type any Contact ID → instantly see all 29 metrics
- Dropdown list for easy contact selection
- XLOOKUP powers every metric cell
- RANK formula shows contact's position among all 95

#### Rank_Analysis Sheet
- SORT + CHOOSE dynamic array
- Auto-sorted by Total Messages (descending)
- Rank 1 = most active contact (12M with 2,420 messages!)

#### Pivot_Analysis Sheet
- 2 interactive Pivot Tables
- Gender Slicer connected to both
- 3 charts: Pie, Bar, Column

---

### 📊 Master_All_Messages.xlsx

#### All_Messages Sheet
- **10,450+ rows** of daily data
- All 95 contacts × 110 days
- 32 columns including Is_Active flag
- Null rows preserved for date continuity

#### Date_Analysis Sheet
- Month number, Year, Weekday number
- TEXT() for day/month names
- EOMONTH for month-end dates
- DATEDIF for days since first message

#### Weekday_Analysis Sheet
Overall pattern (all contacts combined):
| Day | Total Messages |
|-----|---------------|
| Wed | 2,183 ← Most Active |
| Tue | 2,129 |
| Mon | 2,049 |
| Thu | 2,077 |
| Fri | 1,999 |
| Sat | 1,827 |
| Sun | 1,566 ← Least Active |

Per-contact section with dropdown — shows individual weekday pattern with I Started / You Started / I Ended / You Ended breakdown.

#### Month_Analysis Sheet
Per-contact monthly breakdown — Jan to May with Started By / Ended By analysis.

#### Gap_Analysis Sheet
- Per-contact message gap tracker
- Shows exact days between consecutive conversations
- FILTER + INDEX MATCH logic
- Identifies longest silence periods

---

## 💡 Key Insights This Project Reveals

- 🏆 **Who are your top contacts?** — Rank Analysis
- 👥 **Male vs Female messaging behavior** — Pivot + Charts
- 📅 **Which day you chat most?** — Weekday Analysis
- 🗓️ **Which month was most active?** — Month Analysis
- ⏱️ **How long do you go without texting someone?** — Gap Analysis
- 💬 **Who starts conversations more — you or them?** — Started By column
- 📞 **Who calls vs texts?** — Communication Type
- 🔥 **Your most loyal chatters** — Activity Level

---

## 🚀 How to Use

1. **Clone/Download** this repository
2. Open `Contact_Summary_Table.xlsx`
3. Go to **Dashboard** sheet → type any Contact ID (e.g. `12M`)
4. All stats appear instantly! ✨
5. Open `Master_All_Messages.xlsx`
6. Go to **Weekday_Analysis** or **Month_Analysis**
7. Select Contact ID from dropdown → see their pattern

> ⚠️ If prompted, click **"Enable Content"** to allow data connections

---

## 👨‍💻 Author

<div align="center">

**Muhammad Uzair Hussain**

BSCS Student — FAST NUCES CFD Campus

---

*Built with 💚 in Microsoft Excel | No code, just formulas*

</div>
