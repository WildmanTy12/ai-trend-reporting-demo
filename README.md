# AI Trend Reporting Demo

📊 **AI-Powered Escalation Trend Reporting using Google Apps Script + Google Sheets**

This project demonstrates how automation and AI can transform manual operational reporting into a reliable, scalable insights pipeline — all within Google Workspace tools.

---

## ⚙️ Production-Proven Design
This is a **sanitized, portfolio-safe version** of a production system I built and maintain for real escalation workflows.  
The production version automates weekly reporting, AI classification, and insight generation for multiple internal teams.  

All data here is **synthetic**, and all names or references have been anonymized, but the **structure, logic, and functionality** accurately reflect the real system in use.

---

## 🚀 Overview
This demo automates the end-to-end process of turning raw support ticket data into actionable insights.  
It runs entirely inside Google Sheets via Google Apps Script — no servers, no extra tooling, just automation and AI integration.

It handles:
- Data parsing, validation, and filtering  
- AI-based classification for Issue Types and Root Causes  
- Confidence-based inclusion logic  
- Trend summarization and automated insights  
- Transparent debug reporting for explainability  

---

## ✨ Key Features

✅ **Automated Data Pipeline**  
Processes escalation tickets, normalizes fields, and filters by date and confidence (default 30 days, 60%+).

🤖 **AI or Demo Classification**  
Automatically fills missing Issue Type and Root Cause fields:
- `"mock"` → Generates realistic fake data for demo  
- `"mirror"` → Copies existing manual fields for simulation  
- `"gpt"` → Uses OpenAI API for live classification  

📈 **Aggregated Trend Summaries**  
- Groups by Issue Type and Root Cause  
- Calculates counts and average confidence  
- Outputs readable summary tables  

💡 **Insight Generation**  
- GPT-driven or mock-generated insights summarizing top trends, recurring issues, and process opportunities  

🧾 **Comprehensive Debug Log**  
- One `Debug` tab shows every decision: date filter, confidence pass/fail, and inclusion reason  

---

## 🧠 Example Output

### Trend Summary (Demo)
| Issue Type         | Count | Avg Confidence |
|--------------------|-------|----------------|
| Update Requests    | 50    | 90%            |
| Missing Records    | 46    | 90%            |
| Inventory Gaps     | 34    | 90%            |

| Root Cause         | Count | Avg Confidence |
|--------------------|-------|----------------|
| Data Inconsistency | 42    | 67%            |
| Human Error        | 36    | 68%            |
| System Limitation  | 29    | 66%            |

---

### Insights (Demo Mode)
> - Data entry inconsistencies were the most frequent driver, followed by human mistakes and tool limitations.  
> - Onboarding gaps and incorrect inputs appeared equally often, suggesting clearer guidance could reduce these errors.  
> - Confidence was highest for human error cases, showing they’re easier to identify and prevent.  
> - Tool limitation tickets showed lower confidence, suggesting ambiguity or multiple contributing factors.  

---

## 🛠️ Tech Stack
- **Google Apps Script** – Logic, data handling, GPT integration  
- **Google Sheets** – Input, processing, and dashboard visualization  
- **OpenAI API (optional)** – GPT-generated insights and classification  

---

## ⚙️ Setup Instructions

### 1. Copy the Files
- `apps_script.js` → Paste into a bound Apps Script project inside your Sheet  
- `Trend_Reporting_Sanitized_Demo.xlsx` → Upload to Google Drive and open as a Google Sheet  

### 2. Create Tabs
Ensure your Sheet includes:
- `Project A Raw Import` → Input data  
- `Trend Summary (Demo)` → Summary output  
- `Insights (Demo)` → Insights output  
- `Debug` → Auto-created when the script runs  

### 3. Configure Script Properties
In Apps Script → **Project Settings → Script Properties**, add:
| Property | Example Value | Required |
|-----------|----------------|-----------|
| SHEET_ID | `1abcDExyz123...` | Optional if script is bound |
| OPENAI_API_KEY | `sk-...` | Optional (used for GPT mode) |

---

## ⚙️ Customization

| Setting | Purpose | Example |
|----------|----------|----------|
| `DEMO_FILL_MODE` | Controls how AI fields are filled | `"mock"`, `"mirror"`, `"gpt"` |
| `MOCK_IF_NO_API` | Generates demo insights if no API key is set | `true` |
| `CONF_THRESHOLD` | Minimum confidence to include | `60` |
| `DAYS_BACK` | Days of data to include | `30` |

---

## 📁 Project Structure

ai-trend-reporting-demo/
├── apps_script.js # Main Apps Script logic
├── Trend_Reporting_Sanitized_Demo.xlsx # Demo dataset (synthetic)
├── README.md # Project overview
└── .gitignore # Prevents secrets from being committed



---

## 💡 Why This Matters
This system reduces manual reporting time from hours to minutes, while improving consistency and insight quality.  
It demonstrates:
- **Automation** – Replacing manual data checks with scripted logic  
- **AI Integration** – Augmenting decision-making with GPT predictions  
- **Operational Clarity** – Transparent, reproducible summaries for leadership visibility  

---

## 👤 About the Author
Built by **Tyson Wildman**  
**Associate Supervisor, Escalations Team | DoorDash (Merchant Experience Ops)**  
Focused on building scalable automations, AI tools, and data pipelines that simplify complex workflows and empower teams to focus on higher-impact work.

---

## 🪪 License
MIT License — free for demo, educational, and personal use.
