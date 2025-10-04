/***********************
 * Escalation Trend Pipeline (Sanitized Demo)
 * Google Apps Script
 *
 * Tabs expected:
 *   - "Raw Import"
 *   - "Trend Summary (Demo)"
 *   - "Insights (Demo)"
 *   - "Debug" is created automatically
 *
 * Script Properties (Project Settings):
 *   - SHEET_ID = target Sheet ID (optional if script is bound to the Sheet)
 *   - OPENAI_API_KEY = your key (optional)
 ***********************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Escalation Tools")
    .addItem("Run Trend Pipeline", "mainTrendPipeline")
    .addToUi();
}

/* =====================
    Global Config
===================== */

// Confidence threshold for inclusion
const CONF_THRESHOLD = 60;

// Filter window
const DAYS_BACK = 30;

// Insights behavior when no API key is present
const MOCK_IF_NO_API = true;

// How to fill empty AI columns before summarizing:
// "mirror" copies from SOURCE_* fields with 90 conf fallback
// "mock" generates demo values and confidences
// "gpt" calls OpenAI if OPENAI_API_KEY is set, otherwise falls back to mock
const DEMO_FILL_MODE = "mock";

// Source fields if using "mirror" mode
const SOURCE_ISSUE_FIELD = "Issue Type (MNSD)";    // adjust if needed
const SOURCE_CAUSE_FIELD = "Root Cause";           // adjust if you have a non-AI root cause column

// Tabs
const TICKETS_TAB = "Raw Import";
const SUMMARY_TAB = "Trend Summary (Demo)";
const INSIGHTS_TAB = "Insights (Demo)";
const DEBUG_TAB = "Debug";

// Allowed categories for demo fills
const ISSUE_TYPES_DEMO = [
  "Update Requests","Missing Records","Inventory - Tracking","Image Adjustments","Policy - Config",
  "Taxonomy Update","Photo Update","Item Onboarding","General Inquiry","Other"
];
const ROOT_CAUSES_DEMO = [
  "Data Inconsistency","Human Error","System Limitation","Onboarding Gaps","Incorrect Input","Other"
];

/* =====================
    Entry Point
===================== */

function mainTrendPipeline() {
  const props = PropertiesService.getScriptProperties();
  const SHEET_ID = props.getProperty("SHEET_ID");
  const OPENAI_API_KEY = props.getProperty("OPENAI_API_KEY");

  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Could not open spreadsheet. Check SHEET_ID or bind the script.");

  // Ensure AI columns exist and populate empty cells using the selected mode
  ensureAIFieldsPopulated(ss, TICKETS_TAB, OPENAI_API_KEY);

  const ticketsSheet = ss.getSheetByName(TICKETS_TAB);
  if (!ticketsSheet) throw new Error("Missing tab: " + TICKETS_TAB);

  const data = getSheetDataAsObjects(ticketsSheet);
  const debugRows = [];

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - DAYS_BACK);

  // Filter to last N days and build a rich debug row for every ticket
  const recentQualified = data.filter((row, i) => {
    const created = parseCreatedDate(row["Created"]);
    const inWindow = !isNaN(created) && created >= cutoffDate;

    const issueType = row["AI Issue Type"];
    const issueConf = safeParseConfidence(row["AI Issue Type Confidence"]);
    const rootCause = row["AI Root Cause"];
    const rootConf = safeParseConfidence(row["AI Root Cause Confidence"]);

    let reason = "✅";
    if (issueConf === null && rootConf === null) reason = "❌ Missing both confidences";
    else if (issueConf === null) reason = "❌ Invalid issue confidence";
    else if (rootConf === null) reason = "❌ Invalid root confidence";
    else if (issueConf < CONF_THRESHOLD && rootConf < CONF_THRESHOLD) reason = "❌ Both below threshold";
    else if (issueConf < CONF_THRESHOLD) reason = "❌ Issue confidence below threshold";
    else if (rootConf < CONF_THRESHOLD) reason = "❌ Root confidence below threshold";

    debugRows.push([
      i + 2,
      row["Created"],
      !isNaN(created) ? created.toISOString() : "Invalid",
      row["Summary"] || "",
      issueType || "",
      issueConf,
      rootCause || "",
      rootConf,
      inWindow ? "Recent" : "Old",
      reason
    ]);

    return inWindow && reason === "✅";
  });

  // Summaries
  const issueSummary = summarizeByKey(recentQualified, "AI Issue Type", "AI Issue Type Confidence");
  const causeSummary = summarizeByKey(recentQualified, "AI Root Cause", "AI Root Cause Confidence");

  // Output sheets
  writeSummarySheet(ss, SUMMARY_TAB, issueSummary, causeSummary, cutoffDate, new Date(), CONF_THRESHOLD);

  // Insights: live if key present, otherwise mock if allowed
  let insightText;
  if (OPENAI_API_KEY) {
    insightText = getGPTInsights(causeSummary, OPENAI_API_KEY);
  } else if (MOCK_IF_NO_API) {
    insightText =
      "- Data entry inconsistencies are the most frequent driver, followed by human mistakes and tool limitations.\n" +
      "- Onboarding gaps and incorrect inputs appear at comparable rates, suggesting clearer guidance is needed.\n" +
      "- Confidence is typically higher on human error cases, indicating they are easier to identify and prevent.\n" +
      "- Tool limitation tickets show lower confidence, implying ambiguity or multi-causal issues.\n" +
      "- Focus areas for improvement include data hygiene, onboarding guidance, and targeted fixes to tooling.";
  } else {
    insightText = "⚠️ No API key found in Script Properties (OPENAI_API_KEY).";
  }
  writeInsightsSheet(ss, INSIGHTS_TAB, insightText);

  // Single debug tab with the richest detail
  writeSingleDebug(ss, DEBUG_TAB, debugRows);
}

/* =====================
    Fill Empty AI Columns
===================== */

function ensureAIFieldsPopulated(ss, ticketsTabName, apiKey) {
  const sheet = ss.getSheetByName(ticketsTabName);
  if (!sheet) throw new Error("Missing tab: " + ticketsTabName);

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return;

  const headers = values[0];
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  // Ensure target columns exist
  const required = [
    "AI Issue Type",
    "AI Issue Type Confidence",
    "AI Root Cause",
    "AI Root Cause Confidence"
  ];
  required.forEach(col => {
    if (!(col in idx)) {
      headers.push(col);
      values[0] = headers;
      idx[col] = headers.length - 1;
      if (sheet.getLastColumn() < headers.length) {
        sheet.insertColumnsAfter(sheet.getLastColumn(), headers.length - sheet.getLastColumn());
      }
    }
  });

  // Helpers
  const randPick = arr => arr[Math.floor(Math.random() * arr.length)];
  const randConf = () => Math.max(30, Math.min(99, Math.round(65 + (Math.random() - 0.5) * 20)));

  // Fill rows
  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    let aiIssue = row[idx["AI Issue Type"]];
    let aiIssueConf = row[idx["AI Issue Type Confidence"]];
    let aiCause = row[idx["AI Root Cause"]];
    let aiCauseConf = row[idx["AI Root Cause Confidence"]];

    const needsIssue = aiIssue === "" || aiIssue === null;
    const needsCause = aiCause === "" || aiCause === null;
    if (!needsIssue && !needsCause) continue;

    if (DEMO_FILL_MODE === "mirror") {
      const srcIssue = idx.hasOwnProperty(SOURCE_ISSUE_FIELD) ? row[idx[SOURCE_ISSUE_FIELD]] : "";
      const srcCause = idx.hasOwnProperty(SOURCE_CAUSE_FIELD) ? row[idx[SOURCE_CAUSE_FIELD]] : "";
      if (needsIssue) {
        row[idx["AI Issue Type"]] = srcIssue || randPick(ISSUE_TYPES_DEMO);
        row[idx["AI Issue Type Confidence"]] = aiIssueConf || 90;
      }
      if (needsCause) {
        row[idx["AI Root Cause"]] = srcCause || randPick(ROOT_CAUSES_DEMO);
        row[idx["AI Root Cause Confidence"]] = aiCauseConf || 90;
      }
    } else if (DEMO_FILL_MODE === "mock") {
      if (needsIssue) {
        row[idx["AI Issue Type"]] = randPick(ISSUE_TYPES_DEMO);
        row[idx["AI Issue Type Confidence"]] = randConf();
      }
      if (needsCause) {
        row[idx["AI Root Cause"]] = randPick(ROOT_CAUSES_DEMO);
        row[idx["AI Root Cause Confidence"]] = randConf();
      }
    } else if (DEMO_FILL_MODE === "gpt") {
      if (!apiKey) {
        // Fallback to mock if no API key
        if (needsIssue) {
          row[idx["AI Issue Type"]] = randPick(ISSUE_TYPES_DEMO);
          row[idx["AI Issue Type Confidence"]] = randConf();
        }
        if (needsCause) {
          row[idx["AI Root Cause"]] = randPick(ROOT_CAUSES_DEMO);
          row[idx["AI Root Cause Confidence"]] = randConf();
        }
      } else {
        const summary = idx["Summary"] != null ? row[idx["Summary"]] : "";
        const description = idx["Description"] != null ? row[idx["Description"]] : "";
        const pred = predictIssueAndCauseViaGPT_(summary, description, apiKey);
        if (needsIssue) {
          row[idx["AI Issue Type"]] = (pred && pred.issueType) || randPick(ISSUE_TYPES_DEMO);
          row[idx["AI Issue Type Confidence"]] =
            pred && pred.issueConfidence != null ? Math.round(pred.issueConfidence * 100) : randConf();
        }
        if (needsCause) {
          row[idx["AI Root Cause"]] = (pred && pred.rootCause) || randPick(ROOT_CAUSES_DEMO);
          row[idx["AI Root Cause Confidence"]] =
            pred && pred.rootConfidence != null ? Math.round(pred.rootConfidence * 100) : randConf();
        }
      }
    }

    values[r] = row;
  }

  // Commit once
  range.setValues(values);
}

/* =====================
    GPT Helpers
===================== */

function predictIssueAndCauseViaGPT_(summary, description, apiKey) {
  try {
    const prompt = `You are a support triage AI. From the text, output JSON with fields:
{
  "issueType": "one of: ${ISSUE_TYPES_DEMO.join(", ")}",
  "issueConfidence": 0 to 1,
  "rootCause": "one of: ${ROOT_CAUSES_DEMO.join(", ")}",
  "rootConfidence": 0 to 1
}

Summary: ${summary}
Description: ${description}`;

    const res = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        temperature: 0
      })
    });
    const body = JSON.parse(res.getContentText());
    const text = body.choices[0].message.content || "";
    return JSON.parse(text);
  } catch (e) {
    Logger.log("predictIssueAndCauseViaGPT_ error: " + e);
    return null;
  }
}

/* =====================
    Summaries and Output
===================== */

function getSheetDataAsObjects(sheet) {
  const range = sheet.getDataRange().getValues();
  if (!range || range.length < 2) return [];
  const headers = range[0];
  return range.slice(1).map(row =>
    Object.fromEntries(headers.map((key, i) => [key, row[i]]))
  );
}

function parseCreatedDate(raw) {
  if (raw instanceof Date && !isNaN(raw)) return raw;

  if (typeof raw === "number") {
    const d = new Date(Math.round((raw - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d;
  }

  if (typeof raw !== "string") return new Date("Invalid");
  const s = raw.trim();

  const iso = new Date(s);
  if (!isNaN(iso)) return iso;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return new Date("Invalid");
  const mm = parseInt(m[1], 10);
  const dd = parseInt(m[2], 10);
  const yyyy = parseInt(m[3], 10);
  const hh = parseInt(m[4], 10);
  const min = parseInt(m[5], 10);
  const ss = parseInt(m[6] || "0", 10);
  return new Date(yyyy, mm - 1, dd, hh, min, ss);
}

function safeParseConfidence(raw) {
  if (raw === null || raw === undefined || raw === "") return null;

  if (typeof raw === "number") {
    if (isNaN(raw)) return null;
    return raw <= 1 ? raw * 100 : raw;
  }

  const cleaned = String(raw).trim().replace(/[^0-9.]/g, "");
  if (!cleaned) return null;

  const num = parseFloat(cleaned);
  if (isNaN(num)) return null;

  return num <= 1 ? num * 100 : num;
}

function summarizeByKey(data, keyField, confidenceField) {
  const map = {};
  data.forEach(row => {
    const key = row[keyField] || "Unknown";
    const conf = safeParseConfidence(row[confidenceField]) ?? 0;
    if (!map[key]) map[key] = { count: 0, total: 0 };
    map[key].count += 1;
    map[key].total += conf;
  });
  return Object.entries(map).map(([key, stats]) => ({
    key,
    count: stats.count,
    avg: Number((stats.total / stats.count).toFixed(1))
  })).sort((a, b) => b.count - a.count).slice(0, 5);
}

function writeSummarySheet(ss, tabName, issueSummary, causeSummary, startDate, endDate, threshold) {
  let tab = ss.getSheetByName(tabName);
  if (!tab) tab = ss.insertSheet(tabName);
  else tab.clear();

  const title = `📊 Escalation Trends (${startDate.toDateString()} to ${endDate.toDateString()})`;
  tab.getRange("A1").setValue(title);

  if (issueSummary.length === 0 && causeSummary.length === 0) {
    tab.getRange("A3").setValue(`⚠️ No qualifying tickets passed confidence threshold (≥${threshold}%)`);
    return;
  }

  const issueData = [["Issue Type", "Count", "Avg Confidence (%)"]].concat(
    issueSummary.map(r => [r.key, r.count, r.avg])
  );
  tab.getRange(3, 1, issueData.length, 3).setValues(issueData);

  const rootStart = issueData.length + 4;
  const causeData = [["Root Cause", "Count", "Avg Confidence (%)"]].concat(
    causeSummary.map(r => [r.key, r.count, r.avg])
  );
  tab.getRange(rootStart, 1, causeData.length, 3).setValues(causeData);
}

function getGPTInsights(causeSummary, apiKey) {
  const bullets = causeSummary.map(r => `${r.key}: ${r.count} tickets, avg confidence ${r.avg}%`).join("\n");
  const prompt = `You are summarizing support escalation trends.

Root causes from the past 30 days:
${bullets}

Provide 3 to 5 concise bullets suitable for a weekly business review:
- Process gaps
- Suggested improvements
- Actionable follow ups`;

  try {
    const res = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.5
      })
    });
    return JSON.parse(res.getContentText()).choices[0].message.content.trim();
  } catch (e) {
    return "⚠️ GPT insights failed: " + e;
  }
}

function writeInsightsSheet(ss, tabName, text) {
  let tab = ss.getSheetByName(tabName);
  if (!tab) tab = ss.insertSheet(tabName);
  else tab.clear();

  const lines = [["Root Cause Insights (GPT Generated)"]];
  if (!text || typeof text !== "string") {
    lines.push(["⚠️ GPT returned no insight."]);
  } else {
    text.split("\n").forEach(line => lines.push([line.trim()]));
  }
  tab.getRange(1, 1, lines.length, 1).setValues(lines);
}

function writeSingleDebug(ss, tabName, rows) {
  let tab = ss.getSheetByName(tabName);
  if (!tab) tab = ss.insertSheet(tabName);
  else tab.clear();

  const header = [
    "Row #",
    "Raw Created",
    "Parsed ISO",
    "Summary",
    "Issue Type",
    "Issue Conf",
    "Root Cause",
    "Root Conf",
    "Date Window",
    "Confidence Filter"
  ];
  const output = [header].concat(rows);
  tab.getRange(1, 1, output.length, output[0].length).setValues(output);
}
