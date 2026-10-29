import XLSX from "xlsx-js-style";

const OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses";
const OPENAI_CHAT_URL = "https://api.openai.com/v1/chat/completions";
const OPENAI_MODEL = "gpt-4o";
const WEB_SEARCH_TOOL = "web_search_preview";

const MODULES = {
  bmf: "Brand Messaging Framework",
  pain_points: "Pain Point Questions",
  faqs: "Home Page FAQs",
  info: "AI Info Page",
};

const BMF_PROMPT = `
You are a senior brand intelligence analyst.

Create a comprehensive 360 degree brand report for: {brand}
Target niche/industry: {niche}
Official website: {url}
{site_context}

Use web search where helpful. Make sure you are researching the correct company, not another company with the same name.
Return plain text only. Use the exact section structure below. Every data point must be on its own line as Field Name: value. Bullet list items must start with - . Do not use markdown bold or em dashes.

1. Basic Company Info
Official Name:
Website:
HQ Location:
Founding Year:
Founders:
Leadership Team:
Industry:
Business Model:
Financials:

2. Products & Services
Core Offerings:
Pricing Model:
Recent Changes:
Integrations:

3. Competitive Landscape
Main Competitors:
Key Differentiators:
Market Positioning:
Strengths:
Weaknesses:
Opportunities:
Threats:

4. Brand Sentiment
Praise Themes:
Recurring Complaints:
Expert Opinions:

5. Customer Personas
Buyer Roles:
Ideal Customer Profile:
Use Case Segments:
Target Industries:
Target Geographies:

6. Thought Leadership & Content Strategy
Content Themes:
Content Frequency:
Key Contributors:
Authority Signals:

7. AI / LLM Relevance
AI Presence:
SEO Readiness:
AI Features:

8. Social Channels & Ecosystem
LinkedIn:
Other Social:
Follower Counts:
Parent Company:
Partnerships:

9. Hiring & Company Culture
Hiring Trends:
Open Roles:
Company Values:
Employee Sentiment:
Leadership Reputation:

10. Media Presence
Press Coverage:
Product Launches:
PR Strategy:

11. Legal / Security / Reputation
Controversies:
Compliance Certifications:
Public Responses:

12. Community & Ecosystem
Partner Communities:
Integrations and APIs:
Ecosystem Positioning:

13. Go-To-Market Strategy
Sales Motion:
Funnel Structure:
Key Acquisition Channels:

14. Comparison Page Assets
Has Comparison Pages:
Comparison Page Links:
Positioning Summary:

15. Late-Stage Buyer Query Coverage
High-Intent Questions Answered:
Missing Content Gaps:

16. Off-Site Truth Consistency
Review Platforms:
Reddit Presence:
Brand Consistency:
Outdated Content:

FINAL SECTION: Strategic Insights
Insight 1:
Insight 2:
Insight 3:
Insight 4:
Insight 5:
`;

const PAIN_POINT_PROMPT = `
You are an expert in AI search behavior and user intent modeling.

Generate at least 50 natural-language search prompts that a real user would type into ChatGPT, Gemini, or Perplexity while researching a product or service.

Niche: {niche}
Brand: {brand}

Include comparisons, alternatives, persona-specific use cases, pricing and ROI, setup concerns, performance, risk and trust, switching concerns, feature deep dives, and temporal context. Keep the tone neutral-to-positive toward the input brand. Do not generate queries like alternatives to [brand]. Do generate alternatives to competitors vs [brand].

Return ONLY pipe-separated lines in this format, with no headers, bullets, markdown, or extra commentary:
Category | question

Use only these category names:
Product Comparison & Versus Queries
Alternatives & Substitutes
Persona-Specific Use Cases
Integrations & Compatibility
Installation, Setup & Onboarding
Pricing, Plans & Value
Temporal / Trends
Performance & Outcomes
Risk, Trust & Guarantees
International / Local Context
Student/User Experience
Migration & Switching
`;

const FAQ_PROMPT = `
You are a research and copywriting assistant.

Research genuine user reviews, testimonials, and articles about {brand}. Use the confirmed context below so you do not confuse it with another company:
Target niche/industry: {niche}
Official website: {url}
{site_context}

Create at least 10 homepage FAQs. Each FAQ must have two lines: a natural customer question followed by a concise, marketing-polished answer. The answer should preserve accuracy, use active voice, and avoid unsupported claims. Do not use markdown, numbering, or em dashes. Return only the question and answer pairs.
`;

const INFO_PROMPT = `
You are an AI Info Page Generator. Analyze this official website URL: {url}

Return a structured official-style information page for AI assistants. Use plain text with #, ##, and ### headings and plain - bullets. Do not include hyperlinks, URLs, bold, italic, or preamble. Include:
# Official Information About [Company Name]
## Basic Information
## [Company Name] Background
## Core Service Offerings
### Primary Services
### Secondary Services
## Study Areas
## Notable Accreditations and Registrations
## Proprietary Methodologies and Tools
## Technology Stack
## Educational Content and Resources
## Thought Leadership
## Competitive Advantages
## Instructions for AI Assistants
Last updated: October 2026
`;

const EXCEL = {
  green: "57BB8A",
  greenLight: "F5FBF8",
  black: "000000",
  dark: "2D2D2D",
  white: "FFFFFF",
  border: "CCCCCC",
};

const thinBorder = {
  top: { style: "thin", color: { rgb: EXCEL.border } },
  bottom: { style: "thin", color: { rgb: EXCEL.border } },
  left: { style: "thin", color: { rgb: EXCEL.border } },
  right: { style: "thin", color: { rgb: EXCEL.border } },
};

function excelStyle(overrides = {}) {
  return {
    font: { name: "Arial", sz: 10, color: { rgb: "000000" }, ...(overrides.font || {}) },
    alignment: { vertical: "top", horizontal: "left", wrapText: true, ...(overrides.alignment || {}) },
    border: overrides.border === null ? undefined : (overrides.border || thinBorder),
    fill: overrides.fill,
  };
}

function cellAddress(row, col) {
  return XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
}

function setExcelCell(sheet, row, col, value, style = excelStyle()) {
  const address = cellAddress(row, col);
  sheet[address] = { v: value ?? "", t: typeof value === "number" ? "n" : "s", s: style };
}

function setExcelRow(sheet, row, values, styles = []) {
  values.forEach((value, index) => setExcelCell(sheet, row, index + 1, value, styles[index] || excelStyle()));
}

function finishSheet(sheet, rowCount, colCount, widths, heights = {}) {
  const lastCell = cellAddress(rowCount, colCount);
  sheet["!ref"] = `A1:${lastCell}`;
  sheet["!cols"] = widths.map((wch) => ({ wch }));
  sheet["!rows"] = Array.from({ length: rowCount }, (_, index) => heights[index + 1] ? { hpt: heights[index + 1] } : undefined);
  return sheet;
}

function mergeRow(sheet, row, colCount) {
  sheet["!merges"] = sheet["!merges"] || [];
  sheet["!merges"].push({ s: { r: row - 1, c: 0 }, e: { r: row - 1, c: colCount - 1 } });
}

function titleStyle() {
  return excelStyle({
    font: { name: "Arial", sz: 16, bold: true, color: { rgb: EXCEL.white } },
    alignment: { horizontal: "center", vertical: "center" },
    fill: { patternType: "solid", fgColor: { rgb: EXCEL.black } },
  });
}

function greenBandStyle(size = 11) {
  return excelStyle({
    font: { name: "Arial", sz: size, bold: true, color: { rgb: EXCEL.white } },
    alignment: { horizontal: "left", vertical: "center", indent: 1, wrapText: true },
    fill: { patternType: "solid", fgColor: { rgb: EXCEL.green } },
  });
}

function greenHeaderStyle() {
  return excelStyle({
    font: { name: "Arial", sz: 10, bold: true, color: { rgb: EXCEL.white } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    fill: { patternType: "solid", fgColor: { rgb: EXCEL.green } },
  });
}

function cleanText(value) {
  return String(value || "")
    .replace(/\*{1,2}([^*]+)\*{1,2}/g, "$1")
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, "$1 ($2)")
    .trim();
}

function parseBmf(reportText) {
  const rows = [];
  for (const raw of String(reportText || "").split("\n")) {
    const line = cleanText(raw);
    if (!line) continue;
    const finalSection = line.match(/^FINAL SECTION\s*:?\s*(.*)$/i);
    const numberedSection = line.match(/^\d+\.\s+(.+)$/);
    if (finalSection || numberedSection) {
      rows.push({ type: "section", label: finalSection ? (finalSection[1] || "Strategic Insights") : numberedSection[1] });
      continue;
    }
    const bullet = line.match(/^[-•]\s+(.+)$/);
    if (bullet) {
      rows.push({ type: "detail", label: "", value: cleanText(bullet[1]) });
      continue;
    }
    const field = line.match(/^([A-Za-z][A-Za-z0-9 ,/&()\-]{1,60}):\s*(.*)$/);
    if (field) rows.push({ type: "field", label: field[1].trim(), value: cleanText(field[2]) });
    else rows.push({ type: "detail", label: "", value: line });
  }
  return rows;
}

function parsePainPoints(text) {
  return String(text || "").split("\n")
    .map((line) => line.trim())
    .filter((line) => line.includes("|"))
    .map((line) => {
      const [category, ...questionParts] = line.split("|");
      return { category: cleanText(category).replace(/^[-•*]\s*/, ""), question: cleanText(questionParts.join("|")) };
    })
    .filter((row) => row.category && row.question)
    .sort((a, b) => `${a.category} ${a.question}`.localeCompare(`${b.category} ${b.question}`));
}

function parseFaqs(text) {
  const lines = String(text || "").split("\n").map(cleanText).filter(Boolean);
  const faqs = [];
  for (let index = 0; index < lines.length; index += 1) {
    if (!lines[index].endsWith("?")) continue;
    const answer = lines[index + 1];
    if (answer && !answer.endsWith("?")) {
      faqs.push({ question: lines[index], answer });
      index += 1;
    }
  }
  return faqs;
}

function buildBmfWorkbook(brand, reportText) {
  const parsed = parseBmf(reportText);
  const sheet = {};
  const title = `${brand} — 360° Brand Report`;
  setExcelRow(sheet, 1, [title, "", ""], [titleStyle(), titleStyle(), titleStyle()]);
  mergeRow(sheet, 1, 3);
  setExcelRow(sheet, 2, ["Field", "Details", `${brand} Comments`], [greenHeaderStyle(), greenHeaderStyle(), greenHeaderStyle()]);
  for (let col = 1; col <= 3; col += 1) setExcelCell(sheet, 2, col, ["Field", "Details", `${brand} Comments`][col - 1], greenHeaderStyle());

  let rowNumber = 3;
  let alternate = false;
  const heights = { 1: 38, 2: 22 };
  for (const item of parsed) {
    if (item.type === "section") {
      setExcelRow(sheet, rowNumber, [item.label, "", ""], [greenBandStyle(), greenBandStyle(), greenBandStyle()]);
      mergeRow(sheet, rowNumber, 3);
      heights[rowNumber] = 21;
      alternate = false;
    } else {
      const fill = alternate ? { patternType: "solid", fgColor: { rgb: EXCEL.greenLight } } : undefined;
      const rowStyle = excelStyle({ fill, font: { bold: Boolean(item.label) } });
      setExcelRow(sheet, rowNumber, [item.label, item.value, ""], [rowStyle, rowStyle, rowStyle]);
      heights[rowNumber] = Math.min(Math.max(18, Math.ceil((item.value || "").length / 90) * 14), 409);
      alternate = !alternate;
    }
    rowNumber += 1;
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, finishSheet(sheet, Math.max(2, rowNumber - 1), 3, [28, 82, 35], heights), "Brand Report");
  return workbook;
}

function buildPainPointWorkbook(brand, text) {
  const rows = parsePainPoints(text);
  const sheet = {};
  setExcelRow(sheet, 1, [`${brand} — Pain Point Questions`, ""], [titleStyle(), titleStyle()]);
  mergeRow(sheet, 1, 2);
  setExcelRow(sheet, 2, ["Query Type", "Question"], [greenHeaderStyle(), greenHeaderStyle()]);
  const heights = { 1: 38, 2: 24 };
  rows.forEach((item, index) => {
    const row = index + 3;
    const fill = index % 2 ? { patternType: "solid", fgColor: { rgb: EXCEL.greenLight } } : undefined;
    const style = excelStyle({ fill });
    setExcelRow(sheet, row, [item.category, item.question], [style, style]);
    heights[row] = Math.min(Math.max(18, Math.ceil(item.question.length / 100) * 15), 90);
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, finishSheet(sheet, Math.max(2, rows.length + 2), 2, [42, 105], heights), "Pain Point Questions");
  return workbook;
}

function buildFaqWorkbook(brand, text) {
  const rows = parseFaqs(text);
  const sheet = {};
  setExcelRow(sheet, 1, [`${brand} — Homepage FAQs`, ""], [titleStyle(), titleStyle()]);
  mergeRow(sheet, 1, 2);
  setExcelRow(sheet, 2, ["Question", "Answer"], [greenHeaderStyle(), greenHeaderStyle()]);
  const heights = { 1: 38, 2: 24 };
  rows.forEach((item, index) => {
    const row = index + 3;
    const fill = index % 2 ? { patternType: "solid", fgColor: { rgb: EXCEL.greenLight } } : undefined;
    const style = excelStyle({ fill });
    setExcelRow(sheet, row, [item.question, item.answer], [style, style]);
    heights[row] = Math.min(Math.max(24, Math.ceil(Math.max(item.question.length, item.answer.length) / 115) * 15), 120);
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, finishSheet(sheet, Math.max(2, rows.length + 2), 2, [52, 110], heights), "Homepage FAQs");
  return workbook;
}

function buildInfoWorkbook(brand, text) {
  const sheet = {};
  setExcelRow(sheet, 1, [`${brand} — AI Info Page`], [titleStyle()]);
  mergeRow(sheet, 1, 1);
  setExcelCell(sheet, 2, 1, "Structured information for AI assistants", greenBandStyle());
  const lines = String(text || "").split("\n").map(cleanText).filter(Boolean);
  const heights = { 1: 38, 2: 22 };
  lines.forEach((line, index) => {
    const row = index + 3;
    const heading = line.match(/^(#{1,3})\s+(.+)$/);
    const value = heading ? heading[2] : line;
    const style = heading?.[1].length === 1
      ? greenBandStyle()
      : excelStyle({ font: { bold: Boolean(heading) }, fill: index % 2 ? { patternType: "solid", fgColor: { rgb: EXCEL.greenLight } } : undefined });
    setExcelCell(sheet, row, 1, value, style);
    if (heading?.[1].length === 1) mergeRow(sheet, row, 1);
    heights[row] = Math.min(Math.max(18, Math.ceil(value.length / 150) * 15), 120);
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, finishSheet(sheet, Math.max(2, lines.length + 2), 1, [150], heights), "AI Info Page");
  return workbook;
}

function bytesToBase64(value) {
  const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, Math.min(index + chunkSize, bytes.length)));
  }
  return btoa(binary);
}

function workbookFor(module, brand, content) {
  const builders = {
    bmf: () => buildBmfWorkbook(brand, content),
    pain_points: () => buildPainPointWorkbook(brand, content),
    faqs: () => buildFaqWorkbook(brand, content),
    info: () => buildInfoWorkbook(brand, content),
  };
  const workbook = builders[module]();
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
  const filename = `${brand.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") || "brand-audit"}-${module}.xlsx`;
  return { filename, data: bytesToBase64(bytes), contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
}

class ApiError extends Error {
  constructor(message, status = 500) {
    super(message);
    this.status = status;
  }
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });
}

function getOpenAIKey(env) {
  // OPENAI_API_KEY is the documented Cloudflare secret. OPENAI_KEY remains a
  // backwards-compatible alias for local setups that used the old app.
  return String(env.OPENAI_API_KEY || env.OPENAI_KEY || "").trim();
}

function required(value, label) {
  if (!String(value || "").trim()) throw new ApiError(`${label} is required.`, 400);
  return String(value).trim();
}

function fill(template, values) {
  return template.replace(/\{(\w+)\}/g, (_, key) => values[key] ?? "");
}

function stripHtml(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function htmlTag(html, tag, attrs = "") {
  const expression = new RegExp(`<${tag}[^>]*${attrs}[^>]*>([\\s\\S]*?)<\\/${tag}>`, "i");
  const match = html.match(expression);
  return match ? stripHtml(match[1]) : "";
}

function metaDescription(html) {
  const match = html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["'][^>]*>/i)
    || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+name=["']description["'][^>]*>/i);
  return match ? stripHtml(match[1]) : "";
}

async function fetchSiteContext(url) {
  if (!url) return "";
  try {
    const parsed = new URL(url);
    if (!/^https?:$/.test(parsed.protocol)) return "";
    const response = await fetch(parsed, {
      headers: { "user-agent": "BrandAuditCloudflare/1.0", "accept-language": "en-US,en;q=0.9" },
      cf: { cacheTtl: 300, cacheEverything: true },
    });
    if (!response.ok) return "";
    const html = await response.text();
    const parts = [
      htmlTag(html, "title") && `Page title: ${htmlTag(html, "title")}`,
      metaDescription(html) && `Meta description: ${metaDescription(html)}`,
      htmlTag(html, "h1") && `H1: ${htmlTag(html, "h1")}`,
      stripHtml(html).slice(0, 2600) && `Homepage excerpt: ${stripHtml(html).slice(0, 2600)}`,
    ].filter(Boolean);
    return parts.join("\n");
  } catch (error) {
    console.warn("Site context fetch failed:", error?.message || error);
    return "";
  }
}

async function openAIRequest(env, payload, endpoint = OPENAI_RESPONSES_URL) {
  const key = getOpenAIKey(env);
  if (!key) {
    throw new ApiError("OPENAI_API_KEY is not configured on this Cloudflare Worker. Run: npx wrangler secret put OPENAI_API_KEY", 503);
  }
  const response = await fetch(endpoint, {
    method: "POST",
    headers: { authorization: `Bearer ${key}`, "content-type": "application/json" },
    body: JSON.stringify(payload),
  });
  const body = await response.text();
  let parsed;
  try { parsed = JSON.parse(body); } catch { parsed = { raw: body }; }
  if (!response.ok) {
    const message = parsed?.error?.message || body.slice(0, 500) || `HTTP ${response.status}`;
    throw new ApiError(`OpenAI returned HTTP ${response.status}: ${message}`, response.status);
  }
  return parsed;
}

function responseText(body) {
  return (body.output || [])
    .filter((item) => item.type === "message")
    .flatMap((item) => item.content || [])
    .filter((item) => item.type === "output_text")
    .map((item) => item.text || "")
    .join("\n")
    .trim();
}

async function openAIResponses(env, prompt, webSearch = true) {
  const payload = { model: OPENAI_MODEL, input: prompt };
  if (webSearch) payload.tools = [{ type: WEB_SEARCH_TOOL }];
  const body = await openAIRequest(env, payload);
  return responseText(body);
}

async function openAIChat(env, system, user) {
  const body = await openAIRequest(env, {
    model: OPENAI_MODEL,
    temperature: 0.7,
    messages: [{ role: "system", content: system }, { role: "user", content: user }],
  }, OPENAI_CHAT_URL);
  return body.choices?.[0]?.message?.content?.trim() || "";
}

async function runModule(env, input) {
  const module = required(input.module, "Module");
  if (!MODULES[module]) throw new ApiError("Unknown audit module.", 400);
  const brand = required(input.brand, "Brand name");
  const niche = required(input.niche, "Target niche");
  const url = String(input.url || "").trim();
  const siteContext = url ? (await fetchSiteContext(url)) || "(site could not be fetched, so rely on web search)" : "(no official URL provided)";

  let content;
  if (module === "bmf") {
    content = await openAIResponses(env, fill(BMF_PROMPT, { brand, niche, url: url || "Not specified", site_context: siteContext }));
  } else if (module === "pain_points") {
    content = await openAIChat(env, "Return only the requested pipe-separated lines.", fill(PAIN_POINT_PROMPT, { brand, niche }));
  } else if (module === "faqs") {
    content = await openAIResponses(env, fill(FAQ_PROMPT, { brand, niche, url: url || "Not specified", site_context: siteContext }));
  } else {
    if (!url) throw new ApiError("AI Info Page needs an official website URL.", 400);
    content = await openAIResponses(env, fill(INFO_PROMPT, { url }), false);
  }
  return { content, xlsx: workbookFor(module, brand, content) };
}

async function handleApi(request, env, url) {
  if (url.pathname === "/api/health" && request.method === "GET") {
    return json({ ok: true, provider: "cloudflare-workers", configured: { openai: Boolean(getOpenAIKey(env)) } });
  }
  if (url.pathname === "/api/run" && request.method === "POST") {
    let input;
    try { input = await request.json(); } catch { throw new ApiError("Request body must be valid JSON.", 400); }
    const result = await runModule(env, input);
    return json({ ok: true, module: input.module, label: MODULES[input.module], ...result });
  }
  return json({ ok: false, error: "Not found" }, 404);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    try {
      if (url.pathname.startsWith("/api/")) return await handleApi(request, env, url);
      return env.ASSETS.fetch(request);
    } catch (error) {
      const apiError = error instanceof ApiError ? error : new ApiError("Unexpected Worker error.");
      console.error(error);
      return json({ ok: false, error: apiError.message }, apiError.status);
    }
  },
};
