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

  if (module === "bmf") {
    return openAIResponses(env, fill(BMF_PROMPT, { brand, niche, url: url || "Not specified", site_context: siteContext }));
  }
  if (module === "pain_points") {
    return openAIChat(env, "Return only the requested pipe-separated lines.", fill(PAIN_POINT_PROMPT, { brand, niche }));
  }
  if (module === "faqs") {
    return openAIResponses(env, fill(FAQ_PROMPT, { brand, niche, url: url || "Not specified", site_context: siteContext }));
  }
  if (!url) throw new ApiError("AI Info Page needs an official website URL.", 400);
  return openAIResponses(env, fill(INFO_PROMPT, { url }), false);
}

async function handleApi(request, env, url) {
  if (url.pathname === "/api/health" && request.method === "GET") {
    return json({ ok: true, provider: "cloudflare-workers", configured: { openai: Boolean(getOpenAIKey(env)) } });
  }
  if (url.pathname === "/api/run" && request.method === "POST") {
    let input;
    try { input = await request.json(); } catch { throw new ApiError("Request body must be valid JSON.", 400); }
    const content = await runModule(env, input);
    return json({ ok: true, module: input.module, label: MODULES[input.module], content });
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
