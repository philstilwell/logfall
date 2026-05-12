import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const dataPath = path.join(projectRoot, "data", "fallacies.json");
const siteRoot = path.join(projectRoot, "site");
const distRoot = projectRoot;
const dataOutDir = path.join(distRoot, "data");
const workbookOutPath = path.join(distRoot, "logfall-root-edition.xlsx");
const siteUrl = "https://logfall.com/";
const socialImagePath = "assets/logo.jpg";
const socialImageType = "image/jpeg";
const socialImageWidth = 124;
const socialImageHeight = 124;
const buildDate = new Date().toISOString().split("T")[0];
const copyrightNotice = "Copyright © Phil Stilwell";
const cloudflareWebAnalyticsTag =
  `<!-- Cloudflare Web Analytics --><script defer src='https://static.cloudflareinsights.com/beacon.min.js' data-cf-beacon='{"token": "798a30777a6d424f9c4055a02e7bde91"}'></script><!-- End Cloudflare Web Analytics -->`;
const baseSiteKeywords = [
  "logical fallacies",
  "critical thinking",
  "argument analysis",
  "reasoning errors",
  "logic",
];

const featuredNames = [
  "Ad hominem",
  "Appeal to authority",
  "Base rate fallacy",
  "Begging the question",
  "Cherry picking",
  "Correlation is not causation",
  "Equivocation",
  "False dilemma",
  "Moving the goalpost",
  "No True Scotsman",
  "Slippery slope",
  "Straw man argument",
];

const categoryDescriptions = {
  Formal: "Breakdowns in deductive structure where the conclusion does not follow from the form.",
  Mathematical: "Missteps involving probability, statistics, scope, quantity, or numerical expectations.",
  Causal: "Faulty claims about what caused what, or what causal link has actually been shown.",
  Linguistic: "Confusion created by wording, ambiguity, framing, or unstable definitions.",
  Conceptual: "Errors caused by bad categories, weak distinctions, or distorted conceptual boundaries.",
  Evidential: "Arguments that overstate what the evidence shows, ignore what is missing, or misuse support.",
  Perceptual: "Mistakes rooted in appearances, impressions, or the way something seems at first glance.",
  Perspectival: "Errors caused by the wrong vantage point, historical standpoint, or interpretive frame.",
  Epistemic: "Failures in belief management, confidence calibration, or standards for responsible belief.",
  Tactical: "Debate maneuvers that distract, derail, pressure, or strategically reroute the exchange.",
  Emotional: "Arguments that make feeling do the evidential work reasoning should have done.",
};

const diagnosticPrompts = {
  Formal: "If the premises were true, would the conclusion still fail to follow?",
  Mathematical: "What numbers, rates, or probabilities are being ignored or mishandled?",
  Causal: "What evidence actually rules out coincidence, reverse causation, or a third factor?",
  Linguistic: "Has the wording shifted, blurred, or changed meaning mid-argument?",
  Conceptual: "Are the categories being used carefully, or are unlike things being treated as alike?",
  Evidential: "What evidence is missing, selected, or overstretched here?",
  Perceptual: "Is this conclusion being drawn from how things seem rather than what has been shown?",
  Perspectival: "Would the conclusion change if the frame, timeline, or viewpoint were widened?",
  Epistemic: "Is the speaker calibrating confidence to the strength of the evidence?",
  Tactical: "Is the argument still addressing the original issue, or has the conversation been steered away?",
  Emotional: "Would the argument still persuade if the emotional force were removed?",
};

const classroomLevels = ["Middle school+", "High school", "Intro college", "Advanced undergraduate"];
const teachingPathDefinitions = [
  {
    slug: "start-here",
    title: "Start here",
    description: "A foundational sequence for first-time readers and classrooms starting with the fallacies they are most likely to meet in ordinary discussion.",
    audience: "Best for middle school through intro college.",
    names: [
      "Ad hominem",
      "Appeal to authority",
      "Appeal to emotion",
      "Begging the question",
      "Cherry picking",
      "Correlation is not causation",
      "False dilemma",
      "Hasty generalization",
      "No True Scotsman",
      "Slippery slope",
      "Straw man argument",
      "Tu quoque",
    ],
  },
  {
    slug: "public-debate",
    title: "Most common in public debate",
    description: "A classroom-ready path centered on the moves that appear constantly in campaigns, punditry, and online argument.",
    audience: "Best for civics, debate, media literacy, and rhetoric courses.",
    names: [
      "Ad hominem",
      "Appeal to fear",
      "Appeal to authority",
      "Bare assertion fallacy",
      "Cherry picking",
      "False dilemma",
      "False equivalence",
      "Moving the goalpost",
      "Red herring",
      "Slippery slope",
      "Straw man argument",
      "Tu quoque",
    ],
  },
  {
    slug: "often-confused",
    title: "Most often confused",
    description: "A comparison path for near neighbors that students and readers regularly collapse into one another.",
    audience: "Best for high school and college review sessions.",
    names: [
      "Anecdotal fallacy",
      "Appeal to authority",
      "Correlation is not causation",
      "False analogy",
      "False dilemma",
      "False equivalence",
      "Faulty generalization",
      "Hasty generalization",
      "Poisoning the well",
      "Post hoc ergo propter hoc",
      "Ad hominem",
      "False attribution",
    ],
  },
];
const funPromptDefinitions = [
  {
    slug: "balanced-media-fallacy-hunt",
    title: "Balanced Media Fallacy Hunt",
    intro:
      "This prompt asks an AI model with web access to compare recent left-leaning and right-leaning political arguments, identify distinct fallacies, quote the relevant passages, and link each diagnosis back to LogFall. It is designed for classroom comparison, media-literacy practice, and discussion that stays anchored to actual source text rather than vague impressions.",
    requirements:
      "Nothing else needs to be pasted in if the model can browse the live web. The copied prompt is enough, but it works only with a model that can search for and open recent articles on its own.",
    prompt: `Search the web for six recent political news or opinion articles published within the last 7 days. Select three articles that argue from a clearly left-leaning perspective and three that argue from a clearly right-leaning perspective. Each selected article must contain a different logical fallacy, and the fallacy attribution must be defensible from a quoted passage rather than from the article's general tone.

Return your answer as a Markdown table with exactly two columns: "Left-Wing Arguments" and "Right-Wing Arguments", and exactly three data rows.

Inside each cell, use this exact structure with HTML line breaks for spacing:

◉ **[Fallacy Name]**
<br><br>
➘ **Source:** [Publication - Article Title (Date)](URL)
<br><br>
➘ **Quoted Passage:** "[Quote enough of the original passage to make the reasoning misstep clear.]"
<br><br>
➘ **Explanation:** [Write one concise paragraph explaining exactly how the quoted passage commits the fallacy, why that label fits better than close alternatives, and what reasoning move goes wrong.]
<br><br>
➘ **LogFall Reference:** [Direct link to the relevant page at https://logfall.com/fallacies/]
<br><br>
➘ **Discussion Questions:**<br>
— [Question 1]<br>
— [Question 2]<br>
— [Question 3]

Rules:
- Use six different fallacies total.
- Use only sources published within the last 7 days.
- Replace any article whose partisan orientation or fallacy diagnosis is too unclear.
- Keep the explanations evidence-based and quote-driven.
- Use \`◉ \` before each fallacy title, use \`➘ \` before labels and quoted material, and use \`— \` before the discussion-question items.
- For the discussion questions, do not use the word "moral" or any of its derivatives. If needed, use terms like "pro-social," "compassionate," or "cooperative" instead.`,
  },
  {
    slug: "passage-fallacy-analyzer",
    title: "Passage Fallacy Analyzer",
    intro:
      "This prompt is for pasting in any argumentative passage and asking an AI model to find every defensible fallacy within it. It pushes the model to quote enough of the original wording to make the mistake visible, explain the dynamics of the misstep in more depth, and link each diagnosis back to the relevant LogFall page.",
    requirements:
      "This one needs more than the copied prompt itself. You should paste or attach the exact passage you want analyzed in the placeholder area at the end so the model can quote the wording precisely.",
    prompt: `Analyze the passage below and identify every distinct logical fallacy that can be justified from the text. Do not force a fallacy label where the evidence is weak. If no clear fallacy is present, say so explicitly.

Return the result as a Markdown table with these columns:
1. Fallacy
2. Quoted Passage
3. Dynamics of the Misstep
4. LogFall Reference

For each row:
- In the "Fallacy" cell, write \`◉ [Most specific fallacy name]\`.
- In the "Quoted Passage" cell, begin with \`➘ Quoted Passage:\` and then quote enough of the original text to make the misstep understandable on its own.
- In the "Dynamics of the Misstep" cell, begin with \`➘ Dynamics:\` and then write 3 to 5 sentences explaining exactly how the reasoning goes wrong, what argumentative move is being made, and why this fallacy label fits better than close alternatives.
- In the "LogFall Reference" cell, begin with \`➘ LogFall Reference:\` and then provide a direct link to the most relevant page at https://logfall.com/fallacies/.

Additional rules:
- Include multiple rows if the passage contains multiple fallacies.
- Do not list the same fallacy twice unless it occurs in a clearly different place.
- Be charitable: reconstruct the strongest reasonable version of the argument before judging it.
- Use \`◉ \` for the fallacy label and \`➘ \` for the quoted passage, dynamics, and reference labels.
- No discussion questions are needed.

Passage to analyze:
[PASTE PASSAGE HERE]`,
  },
  {
    slug: "fallacy-repair-workshop",
    title: "Fallacy Repair Workshop",
    intro:
      "This prompt takes a weak argument, diagnoses the fallacy, and then rebuilds the argument into the strongest fair version that preserves as much of the original point as possible. It is useful when the goal is not merely critique, but learning how better reasoning actually sounds.",
    requirements:
      "This prompt needs the original argument to be pasted or attached. It works best when you give the exact wording you want repaired rather than a summary of it.",
    prompt: `Analyze the argument below as a repair workshop rather than as a mere takedown.

Your task has four steps:
1. Identify the most important logical fallacy or reasoning failure in the passage.
2. Quote the exact sentence or sentences where the problem appears.
3. Explain in 3 to 5 sentences exactly how the reasoning goes wrong and why this fallacy label fits better than close alternatives.
4. Rewrite the argument into the strongest non-fallacious version you can while preserving as much of the original concern, intuition, or conclusion as possible.

Return the answer in this structure:

◉ **Fallacy Diagnosis**
➘ **Fallacy:** [Name]
➘ **Quoted Passage:** "[Quote enough of the original wording to make the problem clear.]"
➘ **Dynamics of the Misstep:** [3 to 5 sentences]
➘ **LogFall Reference:** [Direct link to the most relevant page at https://logfall.com/fallacies/]

◉ **Stronger Revision**
[Write a revised version of the argument that removes the fallacy while keeping the core point as strong as the evidence allows.]

◉ **What Changed**
➘ [Point 1]
➘ [Point 2]
➘ [Point 3]

Rules:
- Be charitable before criticizing.
- If the passage contains multiple fallacies, focus on the one that most seriously distorts the reasoning, then briefly note any secondary ones.
- Do not merely weaken the claim; improve the reasoning.
- Use \`◉ \` before major section titles and \`➘ \` before labels and list items.

Argument to repair:
[PASTE ARGUMENT HERE]`,
  },
  {
    slug: "near-neighbor-comparator",
    title: "Near-Neighbor Comparator",
    intro:
      "This prompt is for cases where readers keep confusing similar fallacies. It forces the model to choose among close competitors and explain why the rejected labels do not fit, which is often the real learning bottleneck.",
    requirements:
      "This prompt needs a pasted or attached argument to classify. It works best when the passage is short enough to compare carefully against several close rival fallacies.",
    prompt: `Analyze the argument below and decide which fallacy label fits best from a set of near neighbors.

Use these steps:
1. Identify the single best-fitting fallacy.
2. Compare it against at least three close alternatives that might tempt a careless reader.
3. Quote the original passage enough to make the difference understandable.
4. Explain exactly why the winning label fits better than the others.

Return the answer as a Markdown table with these columns:
1. Candidate Fallacy
2. Fits or Not?
3. Why
4. LogFall Reference

Rules:
- In the "Candidate Fallacy" column, write \`◉ [Fallacy name]\`.
- In the "Why" and "LogFall Reference" columns, begin the content with \`➘ \` labels where helpful.
- The first row must be the best-fitting fallacy.
- Include at least three rejected alternatives.
- In each "Why" cell, explain the decisive difference in reasoning structure, not just the surface wording.
- Quote enough of the source passage before the table to make the classification intelligible.
- Use direct LogFall links for every listed fallacy.
- Use \`◉ \` for the winning and rival fallacy labels and \`➘ \` for supporting labels or list-like elements.

After the table, add a short section titled:
◉ **Final Verdict**
[Write one concise paragraph explaining the single best label and the exact reason it wins.]

Argument to classify:
[PASTE ARGUMENT HERE]

Suggested comparison set when relevant:
Ad hominem, Poisoning the well, Straw man argument, Red herring, False dilemma, False equivalence, Hasty generalization, Cherry picking, Correlation is not causation, Post hoc ergo propter hoc.`,
  },
  {
    slug: "steelman-then-diagnose",
    title: "Steelman Then Diagnose",
    intro:
      "This prompt trains the habit of fairness before critique. It first asks the model to reconstruct the strongest reasonable version of an argument, and only then identify any remaining fallacies or weaknesses.",
    requirements:
      "This prompt needs a pasted or attached argument or excerpt. Give enough surrounding context for the model to reconstruct the strongest fair version before it starts diagnosing flaws.",
    prompt: `Read the argument below and follow this sequence strictly:

1. **Steelman first:** Rewrite the argument in its strongest fair form, preserving the speaker's apparent goal while removing ambiguity, filler, and accidental weakness.
2. **Diagnose second:** Analyze the steelmanned version and identify any logical fallacies or remaining reasoning failures that still survive.
3. **Quote the source:** Quote the original passage enough to show where the problematic move first appeared.
4. **Explain the difference:** Show how the steelmanned version improves the argument and what defects, if any, remain even after charitable reconstruction.

Return the answer in this structure:

◉ **Original Passage**
"[Quoted passage]"

◉ **Steelmanned Version**
[Best fair reconstruction]

◉ **Remaining Fallacies or Weaknesses**
Use a Markdown table with these columns:
1. Fallacy or Weakness
2. Present After Steelmanning?
3. Explanation
4. LogFall Reference

Rules:
- In the "Fallacy or Weakness" column, begin each entry with \`◉ \`.
- In the "Explanation" and "LogFall Reference" columns, use \`➘ \` labels where helpful to improve readability.
- If steelmanning removes the apparent fallacy entirely, say so clearly.
- Do not confuse bad wording with bad reasoning.
- In the explanation column, use 3 to 5 sentences whenever a genuine fallacy remains.
- Link to LogFall only when the issue is truly a fallacy rather than a non-fallacious weakness.
- Use \`◉ \` for major section titles and fallacy/weakness labels, and \`➘ \` for sub-points or supporting labels.

Argument to analyze:
[PASTE ARGUMENT HERE]`,
  },
  {
    slug: "argument-map-builder",
    title: "Argument Map Builder",
    intro:
      "This prompt is meant to slow an argument down into premises, hidden assumptions, and conclusion, then identify the exact step where the reasoning stops earning its result. It works especially well for formal, causal, and evidential mistakes.",
    requirements:
      "This prompt needs a pasted or attached passage to map. It works best when the passage contains a fairly clear conclusion and at least one inferential step that can be broken into premises and assumptions.",
    prompt: `Turn the passage below into an explicit argument map, then identify where the reasoning fails.

Return the answer in this structure:

◉ **Argument Map**
➘ **Conclusion:** [Main conclusion]
➘ **Stated Premises:** [List them clearly]
➘ **Hidden Assumptions:** [List any unstated assumptions needed for the argument to work]
➘ **Inference Path:** [Show how the argument moves from premises to conclusion]

◉ **Failure Point**
Use a Markdown table with these columns:
1. Step in the Argument
2. What the Step Claims
3. What Goes Wrong
4. Fallacy Name
5. LogFall Reference

◉ **Short Diagnosis**
[Write one concise paragraph explaining the main reasoning failure in plain language.]

Rules:
- If the argument contains more than one failure point, include more than one row.
- Quote short key phrases from the original passage where needed.
- Be explicit about whether the problem is evidential, causal, formal, conceptual, or rhetorical.
- Use the most specific fallacy label you can justify.
- In the "Fallacy Name" column, begin each entry with \`◉ \`, and use \`➘ \` for labels or list-like sub-points elsewhere.

Passage to map:
[PASTE PASSAGE HERE]`,
  },
];
const foundationalNames = new Set([
  ...featuredNames,
  "Appeal to emotion",
  "Appeal to fear",
  "Hasty generalization",
  "Red herring",
  "Tu quoque",
  "Wishful thinking",
]);
const keywordStopwords = new Set([
  "about",
  "after",
  "again",
  "against",
  "already",
  "also",
  "among",
  "argument",
  "because",
  "before",
  "being",
  "between",
  "claim",
  "conclusion",
  "context",
  "could",
  "does",
  "doing",
  "even",
  "from",
  "have",
  "into",
  "just",
  "many",
  "more",
  "must",
  "need",
  "only",
  "other",
  "over",
  "really",
  "reasoning",
  "same",
  "should",
  "some",
  "still",
  "such",
  "than",
  "that",
  "their",
  "there",
  "these",
  "they",
  "this",
  "those",
  "through",
  "treats",
  "under",
  "using",
  "when",
  "where",
  "which",
  "while",
  "with",
  "without",
]);
const categoryQuizKeywords = {
  Formal: ["premises", "conclusion", "follow", "valid"],
  Mathematical: ["rates", "sample", "numbers", "probability"],
  Causal: ["cause", "correlation", "confounder", "mechanism"],
  Linguistic: ["meaning", "term", "ambiguity", "definition"],
  Conceptual: ["category", "alternatives", "boundary", "classification"],
  Evidential: ["evidence", "missing", "support", "overstated"],
  Perceptual: ["appearance", "seems", "vivid", "shown"],
  Perspectival: ["frame", "viewpoint", "context", "timeline"],
  Epistemic: ["confidence", "know", "evidence", "belief"],
  Tactical: ["issue", "diversion", "claim", "response"],
  Emotional: ["emotion", "pressure", "evidence", "persuasion"],
};
const gaugeCategoryProfiles = {
  Formal: { common: 24, spot: 26, innocent: 50 },
  Mathematical: { common: 46, spot: 34, innocent: 76 },
  Causal: { common: 68, spot: 52, innocent: 82 },
  Linguistic: { common: 64, spot: 40, innocent: 72 },
  Conceptual: { common: 52, spot: 36, innocent: 70 },
  Evidential: { common: 78, spot: 50, innocent: 84 },
  Perceptual: { common: 60, spot: 56, innocent: 78 },
  Perspectival: { common: 42, spot: 32, innocent: 62 },
  Epistemic: { common: 58, spot: 34, innocent: 68 },
  Tactical: { common: 82, spot: 70, innocent: 48 },
  Emotional: { common: 80, spot: 66, innocent: 66 },
};
const rhetoricGaugeOverrides = {
  "Absence of evidence fallacy": { common: 68, spot: 42, innocent: 76 },
  "Ad hominem": { common: 90, spot: 88, innocent: 72 },
  "Anecdotal fallacy": { common: 84, spot: 74, innocent: 86 },
  "Appeal to authority": { common: 80, spot: 72, innocent: 78 },
  "Appeal to emotion": { common: 86, spot: 72, innocent: 78 },
  "Appeal to fear": { common: 84, spot: 78, innocent: 68 },
  "Appeal to nature": { common: 70, spot: 60, innocent: 72 },
  "Argument from ignorance": { common: 72, spot: 56, innocent: 74 },
  "Argument from incredulity": { common: 72, spot: 58, innocent: 76 },
  "Base rate fallacy": { common: 62, spot: 40, innocent: 88 },
  "Begging the question": { common: 78, spot: 28, innocent: 78 },
  "Cherry picking": { common: 86, spot: 66, innocent: 82 },
  "Composition fallacy": { common: 42, spot: 30, innocent: 66 },
  "Correlation is not causation": { common: 88, spot: 58, innocent: 88 },
  "Equivocation": { common: 72, spot: 34, innocent: 72 },
  "False analogy": { common: 70, spot: 56, innocent: 78 },
  "False balance": { common: 76, spot: 52, innocent: 66 },
  "False dilemma": { common: 86, spot: 72, innocent: 84 },
  "False equivalence": { common: 82, spot: 60, innocent: 72 },
  "Faulty generalization": { common: 70, spot: 48, innocent: 82 },
  "Hasty generalization": { common: 84, spot: 64, innocent: 86 },
  "Moving the goalpost": { common: 80, spot: 74, innocent: 56 },
  "No True Scotsman": { common: 74, spot: 54, innocent: 68 },
  "Poisoning the well": { common: 72, spot: 60, innocent: 50 },
  "Post hoc ergo propter hoc": { common: 68, spot: 48, innocent: 84 },
  "Red herring": { common: 82, spot: 56, innocent: 58 },
  "Slippery slope": { common: 82, spot: 70, innocent: 74 },
  "Straw man argument": { common: 90, spot: 76, innocent: 74 },
  "Survivorship bias": { common: 60, spot: 42, innocent: 80 },
  "Tu quoque": { common: 80, spot: 76, innocent: 64 },
  "Wishful thinking": { common: 74, spot: 54, innocent: 82 },
};

function escapeHtml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function slugify(value = "") {
  return value
    .toLowerCase()
    .replaceAll("&", " and ")
    .replaceAll("/", " ")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^-|-$/g, "");
}

function truncate(value, max = 180) {
  if (!value || value.length <= max) return value;
  return `${value.slice(0, max).trimEnd()}...`;
}

function normalizeCaseStudy(item) {
  if (!item) {
    return { summary: "", source: "", title: "", date: "", url: "" };
  }

  if (typeof item === "string") {
    return { summary: item, source: "", title: "", date: "", url: "" };
  }

  return {
    summary: item.summary || "",
    source: item.source || "",
    title: item.title || "",
    date: item.date || "",
    url: item.url || "",
  };
}

function formatCaseStudyCell(item) {
  const study = normalizeCaseStudy(item);
  const parts = [];
  if (study.title) parts.push(study.title);
  if (study.summary) parts.push(study.summary);
  const details = [study.source, study.date].filter(Boolean).join(" | ");
  if (details) parts.push(`Source: ${details}`);
  if (study.url) parts.push(study.url);
  return parts.join("\n");
}

function renderCaseStudy(item) {
  const study = normalizeCaseStudy(item);
  const titleLine = study.title
    ? study.url
      ? `<p class="case-title"><a href="${escapeHtml(study.url)}" title="${escapeHtml(study.title)}">${escapeHtml(study.title)}</a></p>`
      : `<p class="case-title">${escapeHtml(study.title)}</p>`
    : "";
  const sourceLabel = [study.source, study.date].filter(Boolean).join(" · ");
  const sourceBits = [];
  if (sourceLabel) sourceBits.push(escapeHtml(sourceLabel));
  if (study.url && !study.title) {
    sourceBits.push(
      `<a href="${escapeHtml(study.url)}" title="${escapeHtml(study.source || study.url)}">Open source</a>`,
    );
  }
  const sourceLine = sourceBits.join(" · ");

  return `<blockquote class="case-item">
    ${titleLine}
    <p class="case-summary">${escapeHtml(study.summary)}</p>
    ${sourceLine ? `<p class="case-source">${sourceLine}</p>` : ""}
  </blockquote>`;
}

function renderTabGroup(tabKey, items) {
  const buttons = items
    .map(
      (item, index) => `<button
        class="lab-tab${index === 0 ? " active" : ""}"
        type="button"
        role="tab"
        id="${escapeHtml(item.id)}-tab"
        aria-controls="${escapeHtml(item.id)}"
        aria-selected="${index === 0 ? "true" : "false"}"
        data-tab-button
      >${escapeHtml(item.label)}</button>`,
    )
    .join("");

  const panels = items
    .map(
      (item, index) => `<section
        class="lab-panel${index === 0 ? " active" : ""}"
        role="tabpanel"
        id="${escapeHtml(item.id)}"
        aria-labelledby="${escapeHtml(item.id)}-tab"
        ${index === 0 ? "" : "hidden"}
        data-tab-panel
      >
        <p class="lab-panel-kicker">${escapeHtml(item.label)}</p>
        ${item.content}
      </section>`,
    )
    .join("");

  return `<div class="lab-tab-shell" data-tab-group="${escapeHtml(tabKey)}">
    <div class="lab-tablist" role="tablist" aria-label="Rationality lab">${buttons}</div>
    ${panels}
  </div>`;
}

function ensureSentence(value = "") {
  const trimmed = String(value).trim();
  if (!trimmed) return "";
  return /[.?!]["'”’]?$/u.test(trimmed) ? trimmed : `${trimmed}.`;
}

function splitSentences(value = "") {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .split(/(?<=[.?!]["'”’]?)\s+/u)
    .map((item) => item.trim())
    .filter(Boolean);
}

function firstSentence(value = "") {
  return splitSentences(value)[0] || "";
}

function lowerFirst(value = "") {
  if (!value) return "";
  return value.charAt(0).toLowerCase() + value.slice(1);
}

function definitionCore(definition = "") {
  return stripTrailingPunctuation(
    String(definition || "")
      .replace(/^Occurs when\s+/i, "")
      .replace(/^This fallacy occurs when\s+/i, ""),
  );
}

const caveatMisuseOverrides = {
  "Absence of evidence fallacy":
    "Do not use this label whenever no evidence has been found. In many contexts, a long and competent search that turns up nothing is itself evidence against the claim.",
  "Ad hominem":
    "Do not cry ad hominem every time motives, bias, credibility, or conflicts of interest are discussed. Those personal facts can be relevant when trustworthiness or expertise is part of the evidence.",
  "Appeal to authority":
    "Do not use this label just because someone cites an expert. Relevant expertise, strong track records, and broad expert agreement can be genuine evidence.",
  "Appeal to emotion":
    "Do not use this label whenever an argument contains emotion. Emotion becomes fallacious only when feeling is doing work that reasons or evidence should be doing.",
  "Begging the question":
    "Do not use this label as a synonym for raising a question or sounding circular. It applies only when the conclusion is already being assumed inside the support for the conclusion.",
  "Cherry picking":
    "Do not use this label simply because an argument uses a limited amount of evidence. It becomes cherry picking when the missing evidence is relevant and would materially change the conclusion.",
  "Correlation is not causation":
    "Do not use this label just because someone mentions a correlation. Correlations can be valuable clues and can support causal reasoning when mechanism, timing, controls, and alternatives are handled well.",
  "False dilemma":
    "Do not use this label merely because an argument presents two main options. It becomes fallacious when live alternatives are hidden, dismissed, or never allowed onto the table.",
  "False equivalence":
    "Do not use this label whenever two things are compared. Comparison is legitimate when the similarities really do bear on the point at issue.",
  "Hasty generalization":
    "Do not use this label every time someone draws a general conclusion from limited evidence. Sometimes the sample really is enough for a modest claim; the problem is overreaching beyond what the sample can support.",
  "No True Scotsman":
    "Do not use this label every time someone insists on real criteria for a category. Categories can have genuine standards; the fallacy appears when the standards are improvised only to block a counterexample.",
  "Red herring":
    "Do not use this label every time a discussion broadens or adds background. It is a red herring only when the new point diverts attention from the issue that was supposed to be answered.",
  "Slippery slope":
    "Do not dismiss every warning about escalation as a slippery slope. Some chains really are plausible when incentives, precedent, feedback loops, or institutional weaknesses connect the steps.",
  "Straw man argument":
    "Do not use this label simply because a reply is sharp, selective, or uncharitable. It becomes a straw man only when the reply no longer targets the actual position being discussed.",
  "Tu quoque":
    "Do not use this label whenever hypocrisy is mentioned. Hypocrisy can matter when the issue is sincerity, consistency, or credibility; it becomes tu quoque when hypocrisy is treated as a refutation of the original claim.",
};

const caveatMisuseByCategory = {
  Formal:
    "Do not use this label just because an argument feels abstract, technical, or unpersuasive. The label applies only when the logical form itself is defective.",
  Mathematical:
    "Do not use this label every time numbers, odds, or percentages appear in an argument. The problem has to be a specific misuse of rates, samples, frequencies, or statistical comparison.",
  Causal:
    "Do not use this label every time someone proposes a causal story. The label applies only when the causal leap outruns the evidence, mechanism, timing, or controls.",
  Linguistic:
    "Do not use this label just because wording could have been clearer. It applies when ambiguity, redefinition, or verbal drift is doing real argumentative work.",
  Conceptual:
    "Do not use this label every time people disagree about definitions or categories. It applies when the category boundaries themselves are distorting the reasoning.",
  Evidential:
    "Do not use this label simply because the evidence is incomplete. It applies when the argument claims more support than the evidence has actually earned.",
  Perceptual:
    "Do not use this label just because a case feels vivid, memorable, or striking. It applies when appearances or salience are being treated as if they were stronger evidence than they are.",
  Perspectival:
    "Do not use this label every time someone takes a strong point of view. It applies when a missing frame, timescale, comparison class, or standpoint distorts the conclusion.",
  Epistemic:
    "Do not use this label every time someone sounds too confident, too skeptical, or too simplified. It applies when belief or doubt is being managed badly relative to what can responsibly be known.",
  Tactical:
    "Do not use this label every time an argument feels unfair, heated, or evasive. It applies when the move really does distract from, pressure, or replace the reasoning at issue.",
  Emotional:
    "Do not use this label whenever an argument carries emotional force. It applies when emotion is being asked to do evidential or logical work it has not earned.",
};

function caveatQualifierSentence(record) {
  const candidates = splitSentences(record.notes || "");
  const found = candidates.find((sentence) =>
    /(sometimes|can be|can matter|may be|not enough|really is|are real|good evidence|genuine criteria|genuine standards|not merely|not enough)/i.test(
      sentence,
    ),
  );
  return found ? ensureSentence(found) : "";
}

function caveatMisuseText(record) {
  const category = record.categories[0] || "";
  const base =
    caveatMisuseOverrides[record.name] ||
    caveatMisuseByCategory[category] ||
    "Do not use this label as a catch-all for any weak or unconvincing argument. Reserve it for the exact reasoning slip it names.";
  const qualifier = caveatQualifierSentence(record);
  if (!qualifier) return ensureSentence(base);
  if (base.includes(qualifier) || qualifier.includes(base)) return ensureSentence(base);
  return `${ensureSentence(base)} ${qualifier}`;
}

function caveatApplyText(record, records) {
  const confusion = confusionCandidates(record, records, 1)[0];
  const base = `Use this label only when ${lowerFirst(definitionCore(record.definition))}.`;
  if (!confusion) return ensureSentence(base);
  return `${ensureSentence(base)} If the real problem is that ${lowerFirst(definitionCore(confusion.candidate.definition))}, the better label is ${confusion.candidate.name}.`;
}

function caveatTextForWorkbook(record, records) {
  return `${caveatMisuseText(record)}\n${caveatApplyText(record, records)}`;
}

function extractKeywords(text = "", limit = 4) {
  const matches = String(text || "").toLowerCase().match(/[a-z]{5,}/g) || [];
  const unique = [];
  for (const word of matches) {
    if (keywordStopwords.has(word) || unique.includes(word)) continue;
    unique.push(word);
    if (unique.length >= limit) break;
  }
  return unique;
}

function domainTagForRecord(record) {
  const categories = new Set(record.categories);
  if (categories.has("Formal") || categories.has("Mathematical")) return "Formal logic";
  if (categories.has("Causal") || categories.has("Evidential")) return "Scientific reasoning";
  if (categories.has("Tactical") || categories.has("Emotional") || categories.has("Linguistic")) {
    return "Rhetoric / debate";
  }
  return "Critical thinking / philosophy";
}

function difficultyForRecord(record) {
  if (foundationalNames.has(record.name)) return "Foundational";

  let score = 0;
  const categories = new Set(record.categories);
  if (categories.has("Formal") || categories.has("Mathematical")) score += 3;
  if (categories.has("Epistemic") || categories.has("Linguistic") || categories.has("Perspectival")) score += 2;
  if (categories.has("Causal") || categories.has("Conceptual")) score += 1;
  if (record.subCategory || record.subSubCategory) score += 1;
  if ((record.notes || "").length > 260) score += 1;

  if (score <= 1) return "Foundational";
  if (score <= 3) return "Intermediate";
  return "Advanced";
}

function classroomLevelForRecord(record, difficulty = "") {
  if (difficulty === "Foundational") return "Middle school+";
  if (difficulty === "Intermediate") return "High school";
  if (record.categories.includes("Formal") || record.categories.includes("Mathematical") || record.categories.includes("Epistemic")) {
    return "Advanced undergraduate";
  }
  return "Intro college";
}

const pedagogyCache = new Map();
const rhetoricGaugeCache = new Map();

function pedagogyForRecord(record) {
  if (pedagogyCache.has(record.slug)) return pedagogyCache.get(record.slug);

  const difficulty = difficultyForRecord(record);
  const classroomLevel = classroomLevelForRecord(record, difficulty);
  const domainTag = domainTagForRecord(record);
  const teachingPaths = teachingPathDefinitions
    .filter((path) => path.names.includes(record.name))
    .map((path) => ({ slug: path.slug, title: path.title }));

  const meta = {
    difficulty,
    classroomLevel,
    domainTag,
    classroomTags: [classroomLevel, domainTag],
    teachingPaths,
  };
  pedagogyCache.set(record.slug, meta);
  return meta;
}

function clampNumber(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function weightedGaugeBase(record, metric) {
  const weights = [0.56, 0.29, 0.15];
  let total = 0;
  let weightTotal = 0;

  record.categories.forEach((category, index) => {
    const categoryProfile = gaugeCategoryProfiles[category];
    if (!categoryProfile) return;
    const weight = weights[index] || 0.1;
    total += categoryProfile[metric] * weight;
    weightTotal += weight;
  });

  return weightTotal ? total / weightTotal : 50;
}

function commonGaugeNarrative(score) {
  if (score >= 85) return "Shows up constantly in current politics, media, and online argument.";
  if (score >= 70) return "Appears regularly in everyday public rhetoric.";
  if (score >= 55) return "Common enough that most readers will meet it often.";
  if (score >= 40) return "Present, but more situation-dependent than the headline fallacies.";
  return "Relatively uncommon in ordinary rhetoric compared with the better-known fallacies.";
}

function spotGaugeNarrative(score) {
  if (score >= 85) return "Usually visible almost immediately once readers know the pattern.";
  if (score >= 70) return "Often easy to catch with a little attention.";
  if (score >= 55) return "Recognizable, but easy to miss in a fast or heated exchange.";
  if (score >= 40) return "Often hides inside wording, framing, or technical detail.";
  return "Hard to see without slowing down and reconstructing the reasoning.";
}

function innocentGaugeNarrative(score) {
  if (score >= 85) return "Very easy for well-meaning people to commit without noticing.";
  if (score >= 70) return "A frequent unintentional slip in ordinary reasoning.";
  if (score >= 55) return "Sometimes accidental and sometimes more strategic.";
  if (score >= 40) return "Less often innocent; the move usually takes more pressure or steering.";
  return "Usually feels more deliberate than accidental.";
}

function gaugeBandLabel(metric, score) {
  if (metric === "common") {
    if (score >= 85) return "Near-constant";
    if (score >= 70) return "Very common";
    if (score >= 55) return "Recurring";
    if (score >= 40) return "Occasional";
    return "Uncommon";
  }
  if (metric === "spot") {
    if (score >= 85) return "Obvious";
    if (score >= 70) return "Easy to catch";
    if (score >= 55) return "Moderate";
    if (score >= 40) return "Tricky";
    return "Hard to spot";
  }
  if (score >= 85) return "Almost automatic";
  if (score >= 70) return "Very easy to slip into";
  if (score >= 55) return "Common slip";
  if (score >= 40) return "Moderate risk";
  return "Low accidental risk";
}

function rhetoricGaugesForRecord(record) {
  if (rhetoricGaugeCache.has(record.slug)) return rhetoricGaugeCache.get(record.slug);

  const pedagogy = pedagogyForRecord(record);
  const teachingPathSlugs = new Set(pedagogy.teachingPaths.map((item) => item.slug));
  const difficulty = pedagogy.difficulty;
  const foundational = foundationalNames.has(record.name);
  const aliasesBoost = record.aliases.length ? Math.min(4, record.aliases.length * 2) : 0;

  let common =
    weightedGaugeBase(record, "common") +
    (featuredNames.includes(record.name) ? 10 : 0) +
    (teachingPathSlugs.has("public-debate") ? 12 : 0) +
    (teachingPathSlugs.has("start-here") ? 6 : 0) +
    (teachingPathSlugs.has("often-confused") ? 3 : 0) +
    aliasesBoost;

  let spot =
    weightedGaugeBase(record, "spot") +
    (difficulty === "Foundational" ? 8 : difficulty === "Intermediate" ? 2 : -8) +
    (record.categories.includes("Linguistic") || record.categories.includes("Epistemic") ? -4 : 0) +
    (record.categories.includes("Tactical") || record.categories.includes("Emotional") ? 4 : 0) +
    (foundational ? 4 : 0);

  let innocent =
    weightedGaugeBase(record, "innocent") +
    (record.categories.includes("Tactical") ? -8 : 0) +
    (record.categories.includes("Formal") || record.categories.includes("Mathematical") ? 4 : 0) +
    (record.categories.includes("Emotional") || record.categories.includes("Evidential") ? 4 : 0) +
    (foundational ? 3 : 0);

  const override = rhetoricGaugeOverrides[record.name];
  if (override) {
    common = override.common;
    spot = override.spot;
    innocent = override.innocent;
  }

  const gauges = {
    common: {
      title: "Common in today's rhetoric",
      value: clampNumber(Math.round(common / 5) * 5, 15, 95),
      lowLabel: "Rare",
      highLabel: "Constant",
    },
    spot: {
      title: "Easy to spot",
      value: clampNumber(Math.round(spot / 5) * 5, 15, 95),
      lowLabel: "Hidden",
      highLabel: "Obvious",
    },
    innocent: {
      title: "Easy to innocently commit",
      value: clampNumber(Math.round(innocent / 5) * 5, 15, 95),
      lowLabel: "Low risk",
      highLabel: "Easy slip",
    },
  };

  gauges.common.band = gaugeBandLabel("common", gauges.common.value);
  gauges.common.summary = commonGaugeNarrative(gauges.common.value);
  gauges.spot.band = gaugeBandLabel("spot", gauges.spot.value);
  gauges.spot.summary = spotGaugeNarrative(gauges.spot.value);
  gauges.innocent.band = gaugeBandLabel("innocent", gauges.innocent.value);
  gauges.innocent.summary = innocentGaugeNarrative(gauges.innocent.value);

  rhetoricGaugeCache.set(record.slug, gauges);
  return gauges;
}

const confusionCache = new Map();

function confusionCandidates(record, records, limit = 2) {
  if (confusionCache.has(record.slug)) {
    return confusionCache.get(record.slug).slice(0, limit);
  }

  const ranked = records
    .filter((candidate) => candidate.slug !== record.slug)
    .map((candidate) => {
      const sharedCategories = candidate.categories.filter((category) => record.categories.includes(category));
      const score =
        sharedCategories.length * 14 +
        (candidate.categories[0] === record.categories[0] ? 6 : 0) +
        (candidate.family && candidate.family === record.family ? 3 : 0) +
        (candidate.subCategory && candidate.subCategory === record.subCategory ? 2 : 0) +
        (candidate.subSubCategory && candidate.subSubCategory === record.subSubCategory ? 1 : 0);
      return { candidate, sharedCategories, score };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.candidate.name.localeCompare(b.candidate.name));

  confusionCache.set(record.slug, ranked);
  return ranked.slice(0, limit);
}

function overlapTextForConfusion(item) {
  if (!item.sharedCategories.length) {
    return "These can feel similar because they lead readers toward the wrong conclusion through a similar surface move.";
  }
  if (item.sharedCategories.length === 1) {
    return `Both often look like ${item.sharedCategories[0].toLowerCase()} mistakes at first glance.`;
  }
  return `Both often look like ${item.sharedCategories.map((value) => value.toLowerCase()).join(" and ")} mistakes at first glance.`;
}

function modelExplanationForRecord(record, confusionItem) {
  const base = `This is ${record.name} because ${lowerFirst(definitionCore(record.definition))}.`;
  if (!confusionItem) return ensureSentence(base);
  const candidate = confusionItem.candidate;
  return ensureSentence(
    `${base} The key difference from ${candidate.name} is that ${lowerFirst(definitionCore(candidate.definition))}.`,
  );
}

function quizKeywordsForRecord(record) {
  const category = record.categories[0] || "";
  return [...new Set([...(categoryQuizKeywords[category] || []), ...extractKeywords(definitionCore(record.definition)), ...extractKeywords(record.notes || "")])].slice(0, 6);
}

function quizConfigForRecord(record, records) {
  const confusionItems = confusionCandidates(record, records, 3);
  const used = new Set([record.name]);
  const options = [record.name];

  for (const item of confusionItems) {
    if (used.has(item.candidate.name)) continue;
    used.add(item.candidate.name);
    options.push(item.candidate.name);
  }

  if (options.length < 4) {
    for (const candidate of records) {
      if (candidate.slug === record.slug || used.has(candidate.name)) continue;
      if (candidate.categories[0] !== record.categories[0]) continue;
      used.add(candidate.name);
      options.push(candidate.name);
      if (options.length >= 4) break;
    }
  }

  if (options.length < 4) {
    for (const name of featuredNames) {
      if (used.has(name) || name === record.name) continue;
      used.add(name);
      options.push(name);
      if (options.length >= 4) break;
    }
  }

  return {
    options: options.sort((a, b) => a.localeCompare(b)),
    answer: record.name,
    keywords: quizKeywordsForRecord(record),
    model: modelExplanationForRecord(record, confusionItems[0]),
  };
}

function repairModelForRecord(record) {
  const category = record.categories[0] || "";
  const prompt = diagnosticPrompts[category] || "What extra support would the conclusion need before it is justified?";
  const templates = {
    Formal: `These premises do not yet prove the conclusion. A stronger version would either add the missing premise or weaken the conclusion to what really follows.`,
    Mathematical: `A stronger version would state the relevant rate, sample, or comparison explicitly before drawing the conclusion.`,
    Causal: `This pattern may justify further investigation, but it does not by itself show causation; timing, mechanism, controls, and rival explanations still matter.`,
    Linguistic: `A stronger version would keep the key term in one stable sense and restate the claim only after the wording is clear.`,
    Conceptual: `A stronger version would define the category more carefully and say what happens once the omitted alternatives are put back on the table.`,
    Evidential: `A stronger version would say only what the evidence currently supports and name what evidence is still missing.`,
    Perceptual: `A stronger version would separate what looks vivid or striking from what has actually been shown.`,
    Perspectival: `A stronger version would widen the frame and restate the conclusion so it still holds once the missing context is included.`,
    Epistemic: `A stronger version would lower the confidence of the claim to match what can responsibly be known from the evidence.`,
    Tactical: `A stronger version would answer the original claim directly instead of shifting the conversation away from it.`,
    Emotional: `A stronger version would remove the emotional pressure and defend the conclusion with reasons or evidence.`,
  };
  return ensureSentence(templates[category] || `A stronger version would say only what the reasoning has actually earned. ${prompt}`);
}

function repairChecklistForRecord(record) {
  const category = record.categories[0] || "";
  const shared = [
    "State only what the argument or evidence really supports.",
    "Remove the exact move that made the original version fallacious.",
  ];
  const byCategory = {
    Formal: ["Check whether the conclusion really follows from the stated premises."],
    Mathematical: ["Make the relevant comparison, rate, or sample size explicit."],
    Causal: ["Name at least one rival explanation, missing mechanism, or control."],
    Linguistic: ["Keep the key term in one stable sense throughout the rewrite."],
    Conceptual: ["Define the category carefully and restore the omitted alternatives."],
    Evidential: ["Say what evidence would need to be added before the stronger claim could be justified."],
    Perceptual: ["Separate how the case feels or looks from what has been demonstrated."],
    Perspectival: ["Add the missing frame, timeline, or viewpoint before concluding."],
    Epistemic: ["Calibrate the confidence of the rewrite to the actual evidence."],
    Tactical: ["Answer the original issue rather than the distraction or attack."],
    Emotional: ["Make sure the rewrite would still persuade without the emotional push."],
  };
  return [...shared, ...(byCategory[category] || [])];
}

function argumentMapForRecord(record) {
  const category = record.categories[0] || "";
  const problem = ensureSentence(record.mainReasoningProblem || definitionCore(record.definition));
  const prompt = diagnosticPrompts[category] || "What is missing before the conclusion is earned?";

  if (category === "Formal") {
    return {
      variant: "formal",
      title: "Argument map",
      intro: "This map highlights the gap between the stated structure and the conclusion the argument tries to force.",
      nodes: [
        { label: "Premise pattern", text: ensureSentence(record.example || record.definition) },
        { label: "Invalid step", text: `The structure fails when ${lowerFirst(problem)}` },
        { label: "What the premises still allow", text: ensureSentence(record.notes || "The observed consequence could still have other explanations or supporting structures.") },
        { label: "What a valid repair needs", text: ensureSentence(prompt) },
      ],
    };
  }

  if (category === "Causal") {
    return {
      variant: "causal",
      title: "Argument map",
      intro: "This map shows where an observed pattern gets promoted into a stronger causal story than the evidence can support.",
      nodes: [
        { label: "Observed pattern", text: ensureSentence(record.example || record.definition) },
        { label: "Claimed cause", text: `The leap happens when ${lowerFirst(problem)}` },
        { label: "Missing checks", text: ensureSentence(record.notes || "Alternative causes, timing, mechanisms, and controls still need to be checked.") },
        { label: "Safer conclusion", text: ensureSentence(prompt) },
      ],
    };
  }

  return null;
}

function toolModeForCategory(category = "") {
  if (["Causal"].includes(category)) return "causal";
  if (["Evidential", "Epistemic"].includes(category)) return "evidential";
  if (["Tactical", "Emotional"].includes(category)) return "tactical";
  if (["Conceptual", "Linguistic", "Perceptual", "Perspectival"].includes(category)) return "conceptual";
  if (["Formal", "Mathematical"].includes(category)) return "structural";
  return "default";
}

function buildAuditPreset(record, prompt, primaryProfile) {
  const category = record.categories[0] || "";
  const mode = toolModeForCategory(category);
  const baseExample = ensureSentence(record.example || record.definition);
  const problem = ensureSentence(record.mainReasoningProblem || record.definition);
  const notes = ensureSentence(record.notes || record.feedbackLogic || record.rationalityDanger || "");
  const repairPrompt = ensureSentence(prompt || record.repairPrompts || "Ask what extra step the conclusion needs before it is earned.");

  const presets = {
    causal: {
      title: "Causal jump walkthrough",
      intro: "Step through where the example moves from a pattern or event to a stronger causal claim than the evidence can bear.",
      steps: [
        {
          label: "Pattern",
          title: "What is actually observed",
          body: `The example starts with this observed pattern or event: ${baseExample}`,
        },
        {
          label: "Jump",
          title: "Where the causal leap enters",
          body: `The reasoning goes wrong when ${lowerFirst(problem)}`,
        },
        {
          label: "Missing test",
          title: "What still has to be checked",
          body: notes || "A stronger causal claim still has to rule out coincidence, reverse causation, or a third factor.",
        },
        {
          label: "Repair",
          title: "How to challenge the move",
          body: repairPrompt,
        },
      ],
    },
    evidential: {
      title: "Evidence gap walkthrough",
      intro: "Use the example to see exactly where the conclusion outruns what the evidence, sampling, or confidence level can support.",
      steps: [
        {
          label: "Shown",
          title: "What the example actually gives you",
          body: `The page begins with this evidence claim or data point: ${baseExample}`,
        },
        {
          label: "Overreach",
          title: "Where the conclusion goes too far",
          body: `The overreach happens when ${lowerFirst(problem)}`,
        },
        {
          label: "Gap",
          title: "What support is still missing",
          body: notes || "The argument still owes you stronger support, a fairer sample, or a better calibration of confidence.",
        },
        {
          label: "Repair",
          title: "What to ask next",
          body: repairPrompt,
        },
      ],
    },
    tactical: {
      title: "Debate-move walkthrough",
      intro: "Follow the example to see how attention gets redirected away from the real issue and toward pressure, distraction, or emotional force.",
      steps: [
        {
          label: "Surface",
          title: "What the exchange looks like on the surface",
          body: `At first glance, the move sounds like this: ${baseExample}`,
        },
        {
          label: "Pressure",
          title: "Where the persuasive shove happens",
          body: `The maneuver enters when ${lowerFirst(problem)}`,
        },
        {
          label: "Displaced issue",
          title: "What gets pushed out of focus",
          body: notes || "The real issue is displaced by pressure, identity cues, or a strategically chosen detour.",
        },
        {
          label: "Reset",
          title: "How to bring the discussion back",
          body: repairPrompt,
        },
      ],
    },
    conceptual: {
      title: "Category and wording walkthrough",
      intro: "Step through the example to see which distinction, definition, or frame is being blurred and why that matters to the conclusion.",
      steps: [
        {
          label: "Frame",
          title: "How the claim is initially framed",
          body: `The reasoning is first presented like this: ${baseExample}`,
        },
        {
          label: "Blur",
          title: "Where the categories or wording slip",
          body: `The trouble begins when ${lowerFirst(problem)}`,
        },
        {
          label: "Confusion",
          title: "What becomes unclear or unstable",
          body: notes || "An important boundary, meaning, or point of view is getting blurred before the conclusion is drawn.",
        },
        {
          label: "Repair",
          title: "How to sharpen the distinction",
          body: repairPrompt,
        },
      ],
    },
    structural: {
      title: "Structure check walkthrough",
      intro: "Use the example to track the setup, the rule being assumed, and the exact place where the structure or numbers stop supporting the conclusion.",
      steps: [
        {
          label: "Setup",
          title: "How the argument is set up",
          body: `The setup looks like this: ${baseExample}`,
        },
        {
          label: "Rule",
          title: "What rule the reasoning is relying on",
          body: `The argument depends on the idea that ${lowerFirst(problem)}`,
        },
        {
          label: "Break",
          title: "Where the structure fails",
          body: notes || "The structure, probability move, or quantitative assumption is not enough to justify the conclusion.",
        },
        {
          label: "Repair",
          title: "How to test the structure",
          body: repairPrompt,
        },
      ],
    },
    default: {
      title: "Reasoning path walkthrough",
      intro: "Walk through the example step by step and identify the exact point where the conclusion starts asking for more than the reasoning has earned.",
      steps: [
        {
          label: "Start",
          title: "What the example begins with",
          body: `The page starts with this move: ${baseExample}`,
        },
        {
          label: "Leap",
          title: "Where the reasoning slips",
          body: `The problem appears when ${lowerFirst(problem)}`,
        },
        {
          label: "Cost",
          title: "What that slip hides or distorts",
          body: notes || "The conclusion is being made easier to accept before the missing reasoning work has been done.",
        },
        {
          label: "Repair",
          title: "What to ask instead",
          body: repairPrompt,
        },
      ],
    },
  };

  return {
    ...presets[mode],
    highlight: ensureSentence(record.feedbackLogic || primaryProfile.feedback || record.notes || ""),
  };
}

function buildReasoningAudit(record, prompt, primaryProfile) {
  const tabKey = `audit-${record.slug}`;
  const preset = buildAuditPreset(record, prompt, primaryProfile);
  const buttons = preset.steps
    .map(
      (step, index) => `<button
        class="audit-step-button${index === 0 ? " active" : ""}"
        type="button"
        role="tab"
        id="${tabKey}-button-${index}"
        aria-controls="${tabKey}-panel-${index}"
        aria-selected="${index === 0 ? "true" : "false"}"
        data-audit-button
      ><span class="audit-step-index">${index + 1}</span><span class="audit-step-label">${escapeHtml(step.label)}</span></button>`,
    )
    .join("");

  const panels = preset.steps
    .map(
      (step, index) => `<section
        class="audit-panel${index === 0 ? " active" : ""}"
        id="${tabKey}-panel-${index}"
        role="tabpanel"
        aria-labelledby="${tabKey}-button-${index}"
        ${index === 0 ? "" : "hidden"}
        data-audit-panel
      >
        <p class="audit-panel-title">${escapeHtml(step.title)}</p>
        <p class="audit-panel-body">${escapeHtml(step.body)}</p>
      </section>`,
    )
    .join("");

  return `<div class="audit-widget" data-audit-widget>
    <div class="audit-header">
      <h4>${escapeHtml(preset.title)}</h4>
      <p class="muted">${escapeHtml(preset.intro)}</p>
    </div>
    <div class="audit-step-list" role="tablist" aria-label="${escapeHtml(record.name)} reasoning walkthrough">
      ${buttons}
    </div>
    <div class="audit-progress-track" aria-hidden="true">
      <div class="audit-progress-fill" data-audit-progress></div>
    </div>
    <div class="audit-panel-wrap">
      ${panels}
    </div>
    ${preset.highlight ? `<p class="muted audit-feedback"><strong>What this walkthrough highlights:</strong> ${escapeHtml(preset.highlight)}</p>` : ""}
    <p class="muted lab-example"><strong>Current example:</strong> ${escapeHtml(record.example)}</p>
  </div>`;
}

function buildAuditSupportPanels(record, prompt, primaryProfile) {
  const preset = buildAuditPreset(record, prompt, primaryProfile);
  const stepLabels = preset.steps.map((step) => step.label).join(" -> ");
  const flowText = ensureSentence(
    `It moves through ${preset.steps[0].label}, ${preset.steps[1].label}, ${preset.steps[2].label}, and ${preset.steps[3].label} so the reader can follow the example from setup to diagnosis to repair`,
  );
  const usageText = ensureSentence(
    `Move through ${stepLabels} in order, compare each stage to the current example, and stop as soon as the later step adds more than the earlier one has earned`,
  );

  return `<div class="two-column compact-columns">
    <div class="note-panel">
      <h4>What this walkthrough follows</h4>
      <p class="muted">${escapeHtml(`${preset.intro} ${flowText}`)}</p>
    </div>
    <div class="note-panel">
      <h4>Where the slip happens</h4>
      <p class="muted">${escapeHtml(preset.steps[1].body)}</p>
    </div>
    <div class="note-panel">
      <h4>How to use the stages</h4>
      <p class="muted">${escapeHtml(usageText)}</p>
    </div>
    <div class="note-panel">
      <h4>What the repair step asks</h4>
      <p class="muted">${escapeHtml(preset.steps[3].body)}</p>
    </div>
  </div>`;
}

function normalizeRecordCategories(record) {
  const categories = [];
  const seen = new Set();

  for (const category of record.categories || []) {
    if (!category || seen.has(category)) continue;
    if (!categoryDescriptions[category]) {
      throw new Error(`Record "${record.name}" has unknown category "${category}".`);
    }
    seen.add(category);
    categories.push(category);
  }

  if (categories.length === 0) {
    throw new Error(`Record "${record.name}" is missing categories.`);
  }

  if (categories.length > 3) {
    throw new Error(`Record "${record.name}" has ${categories.length} categories; expected at most 3.`);
  }

  return categories;
}

function absoluteUrl(relativePath = "") {
  return new URL(relativePath, siteUrl).toString();
}

function buildSitemap(entries) {
  const body = entries
    .map(
      (entry) => `  <url>
    <loc>${escapeHtml(absoluteUrl(entry.path))}</loc>
    <lastmod>${buildDate}</lastmod>
  </url>`,
    )
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${body}
</urlset>
`;
}

function jsonLdMarkup(items = []) {
  const list = Array.isArray(items) ? items.filter(Boolean) : [items].filter(Boolean);
  return list
    .map(
      (item) =>
        `<script type="application/ld+json">${JSON.stringify(item).replace(/</g, "\\u003c")}</script>`,
    )
    .join("\n    ");
}

function publisherSchema() {
  return {
    "@type": "Person",
    name: "Phil Stilwell",
    url: absoluteUrl("about/"),
  };
}

function breadcrumbSchema(items = []) {
  if (!items.length) return null;
  return {
    "@context": "https://schema.org",
    "@type": "BreadcrumbList",
    itemListElement: items.map((item, index) => ({
      "@type": "ListItem",
      position: index + 1,
      name: item.name,
      item: absoluteUrl(item.path),
    })),
  };
}

function mergeKeywords(...groups) {
  return [...new Set(
    groups
      .flatMap((group) => (Array.isArray(group) ? group : [group]))
      .map((item) => String(item || "").trim())
      .filter(Boolean),
  )];
}

function learningResourceSchema({
  name,
  path,
  description,
  about = [],
  teaches = [],
  learningResourceType = ["Reference"],
  educationalUse = ["teaching", "self-study"],
  audienceType = "",
  keywords = [],
}) {
  const aboutItems = (Array.isArray(about) ? about : [about]).filter(Boolean);
  const teachesItems = (Array.isArray(teaches) ? teaches : [teaches]).filter(Boolean);
  const keywordItems = mergeKeywords(keywords);

  return {
    "@context": "https://schema.org",
    "@type": "LearningResource",
    name,
    url: absoluteUrl(path),
    description,
    inLanguage: "en-US",
    isAccessibleForFree: true,
    publisher: publisherSchema(),
    learningResourceType,
    educationalUse,
    audience: audienceType
      ? {
          "@type": "Audience",
          audienceType,
        }
      : undefined,
    about: aboutItems.length === 1 ? aboutItems[0] : aboutItems.length ? aboutItems : undefined,
    teaches: teachesItems.length === 1 ? teachesItems[0] : teachesItems.length ? teachesItems : undefined,
    keywords: keywordItems.length ? keywordItems.join(", ") : undefined,
  };
}

function pageShell({
  title,
  description,
  prefix,
  content,
  currentSection = "",
  canonicalPath = "",
  ogType = "website",
  robots = "index,follow",
  extraHeadHtml = cloudflareWebAnalyticsTag,
  structuredData = [],
  socialImageAlt = "LogFall logo",
  keywords = [],
}) {
  const homeHref = prefix || "./";
  const navItems = [
    { href: homeHref, label: "Home", key: "home" },
    { href: `${prefix}fallacies/`, label: "All Fallacies", key: "fallacies" },
    { href: `${prefix}categories/`, label: "Categories", key: "categories" },
    { href: `${prefix}check-yourself/`, label: "Check Yourself", key: "check-yourself" },
    { href: `${prefix}prompts/`, label: "Fun AI Prompts", key: "prompts" },
    { href: `${prefix}about/`, label: "About", key: "about" },
  ];

  const nav = navItems
    .map(
      (item) =>
        `<a href="${item.href}"${item.key === currentSection ? ' aria-current="page"' : ""}>${escapeHtml(item.label)}</a>`,
    )
    .join("");
  const canonicalUrl = absoluteUrl(canonicalPath);
  const socialImageUrl = absoluteUrl(socialImagePath);
  const structuredDataHead = jsonLdMarkup(structuredData);
  const keywordContent = mergeKeywords(baseSiteKeywords, keywords).join(", ");

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(title)}</title>
    <meta name="description" content="${escapeHtml(description)}" />
    <meta name="robots" content="${escapeHtml(robots)}" />
    <meta name="author" content="Phil Stilwell" />
    <meta name="creator" content="Phil Stilwell" />
    <meta name="application-name" content="LogFall" />
    <meta name="keywords" content="${escapeHtml(keywordContent)}" />
    <link rel="canonical" href="${escapeHtml(canonicalUrl)}" />
    <link rel="alternate" hreflang="en-US" href="${escapeHtml(canonicalUrl)}" />
    <link rel="alternate" hreflang="x-default" href="${escapeHtml(canonicalUrl)}" />
    <link rel="icon" type="image/x-icon" href="${prefix}assets/favicon.ico" />
    <link rel="icon" type="image/png" sizes="32x32" href="${prefix}assets/favicon-32x32.png" />
    <link rel="icon" type="image/png" sizes="16x16" href="${prefix}assets/favicon-16x16.png" />
    <link rel="apple-touch-icon" sizes="180x180" href="${prefix}assets/apple-touch-icon.png" />
    <link rel="manifest" href="${prefix}site.webmanifest" />
    <meta name="theme-color" content="#0f172a" />
    <meta property="og:site_name" content="LogFall" />
    <meta property="og:locale" content="en_US" />
    <meta property="og:type" content="${escapeHtml(ogType)}" />
    <meta property="og:title" content="${escapeHtml(title)}" />
    <meta property="og:description" content="${escapeHtml(description)}" />
    <meta property="og:url" content="${escapeHtml(canonicalUrl)}" />
    <meta property="og:image" content="${escapeHtml(socialImageUrl)}" />
    <meta property="og:image:secure_url" content="${escapeHtml(socialImageUrl)}" />
    <meta property="og:image:type" content="${socialImageType}" />
    <meta property="og:image:width" content="${socialImageWidth}" />
    <meta property="og:image:height" content="${socialImageHeight}" />
    <meta property="og:image:alt" content="${escapeHtml(socialImageAlt)}" />
    <meta name="twitter:card" content="summary" />
    <meta name="twitter:title" content="${escapeHtml(title)}" />
    <meta name="twitter:description" content="${escapeHtml(description)}" />
    <meta name="twitter:image" content="${escapeHtml(socialImageUrl)}" />
    <meta name="twitter:image:alt" content="${escapeHtml(socialImageAlt)}" />
    ${structuredDataHead}
    ${extraHeadHtml}
    <link rel="stylesheet" href="${prefix}styles.css" />
    <script defer src="${prefix}app.js"></script>
  </head>
  <body>
    <div class="site-shell">
      <header class="masthead">
        <div class="masthead-inner">
          <div class="brand-row">
            <div class="brand-lockup">
              <img class="brand-logo" src="${prefix}assets/logo.jpg" alt="LogFall logo" />
              <div>
                <p class="brand-kicker">Logical Fallacies</p>
                <h1 class="brand-title">LogFall</h1>
                <p class="brand-subtitle">A practical logical-fallacies reference with clear explanations, usable examples, and teaching tools.</p>
              </div>
            </div>
          </div>
          <nav class="top-nav" aria-label="Primary">${nav}</nav>
        </div>
      </header>
      <main class="page-wrap">${content}</main>
      <footer class="footer">
        <div class="footer-inner">
          <p class="footer-note">A reference for spotting, naming, comparing, and repairing reasoning errors.</p>
          <p class="footer-note">${escapeHtml(copyrightNotice)}</p>
        </div>
      </footer>
    </div>
  </body>
</html>`;
}

function renderPills(categories) {
  return `<div class="pill-row">${categories
    .map((category) => `<span class="pill pill-${escapeHtml(category)}">${escapeHtml(category)}</span>`)
    .join("")}</div>`;
}

function renderTeacherPills(record) {
  const pedagogy = pedagogyForRecord(record);
  return `<div class="teaching-pill-row">
    <span class="teaching-pill">${escapeHtml(pedagogy.difficulty)}</span>
    <span class="teaching-pill">${escapeHtml(pedagogy.classroomLevel)}</span>
  </div>`;
}

function renderRhetoricGaugeCard(gauge) {
  return `<article class="gauge-card">
    <div class="gauge-card-top">
      <p class="gauge-kicker">${escapeHtml(gauge.band)}</p>
      <p class="gauge-score">${gauge.value}</p>
    </div>
    <h4>${escapeHtml(gauge.title)}</h4>
    <p class="muted gauge-summary">${escapeHtml(gauge.summary)}</p>
    <div class="gauge-meter" style="--value:${gauge.value}">
      <div class="gauge-meter-fill"></div>
      <div class="gauge-meter-marker" aria-hidden="true"></div>
    </div>
    <div class="gauge-scale" aria-hidden="true">
      <span>${escapeHtml(gauge.lowLabel)}</span>
      <span>${escapeHtml(gauge.highLabel)}</span>
    </div>
  </article>`;
}

function renderRhetoricGaugeSection(record) {
  const gauges = rhetoricGaugesForRecord(record);
  return `<div class="section-block gauge-section">
    <p class="detail-card-label">Teaching gauges</p>
    <p class="muted gauge-note">0-100 editorial estimates for classroom use rather than measured statistics.</p>
    <div class="gauge-grid">
      ${renderRhetoricGaugeCard(gauges.common)}
      ${renderRhetoricGaugeCard(gauges.spot)}
      ${renderRhetoricGaugeCard(gauges.innocent)}
    </div>
  </div>`;
}

function renderFallacyCard(record, prefix) {
  const pedagogy = pedagogyForRecord(record);
  const aliases = record.aliases.join(" ");
  const caseStudyText = record.caseStudies.map((item) => normalizeCaseStudy(item).summary).join(" ");
  const body = `${record.definition} ${record.example} ${record.notes} ${caseStudyText} ${pedagogy.classroomTags.join(" ")} ${pedagogy.teachingPaths.map((item) => item.title).join(" ")}`;
  return `<article
    class="fallacy-card"
    data-fallacy-card
    data-name="${escapeHtml(record.name)}"
    data-aliases="${escapeHtml(aliases)}"
    data-categories="${escapeHtml(record.categories.join("|"))}"
    data-difficulty="${escapeHtml(pedagogy.difficulty)}"
    data-classroom="${escapeHtml(pedagogy.classroomLevel)}"
    data-body="${escapeHtml(body)}"
  >
    <h3><a href="${prefix}fallacies/${record.slug}/">${escapeHtml(record.name)}</a></h3>
    <p class="card-copy">${escapeHtml(truncate(record.definition, 170))}</p>
    ${renderPills(record.categories)}
    ${renderTeacherPills(record)}
  </article>`;
}

function renderAssessmentCard(record, prefix) {
  const pedagogy = pedagogyForRecord(record);
  const aliases = record.aliases.join(" ");
  const caseStudyText = record.caseStudies.map((item) => normalizeCaseStudy(item).summary).join(" ");
  const body = `${record.definition} ${record.example} ${record.notes} ${caseStudyText} ${pedagogy.classroomTags.join(" ")} ${pedagogy.teachingPaths.map((item) => item.title).join(" ")}`;
  return `<article
    class="fallacy-card assessment-card"
    data-fallacy-card
    data-name="${escapeHtml(record.name)}"
    data-aliases="${escapeHtml(aliases)}"
    data-categories="${escapeHtml(record.categories.join("|"))}"
    data-difficulty="${escapeHtml(pedagogy.difficulty)}"
    data-classroom="${escapeHtml(pedagogy.classroomLevel)}"
    data-body="${escapeHtml(body)}"
  >
    <p class="eyebrow">Check yourself</p>
    <h3><a href="${prefix}check-yourself/${record.slug}/">${escapeHtml(record.name)}</a></h3>
    <p class="card-copy"><strong>Example claim:</strong> ${escapeHtml(truncate(record.example, 165))}</p>
    ${renderPills(record.categories)}
    ${renderTeacherPills(record)}
    <p class="assessment-card-link"><a class="inline-link" href="${prefix}check-yourself/${record.slug}/">Open the quiz</a></p>
  </article>`;
}

function posterAssetNameForRecord(record) {
  return `fallacy-${record.slug}-poster`;
}

function resolvePosterAssetForRecord(record, posterAssets) {
  const base = posterAssetNameForRecord(record);
  for (const extension of ["webp", "png", "jpeg", "jpg"]) {
    const candidate = `${base}.${extension}`;
    if (posterAssets.has(candidate)) {
      return candidate;
    }
  }
  return null;
}

function posterAltTextForRecord(record) {
  return `${record.name} companion poster in a retro mid-century editorial cartoon style, illustrating ${record.definition.toLowerCase()}`;
}

function stripTrailingPunctuation(value = "") {
  return String(value)
    .trim()
    .replace(/([.?!])(["'”’])$/u, "$2")
    .replace(/[.?!]+$/u, "");
}

function lowercaseFirst(value = "") {
  if (!value) return "";
  return value.charAt(0).toLowerCase() + value.slice(1);
}

function posterExplanationForRecord(record) {
  const example = stripTrailingPunctuation(record.example || "");
  const definition = stripTrailingPunctuation(
    String(record.definition || "")
      .replace(/^Occurs when\s+/i, "")
      .replace(/^This fallacy occurs when\s+/i, ""),
  );

  const parts = [];

  if (example) {
    parts.push(`This image turns the fallacy into a concrete scene: ${example}.`);
  }

  if (definition) {
    parts.push(`The key point is that ${lowercaseFirst(definition)}.`);
  }

  return parts.join(" ");
}

function renderReferenceMeta(record, prompts) {
  const pedagogy = pedagogyForRecord(record);
  const pathMarkup = pedagogy.teachingPaths.length
    ? `<div class="path-link-row">
        ${pedagogy.teachingPaths
          .map(
            (pathMeta) => `<a class="path-link-chip" href="../../paths/${escapeHtml(pathMeta.slug)}/">${escapeHtml(pathMeta.title)}</a>`,
          )
          .join("")}
      </div>`
    : `<p class="muted">This entry is not currently in one of the curated teaching paths.</p>`;
  return `<div class="meta-grid">
    <div class="note-panel">
      <h4>Family</h4>
      <p class="muted">${escapeHtml(record.family || "Unspecified")}</p>
    </div>
    ${
      record.aliases.length
        ? `<div class="note-panel">
      <h4>Aliases</h4>
      <p class="muted">${escapeHtml(record.aliases.join(", "))}</p>
    </div>`
        : ""
    }
    <div class="note-panel">
      <h4>Quick check</h4>
      <p class="muted">${escapeHtml(prompts[0] || "Ask what evidence or reasoning step is doing too much work.")}</p>
    </div>
    <div class="note-panel">
      <h4>Difficulty</h4>
      <p class="muted">${escapeHtml(pedagogy.difficulty)}</p>
      <div class="teaching-pill-row">
        ${pedagogy.classroomTags.map((tag) => `<span class="teaching-pill">${escapeHtml(tag)}</span>`).join("")}
      </div>
    </div>
    <div class="note-panel">
      <h4>Teaching paths</h4>
      ${pathMarkup}
    </div>
  </div>`;
}

function renderPosterIllustration(record, prefix, posterAssets) {
  const asset = resolvePosterAssetForRecord(record, posterAssets);
  if (!asset) return "";

  return `<aside class="detail-section detail-illustration-shell">
    <img
      class="detail-illustration-image"
      src="${prefix}assets/${escapeHtml(asset)}"
      alt="${escapeHtml(posterAltTextForRecord(record))}"
      loading="lazy"
    />
    <div class="detail-illustration-copy">
      <p class="detail-illustration-label">What this image shows</p>
      <p class="detail-illustration-text">${escapeHtml(posterExplanationForRecord(record))}</p>
    </div>
  </aside>`;
}

function relatedFallacies(record, records) {
  return records
    .filter((candidate) => candidate.name !== record.name)
    .map((candidate) => {
      const shared = candidate.categories.filter((category) => record.categories.includes(category));
      const score =
        shared.length * 10 +
        (candidate.family && candidate.family === record.family ? 2 : 0) +
        (candidate.subCategory && candidate.subCategory === record.subCategory ? 1 : 0);
      return { candidate, score, shared };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || a.candidate.name.localeCompare(b.candidate.name))
    .slice(0, 6)
    .map((item) => item.candidate);
}

function renderRationalityLab(record, categoryProfiles) {
  const primaryCategory = record.categories[0] || "";
  const primaryProfile = categoryProfiles[primaryCategory] || {};
  const tabKey = `lab-${record.slug}`;
  const tabs = [];

  if (record.rationalityDanger || record.mainReasoningProblem || primaryProfile.danger) {
    tabs.push({
      id: `${tabKey}-danger`,
      label: "Why it matters",
      content: `
        <div class="two-column compact-columns">
          ${
            record.rationalityDanger
              ? `<div class="note-panel">
            <h4>Why this mistake matters</h4>
            <p class="muted">${escapeHtml(record.rationalityDanger)}</p>
          </div>`
              : ""
          }
          ${
            record.mainReasoningProblem
              ? `<div class="note-panel">
            <h4>Main reasoning problem</h4>
            <p class="muted">${escapeHtml(record.mainReasoningProblem)}</p>
          </div>`
              : ""
          }
          ${
            primaryProfile.danger
              ? `<div class="note-panel">
            <h4>Why this kind of mistake matters</h4>
            <p class="muted">${escapeHtml(primaryProfile.danger)}</p>
          </div>`
              : ""
          }
        </div>`,
    });
  }

  if (record.dynamicsToNotice || record.warningSigns) {
    tabs.push({
      id: `${tabKey}-dynamics`,
      label: "What to watch for",
      content: `
        <div class="two-column compact-columns">
          ${
            record.dynamicsToNotice
              ? `<div class="note-panel">
            <h4>What to watch for</h4>
            <p class="muted">${escapeHtml(record.dynamicsToNotice)}</p>
          </div>`
              : ""
          }
          ${
            record.warningSigns
              ? `<div class="note-panel">
            <h4>Common warning signs</h4>
            <p class="muted">${escapeHtml(record.warningSigns)}</p>
          </div>`
              : ""
          }
        </div>`,
    });
  }

  if (record.repairPrompts) {
    tabs.push({
      id: `${tabKey}-repair`,
      label: "How to fix it",
      content: `<div class="note-panel">
        <h4>How to fix it</h4>
        <p class="muted">${escapeHtml(record.repairPrompts)}</p>
      </div>`,
    });
  }

  if (record.interactiveMechanic || record.userAction || record.feedbackLogic || primaryProfile.mechanic) {
    tabs.push({
      id: `${tabKey}-tool`,
      label: "Try it",
      content: buildAuditSupportPanels(
        record,
        diagnosticPrompts[primaryCategory] || diagnosticPrompts[record.categories[1]] || "",
        primaryProfile,
      ),
    });
    tabs[tabs.length - 1].content += buildReasoningAudit(record, diagnosticPrompts[primaryCategory] || diagnosticPrompts[record.categories[1]] || "", primaryProfile);
  }

  if (!tabs.length) {
    return "";
  }

  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">Practice And Repair</h3>
        <p class="section-copy">Extra teaching tools that show why the fallacy is persuasive, what to look for, and how to correct it.</p>
      </div>
    </div>
    <article class="detail-section">
      ${renderTabGroup(tabKey, tabs)}
    </article>
  </section>`;
}

function renderConfusionSection(record, records, prefix) {
  const confusions = confusionCandidates(record, records, 2);
  if (!confusions.length) return "";

  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">Often confused with</h3>
        <p class="section-copy">These near neighbors are easy to mix up, so use the comparison to see the exact difference.</p>
      </div>
    </div>
    <div class="two-column comparison-grid">
      ${confusions
        .map((item) => {
          const candidate = item.candidate;
          const candidatePrompt =
            diagnosticPrompts[candidate.categories[0]] ||
            "What specific reasoning step is being asked to do too much work?";
          return `<article class="note-panel comparison-card">
            <p class="eyebrow">Comparison</p>
            <h4><a href="${prefix}fallacies/${candidate.slug}/">${escapeHtml(candidate.name)}</a></h4>
            <p class="muted"><strong>Why people mix them up:</strong> ${escapeHtml(overlapTextForConfusion(item))}</p>
            <p class="muted"><strong>Exact difference:</strong> ${escapeHtml(record.name)} happens when ${escapeHtml(lowerFirst(definitionCore(record.definition)))}. ${escapeHtml(candidate.name)} happens when ${escapeHtml(lowerFirst(definitionCore(candidate.definition)))}.</p>
            <p class="muted"><strong>Quick split:</strong> ${escapeHtml(diagnosticPrompts[record.categories[0]] || "Ask what the real reasoning problem is.")} Then compare it with ${escapeHtml(candidatePrompt)}</p>
          </article>`;
        })
        .join("")}
    </div>
  </section>`;
}

function renderCaveatSection(record, records) {
  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">Caveat</h3>
        <p class="section-copy">This label is easy to overuse. The point here is not to call every weak argument by this name, but to reserve it for the exact misstep it describes.</p>
      </div>
    </div>
    <div class="two-column compact-columns">
      <div class="note-panel">
        <h4>Common misapplication</h4>
        <p class="muted">${escapeHtml(caveatMisuseText(record))}</p>
      </div>
      <div class="note-panel">
        <h4>Use the label only when...</h4>
        <p class="muted">${escapeHtml(caveatApplyText(record, records))}</p>
      </div>
    </div>
  </section>`;
}

function renderArgumentMapSection(record) {
  const map = argumentMapForRecord(record);
  if (!map) return "";

  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">Visual argument map</h3>
        <p class="section-copy">${escapeHtml(map.intro)}</p>
      </div>
    </div>
    <article class="detail-section">
      <div class="argument-map argument-map-${escapeHtml(map.variant)}">
        ${map.nodes
          .map(
            (node) => `<div class="argument-node">
              <p class="argument-node-label">${escapeHtml(node.label)}</p>
              <p class="argument-node-text">${escapeHtml(node.text)}</p>
            </div>`,
          )
          .join("")}
      </div>
    </article>
  </section>`;
}

function renderQuizAndRepairSection(record, records, options = {}) {
  const { includeHeader = true } = options;
  const quiz = quizConfigForRecord(record, records);
  const repairChecklist = repairChecklistForRecord(record);
  const repairModel = repairModelForRecord(record);
  const headerMarkup = includeHeader
    ? `<div class="section-header">
      <div>
        <h3 class="section-title">Check yourself</h3>
        <p class="section-copy">Use this short quiz to identify the fallacy, explain it, and then practice repairing the claim.</p>
      </div>
    </div>`
    : "";

  return `<section class="section-block">
    ${headerMarkup}
    <div class="two-column quiz-grid">
      <article
        class="detail-section quiz-card"
        data-quiz-widget
        data-quiz-answer="${escapeHtml(quiz.answer)}"
        data-quiz-keywords="${escapeHtml(JSON.stringify(quiz.keywords))}"
        data-quiz-model="${escapeHtml(quiz.model)}"
      >
        <h4>1. Identify the fallacy</h4>
        <p class="muted">Choose the best label for the example claim below. Base your answer on this exact wording, not on a broader topic guess.</p>
        <div class="quiz-example-shell">
          <p class="quiz-example-label">Example claim to diagnose</p>
          <p class="quiz-example-text">${escapeHtml(record.example)}</p>
          <p class="quiz-example-note">Read this claim carefully. Every answer in this check should be anchored to this exact example.</p>
        </div>
        <div class="quiz-options">
          ${quiz.options
            .map(
              (option, index) => `<label class="quiz-option">
                <input type="radio" name="quiz-${escapeHtml(record.slug)}" value="${escapeHtml(option)}" ${index === 0 ? "" : ""} />
                <span>${escapeHtml(option)}</span>
              </label>`,
            )
            .join("")}
        </div>
        <h4>2. Explain why</h4>
        <p class="muted">In one or two sentences, name the exact reasoning slip.</p>
        <textarea class="quiz-textarea" data-quiz-response placeholder="Explain what makes the reasoning fallacious..."></textarea>
        <button class="button button-primary button-compact" type="button" data-quiz-grade>Grade this check</button>
        <div class="quiz-feedback hidden" data-quiz-feedback role="status" aria-live="polite"></div>
      </article>

      <article class="detail-section repair-card">
        <h4>Repair the argument</h4>
        <p class="muted">Rewrite the example so it says only what the evidence or reasoning has actually earned.</p>
        <textarea class="quiz-textarea" placeholder="Write a stronger, fairer version of the claim..."></textarea>
        <div class="repair-checklist">
          <p class="repair-checklist-title">What a good repair should do</p>
          <ul>
            ${repairChecklist.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
          </ul>
        </div>
        <details class="repair-model">
          <summary>Show one reasonable repair</summary>
          <p class="muted">${escapeHtml(repairModel)}</p>
        </details>
      </article>
    </div>
  </section>`;
}

function renderAssessmentTeaserSection(record) {
  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">Check yourself</h3>
        <p class="section-copy">The assessment area now uses mixed 10-question sets, so the fallacy is not announced in the title before the quiz begins.</p>
      </div>
    </div>
    <article class="detail-section">
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>What the assessment does</h4>
          <p class="muted">You will work through a mixed set of fallacy-identification questions. Focused links from a fallacy page will quietly include this fallacy among nearby look-alikes without announcing the answer in the page title.</p>
        </div>
        <div class="note-panel">
          <h4>Open the check</h4>
          <p class="muted"><a class="text-link" href="../../check-yourself/?focus=${escapeHtml(record.slug)}">Open a mixed Check Yourself set that includes this fallacy</a>.</p>
        </div>
      </div>
    </article>
  </section>`;
}

function recordsForTeachingPath(pathDefinition, records) {
  return pathDefinition.names
    .map((name) => records.find((record) => record.name === name))
    .filter(Boolean);
}

function renderTeachingPathCard(pathDefinition, records, prefix) {
  const members = recordsForTeachingPath(pathDefinition, records);
  return `<article class="category-card path-card">
    <p class="eyebrow">Teaching path</p>
    <h3><a href="${prefix}paths/${pathDefinition.slug}/">${escapeHtml(pathDefinition.title)}</a></h3>
    <p class="card-copy">${escapeHtml(pathDefinition.description)}</p>
    <div class="teaching-pill-row">
      <span class="teaching-pill">${members.length} fallacies</span>
      <span class="teaching-pill">${escapeHtml(pathDefinition.audience)}</span>
    </div>
  </article>`;
}

function homeSeoDescription() {
  return "Explore logical fallacies with clear definitions, examples, case studies, teaching paths, AI prompts, and classroom-ready critical thinking tools.";
}

function fallaciesIndexSeoDescription(recordCount) {
  return `Browse all ${recordCount} logical fallacies by name, category, difficulty, and classroom level, with definitions, examples, related entries, and teaching tools.`;
}

function categorySeoDescription(category) {
  return `Explore ${category.count} ${category.name.toLowerCase()} logical fallacies with definitions, examples, related entries, and teaching tools for argument analysis.`;
}

function promptsSeoDescription() {
  return "Copy AI prompts for identifying logical fallacies in passages, debates, media, and argument maps, with direct links back to LogFall.";
}

function assessmentIndexSeoDescription(recordCount) {
  return `Take mixed 10-question logical fallacy assessments built from ${recordCount} fallacies, with clear example claims, answer feedback, and links back to the relevant LogFall entries.`;
}

function seoFallacyName(record) {
  return /\bfallacy\b/i.test(record.name) ? record.name : `${record.name} logical fallacy`;
}

function seoFallacyKeyword(record) {
  return /\bfallacy\b/i.test(record.name) ? record.name : `${record.name} fallacy`;
}

function assessmentSeoTitle(record) {
  return `${seoFallacyName(record)}: check-yourself quiz and repair practice | LogFall`;
}

function assessmentSeoDescription(record) {
  return truncate(
    `Take the ${seoFallacyName(record)} check-yourself quiz with a clear example claim, fallacy identification, explanation grading, and repair practice.`,
    158,
  );
}

function buildAssessmentBank(records) {
  return records.map((record) => {
    const quiz = quizConfigForRecord(record, records);
    const pedagogy = pedagogyForRecord(record);
    return {
      slug: record.slug,
      name: record.name,
      example: record.example,
      answer: quiz.answer,
      options: quiz.options,
      model: quiz.model,
      categories: record.categories,
      difficulty: pedagogy.difficulty,
      classroomLevel: pedagogy.classroomLevel,
      fallacyUrl: `../fallacies/${record.slug}/`,
    };
  });
}

function safeJsonForScript(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026");
}

function aboutSeoDescription() {
  return "Meet Phil Stilwell and learn how LogFall grew from university critical-thinking classes into a teaching-focused logical fallacies reference.";
}

function pathSeoDescription(pathDefinition, memberCount) {
  return `${pathDefinition.title}: a ${memberCount}-fallacy teaching path for ${pathDefinition.audience.toLowerCase()}, built for comparison, discussion, and critical thinking instruction.`;
}

function fallacySeoTitle(record) {
  return `${seoFallacyName(record)}: definition, examples, and how to spot it | LogFall`;
}

function fallacySeoDescription(record) {
  return truncate(
    `Learn ${seoFallacyName(record)} with a clear definition, example, case studies, related fallacies, and teaching tools for critical thinking.`,
    158,
  );
}

function buildTeachingPathsIndexPage(records) {
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Teaching Paths</strong>
    </div>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h2 class="section-title">Teaching paths</h2>
          <p class="section-copy">Curated routes through the taxonomy for teachers, discussion leaders, and first-time readers.</p>
        </div>
      </div>
      <div class="category-grid">
        ${teachingPathDefinitions.map((pathDefinition) => renderTeachingPathCard(pathDefinition, records, "../")).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: "Teaching Paths for Learning Logical Fallacies | LogFall",
    description: "Curated teaching paths through logical fallacies for classrooms, review sessions, and first-time readers.",
    prefix: "../",
    currentSection: "",
    canonicalPath: "paths/",
    keywords: [
      "logical fallacy teaching paths",
      "critical thinking lessons",
      "fallacy lesson sequence",
      "classroom logical fallacies",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Teaching Paths", path: "paths/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Teaching Paths",
        url: absoluteUrl("paths/"),
        description: "Curated teaching paths through logical fallacies for classrooms, review sessions, and first-time readers.",
        publisher: publisherSchema(),
        hasPart: teachingPathDefinitions.map((pathDefinition) => ({
          "@type": "CreativeWork",
          name: pathDefinition.title,
          url: absoluteUrl(`paths/${pathDefinition.slug}/`),
          description: pathDefinition.description,
        })),
      },
      learningResourceSchema({
        name: "Teaching Paths for Learning Logical Fallacies",
        path: "paths/",
        description: "Curated teaching paths through logical fallacies for classrooms, review sessions, and first-time readers.",
        about: ["logical fallacies", "critical thinking", "argument analysis"],
        teaches: ["logical fallacies", "comparison of reasoning mistakes"],
        learningResourceType: ["Teaching path", "Reference"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacy teaching paths", "critical thinking lessons", "fallacy lesson sequence"],
      }),
    ],
    content,
  });
}

function buildTeachingPathPage(pathDefinition, records) {
  const members = recordsForTeachingPath(pathDefinition, records);
  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Teaching Paths</a><span>/</span><strong>${escapeHtml(pathDefinition.title)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Teaching path</p>
      <h2 class="detail-title">${escapeHtml(pathDefinition.title)}</h2>
      <p class="detail-deck">${escapeHtml(pathDefinition.description)}</p>
      <div class="meta-grid section-block">
        <div class="note-panel">
          <h4>Best for</h4>
          <p class="muted">${escapeHtml(pathDefinition.audience)}</p>
        </div>
        <div class="note-panel">
          <h4>Sequence size</h4>
          <p class="muted">${members.length} fallacies in recommended teaching order.</p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Recommended sequence</h3>
          <p class="section-copy">Use the order below as a lesson sequence, review set, or comparison track.</p>
        </div>
      </div>
      <div class="fallacy-grid">
        ${members.map((record) => renderFallacyCard(record, "../../")).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: `${pathDefinition.title}: Logical Fallacy Teaching Path | LogFall`,
    description: pathSeoDescription(pathDefinition, members.length),
    prefix: "../../",
    currentSection: "",
    canonicalPath: `paths/${pathDefinition.slug}/`,
    keywords: [
      pathDefinition.title,
      "logical fallacy teaching path",
      "critical thinking sequence",
      pathDefinition.audience,
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Teaching Paths", path: "paths/" },
        { name: pathDefinition.title, path: `paths/${pathDefinition.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: pathDefinition.title,
        url: absoluteUrl(`paths/${pathDefinition.slug}/`),
        description: pathSeoDescription(pathDefinition, members.length),
        publisher: publisherSchema(),
        audience: {
          "@type": "Audience",
          audienceType: pathDefinition.audience,
        },
        mainEntity: {
          "@type": "ItemList",
          numberOfItems: members.length,
          itemListElement: members.map((record, index) => ({
            "@type": "ListItem",
            position: index + 1,
            name: record.name,
            url: absoluteUrl(`fallacies/${record.slug}/`),
          })),
        },
      },
      learningResourceSchema({
        name: `${pathDefinition.title}: Logical Fallacy Teaching Path`,
        path: `paths/${pathDefinition.slug}/`,
        description: pathSeoDescription(pathDefinition, members.length),
        about: ["logical fallacies", "critical thinking", "argument comparison"],
        teaches: members.map((record) => record.name),
        learningResourceType: ["Teaching path", "Lesson sequence"],
        educationalUse: ["teaching", "review", "self-study"],
        audienceType: pathDefinition.audience,
        keywords: [pathDefinition.title, "logical fallacy teaching path", pathDefinition.audience],
      }),
    ],
    content,
  });
}

function buildHomePage(records, categories) {
  const caseStudyCount = records.reduce((sum, record) => sum + record.caseStudies.length, 0);
  const featured = featuredNames
    .map((name) => records.find((record) => record.name === name))
    .filter(Boolean);

  const content = `
    <section class="hero">
      <div class="hero-panel">
        <h2 class="hero-title">Logical fallacies made clearer, more teachable, and easier to compare.</h2>
        <p class="hero-lead">
          LogFall is a practical reference for spotting, naming, comparing, and correcting reasoning mistakes.
          Each entry combines a definition, a concrete example, case studies, a companion illustration,
          related fallacies, and a guided practice tool.
        </p>
        <div class="hero-actions">
          <a class="button button-primary" href="fallacies/">Browse All Fallacies</a>
          <a class="button button-secondary" href="categories/">Explore Categories</a>
          <a class="button button-secondary" href="about/">How To Use LogFall</a>
        </div>
      </div>
      <aside class="hero-side hero-panel">
        <p class="eyebrow">At a glance</p>
        <div class="stat-grid">
          <div class="stat-card">
            <span class="stat-value">${records.length}</span>
            <span class="stat-label">Fallacies</span>
          </div>
          <div class="stat-card">
            <span class="stat-value">${categories.length}</span>
            <span class="stat-label">Taxonomy categories</span>
          </div>
          <div class="stat-card">
            <span class="stat-value">${caseStudyCount}</span>
            <span class="stat-label">Case studies</span>
          </div>
        </div>
        <div class="section-block">
          <p class="eyebrow">How to use the site</p>
          <div class="note-panel">
            <h4>1. Start with a category</h4>
            <p class="muted">Browse the taxonomy when you know the kind of reasoning error but not the exact fallacy name.</p>
          </div>
          <div class="note-panel" style="margin-top:12px;">
            <h4>2. Jump to a detail page</h4>
            <p class="muted">Each fallacy page brings together a concise definition, a concrete example, explanatory notes, a rationality lab, case studies, and nearby entries.</p>
          </div>
          <div class="note-panel" style="margin-top:12px;">
            <h4>3. Compare similar mistakes</h4>
            <p class="muted">Use the related fallacies, case studies, and practice tool to separate look-alike mistakes that often get confused in real arguments.</p>
          </div>
        </div>
      </aside>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h2 class="section-title">Teaching paths</h2>
          <p class="section-copy">Curated routes through the site for first-time readers, public-debate analysis, and high-confusion review sessions.</p>
        </div>
        <a class="inline-link" href="paths/">See all paths</a>
      </div>
      <div class="category-grid">
        ${teachingPathDefinitions.map((pathDefinition) => renderTeachingPathCard(pathDefinition, records, "")).join("")}
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h2 class="section-title">Browse the taxonomy</h2>
          <p class="section-copy">The taxonomy groups fallacies by the main kind of reasoning failure involved, which makes nearby mistakes easier to compare.</p>
        </div>
      </div>
      <div class="category-grid">
        ${categories
          .map(
            (category) => `<article class="category-card">
              <h3><a href="categories/${category.slug}/">${escapeHtml(category.name)}</a></h3>
              <p class="card-copy">${escapeHtml(category.description)}</p>
              <div class="pill-row">
                <span class="pill pill-${escapeHtml(category.name)}">${category.count} entries</span>
              </div>
            </article>`,
          )
          .join("")}
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h2 class="section-title">Featured fallacies</h2>
          <p class="section-copy">A starter set of especially common reasoning errors that are useful for classrooms, debates, and self-audit.</p>
        </div>
        <a class="inline-link" href="fallacies/">See the full index</a>
      </div>
      <div class="fallacy-grid">
        ${featured.map((record) => renderFallacyCard(record, "")).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: "Logical Fallacies Explained: Definitions, Examples, and Teaching Tools | LogFall",
    description: homeSeoDescription(),
    prefix: "",
    currentSection: "home",
    canonicalPath: "",
    keywords: [
      "logical fallacies list",
      "critical thinking",
      "fallacy examples",
      "argument analysis",
      "reasoning errors",
      "teaching tools",
    ],
    structuredData: [
      {
        "@context": "https://schema.org",
        "@type": "WebSite",
        name: "LogFall",
        url: absoluteUrl(""),
        description: homeSeoDescription(),
        publisher: publisherSchema(),
        inLanguage: "en-US",
        potentialAction: {
          "@type": "SearchAction",
          target: `${absoluteUrl("fallacies/")}?q={search_term_string}`,
          "query-input": "required name=search_term_string",
        },
      },
      learningResourceSchema({
        name: "LogFall",
        path: "",
        description: homeSeoDescription(),
        about: ["logical fallacies", "critical thinking", "argument analysis"],
        teaches: ["logical fallacies", "how to spot reasoning errors", "comparison of similar fallacies"],
        learningResourceType: ["Reference", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacies list", "critical thinking", "fallacy examples", "teaching tools"],
      }),
    ],
    content,
  });
}

function buildAboutPage() {
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>About</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">About LogFall</p>
      <h2 class="detail-title">A teaching-focused reference built by Phil Stilwell and his students.</h2>
      <p class="detail-deck">
        LogFall grew out of years of teaching critical thinking, logic, and argument analysis in university classrooms.
        It is designed to help readers recognize reasoning mistakes, distinguish near neighbors, and practice clearer habits of interpretation, comparison, and repair.
      </p>
      <div class="about-profile-grid section-block">
        <div class="note-panel about-portrait-panel">
          <img class="about-portrait" src="../assets/phil-stilwell-profile.png" alt="Phil Stilwell" />
          <p class="muted about-portrait-caption">
            Phil Stilwell is an essentially retired academic consultant, researcher, and university instructor whose work has centered on philosophy,
            epistemology, critical thinking, and macroeconomic theory.
          </p>
        </div>
        <div class="about-card-stack">
          <div class="note-panel">
            <h4>About Phil Stilwell</h4>
            <p class="muted about-summary">
              Phil Stilwell&rsquo;s career has spanned more than 26 years, much of it in Japan, across university instruction,
              academic consulting, and research. His main areas of focus include philosophy of science, epistemology, induction,
              analytic philosophy, applied logic, and critical thinking, with continuing work connected to
              <a class="text-link" href="https://credencing.com">Credencing</a>.
              He also remains engaged with questions in cognitive science, economic modeling, and technological change.
            </p>
            <ul class="about-credential-list">
              <li><strong>BA Philosophy, 1996</strong> &mdash; Summa Cum Laude, The University of Kansas</li>
              <li><strong>MA Education, 1998</strong> &mdash; The University of Kansas</li>
              <li><strong>Teaching focus</strong> &mdash; critical thinking, philosophy, epistemology, macroeconomics, and technical writing</li>
              <li><strong>Professional base</strong> &mdash; university instruction and academic consulting, primarily in Japan</li>
            </ul>
          </div>
          <div class="two-column compact-columns">
            <div class="note-panel">
              <h4>Academic background</h4>
              <p class="muted">
                Phil has designed and taught university-level philosophy curricula, including a general survey course with emphasis on analytic philosophy and existentialism.
                His research interests have centered on epistemology, induction, the philosophy of science, and the structure of rational belief.
              </p>
            </div>
            <div class="note-panel">
              <h4>Applied logic and instruction</h4>
              <p class="muted">
                A central aim of his teaching has been the cultivation of applied logic and critical analysis.
                He taught critical thinking courses for university and professional audiences, including work at the NYU School of Professional Studies in Tokyo,
                and later served as a Speech &amp; Debate Judge for the Japan Customs Bureau national contest, including a term as Chief Judge from 2016 to 2017.
              </p>
            </div>
          </div>
        </div>
      </div>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>Macroeconomics and systems analysis</h4>
          <p class="muted">
            Phil&rsquo;s analytical work also extends to macroeconomic theory and forecasting.
            From 2011 to 2015, he taught macroeconomics at Gakushuin University and developed &ldquo;The Shape of the Future,&rdquo;
            a course focused on futurology and systemic forecasting.
          </p>
        </div>
        <div class="note-panel">
          <h4>Academic consulting and technical communication</h4>
          <p class="muted">
            Through Stilwell Consulting, which he has operated since 2003, Phil has helped professors and graduate researchers refine technical writing into journal-level English
            and prepare papers and presentations for elite academic settings. From 2015 to 2017, he also taught technical writing at The University of Tokyo,
            focusing on style guides, composition, and publication practice for international graduate students.
          </p>
        </div>
      </div>
    </section>

    <section class="detail-section section-block">
      <p class="eyebrow">Why This Site Exists</p>
      <h2 class="detail-title about-origin-title">A classroom-built resource for teaching better reasoning.</h2>
      <p class="detail-deck">
        The inspiration for LogFall came from the many critical thinking classes Phil taught over the years.
        The project began as a practical teaching tool and gradually became a broader reference for students, instructors, debaters, and careful readers.
      </p>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>Classroom origin</h4>
          <p class="muted">
            In class, students often needed more than a list of fallacy names. They needed help distinguishing similar mistakes,
            seeing exactly where an argument slipped, and learning how to repair a weak argument instead of merely labeling it.
            LogFall was shaped around those repeated classroom needs.
          </p>
        </div>
        <div class="note-panel">
          <h4>Joint effort with students</h4>
          <p class="muted">
            This collection of logical fallacies was a joint effort between Phil and his students.
            Class discussions, example-hunting, objections, revisions, and moments of confusion all helped sharpen the entries,
            making the site less like a static glossary and more like a teaching resource tested against real classroom use.
          </p>
        </div>
      </div>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>What each page is for</h4>
          <p class="muted">
            Each fallacy page combines a definition, a concrete example, explanatory notes, case studies, related fallacies, companion imagery,
            and a guided practice tool. The goal is not just to name a mistake, but to help readers see the structure of the mistake clearly enough to avoid repeating it.
          </p>
        </div>
        <div class="note-panel">
          <h4>How the categories help</h4>
          <p class="muted">
            Categories group fallacies by the main kind of reasoning failure involved, so readers can compare similar mistakes instead of memorizing isolated labels.
            That structure makes the site more useful for teaching, self-study, and side-by-side diagnosis.
          </p>
        </div>
      </div>
    </section>

    <section class="detail-section section-block">
      <p class="eyebrow">How To Use This Knowledge</p>
      <h2 class="detail-title about-origin-title">Fallacy study works best when it is aimed inward first.</h2>
      <p class="detail-deck">
        The best reason to learn logical fallacies is not to collect clever ways of defeating other people in argument.
        It is to become better at noticing one&rsquo;s own reasoning shortcuts, rhetorical temptations, and avoidable mistakes.
      </p>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>Use it for self-correction</h4>
          <p class="muted">
            Fallacy knowledge becomes shallow when it is used mainly as a weapon against opponents. Used well, it is an inward-facing discipline:
            a way of slowing down, checking one&rsquo;s own inferences, and asking whether one&rsquo;s preferred conclusion has been reached too quickly,
            too emotionally, or on too little evidence.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not treat labels as knockouts</h4>
          <p class="muted">
            Naming a fallacy does not automatically settle a dispute. A weak argument can defend a true conclusion, a strong argument can be stated badly,
            and real reasoning often needs patient reconstruction before it can be judged fairly. LogFall is meant to support clearer thinking,
            not point-scoring.
          </p>
        </div>
      </div>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>One part of a rational mind</h4>
          <p class="muted">
            Knowledge of logical fallacies is only one small part of a well-rounded rational mind. Good judgment also depends on familiarity with
            cognitive biases, probability theory, statistical reasoning, and the difference between stronger and weaker forms of evidence.
          </p>
        </div>
        <div class="note-panel">
          <h4>Build the wider toolkit</h4>
          <p class="muted">
            Readers who want to reason well should also understand how deductive and inductive arguments work, how confidence should track evidence,
            and how sampling, uncertainty, and causal inference can distort judgment. Fallacy study matters most when it is integrated into that larger
            discipline of rational inquiry.
          </p>
        </div>
      </div>
    </section>
  `;

  return pageShell({
    title: "About Phil Stilwell and LogFall | Logical Fallacies Teaching Resource",
    description: aboutSeoDescription(),
    prefix: "../",
    currentSection: "about",
    canonicalPath: "about/",
    keywords: [
      "Phil Stilwell",
      "about LogFall",
      "critical thinking teacher",
      "logical fallacies instructor",
      "university critical thinking",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "About", path: "about/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "AboutPage",
        name: "About LogFall",
        url: absoluteUrl("about/"),
        description: aboutSeoDescription(),
        mainEntity: {
          "@type": "Person",
          name: "Phil Stilwell",
          url: absoluteUrl("about/"),
          sameAs: ["https://credencing.com"],
          alumniOf: "The University of Kansas",
          knowsAbout: [
            "philosophy",
            "epistemology",
            "critical thinking",
            "macroeconomics",
            "technical writing",
            "logical fallacies",
          ],
        },
      },
      learningResourceSchema({
        name: "About LogFall",
        path: "about/",
        description: aboutSeoDescription(),
        about: ["Phil Stilwell", "critical thinking instruction", "logical fallacies"],
        teaches: ["how to use logical fallacies well", "self-correction in reasoning"],
        learningResourceType: ["About page", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["Phil Stilwell", "about LogFall", "critical thinking teacher"],
      }),
    ],
    content,
  });
}

function renderPromptCard(promptDefinition) {
  const promptId = `prompt-${promptDefinition.slug}`;
  return `<article class="detail-section prompt-card" id="${escapeHtml(promptDefinition.slug)}">
    <p class="eyebrow">AI prompt</p>
    <h3 class="section-title">${escapeHtml(promptDefinition.title)}</h3>
    <p class="section-copy">${escapeHtml(promptDefinition.intro)}</p>
    <div class="note-panel prompt-requirement">
      <p class="prompt-requirement-label">What else is needed</p>
      <p class="muted">${escapeHtml(promptDefinition.requirements)}</p>
    </div>
    <div class="prompt-toolbar">
      <button class="button button-secondary prompt-copy-button" type="button" data-copy-button="${escapeHtml(promptId)}">Copy prompt</button>
    </div>
    <textarea
      class="prompt-textarea"
      id="${escapeHtml(promptId)}"
      readonly
      spellcheck="false"
    >${escapeHtml(promptDefinition.prompt)}</textarea>
  </article>`;
}

function buildPromptsPage() {
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Fun AI Prompts</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Fun AI Prompts</p>
      <h2 class="detail-title">Copy-ready prompts for fallacy hunting, comparison, and analysis.</h2>
      <p class="detail-deck">
        These prompts are meant to make LogFall more usable with AI tools. They work best when the model is asked to quote source material,
        justify every fallacy label carefully, and link the analysis back to the relevant LogFall page instead of relying on loose impressions.
      </p>
      <div class="prompt-directory">
        ${funPromptDefinitions
          .map(
            (promptDefinition) =>
              `<a class="path-link-chip" href="#${escapeHtml(promptDefinition.slug)}">${escapeHtml(promptDefinition.title)}</a>`,
          )
          .join("")}
      </div>
    </section>

    <section class="section-block">
      <div class="two-column prompt-grid">
        ${funPromptDefinitions.map((promptDefinition) => renderPromptCard(promptDefinition)).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: "AI Prompts for Finding Logical Fallacies in Text and Media | LogFall",
    description: promptsSeoDescription(),
    prefix: "../",
    currentSection: "prompts",
    canonicalPath: "prompts/",
    keywords: [
      "AI prompts for logical fallacies",
      "logical fallacy prompts",
      "argument analysis prompts",
      "critical thinking prompts",
      "fallacy detection prompts",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Fun AI Prompts", path: "prompts/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Fun AI Prompts",
        url: absoluteUrl("prompts/"),
        description: promptsSeoDescription(),
        publisher: publisherSchema(),
        hasPart: funPromptDefinitions.map((promptDefinition) => ({
          "@type": "CreativeWork",
          name: promptDefinition.title,
          description: promptDefinition.intro,
        })),
      },
      learningResourceSchema({
        name: "AI Prompts for Finding Logical Fallacies in Text and Media",
        path: "prompts/",
        description: promptsSeoDescription(),
        about: ["logical fallacies", "AI prompting", "argument analysis"],
        teaches: ["logical fallacy identification", "comparison of arguments", "argument mapping"],
        learningResourceType: ["Prompt library", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["AI prompts for logical fallacies", "argument analysis prompts", "fallacy detection prompts"],
      }),
    ],
    content,
  });
}

function buildAssessmentIndexPage(records, categories) {
  const assessmentBank = buildAssessmentBank(records);
  const overviewCategories = categories
    .slice()
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
    .slice(0, 4)
    .map((category) => category.name)
    .join(", ");

  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Check Yourself</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Check Yourself</p>
      <h2 class="detail-title">A mixed 10-question assessment that does not reveal the answer in the title.</h2>
      <p class="detail-deck">
        Each set draws from the full LogFall library and mixes fallacies so visitors have to identify the reasoning mistake from the claim itself.
        Load another set anytime for a new semi-random group of examples.
      </p>
    </section>

    <section class="panel search-panel assessment-runner-panel" data-assessment-shell data-assessment-size="10">
      <div class="section-header">
        <div>
          <h3 class="section-title">Mixed assessment runner</h3>
          <p class="section-copy">The assessment uses clear example claims, multiple-choice identification, and immediate answer review with links back to the relevant fallacy pages.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>How it works</h4>
          <p class="muted">Each run gives you 10 example claims. Choose the best fallacy label for each one, then grade the whole set at once to see your score and review the correct answers.</p>
        </div>
        <div class="note-panel">
          <h4>What is in the mix</h4>
          <p class="muted">Sets are drawn from all ${records.length} fallacies in LogFall, with especially common categories such as ${escapeHtml(overviewCategories)} appearing often enough to keep the assessment realistic.</p>
        </div>
      </div>
      <div class="note-panel assessment-banner" data-assessment-banner>
        <h4>Loading your set</h4>
        <p class="muted">The assessment is preparing a mixed group of fallacy questions now.</p>
      </div>
      <div class="assessment-toolbar">
        <button class="button button-primary button-compact" type="button" data-assessment-new>Load another set</button>
        <a class="button button-secondary button-compact" href="../fallacies/">Study the full reference</a>
      </div>
      <div class="assessment-items" data-assessment-items></div>
      <div class="assessment-actions">
        <button class="button button-primary" type="button" data-assessment-grade>Grade this assessment</button>
      </div>
      <div class="detail-section assessment-results hidden" data-assessment-results role="status" aria-live="polite"></div>
      <script id="assessment-bank" type="application/json">${safeJsonForScript(assessmentBank)}</script>
      <noscript>
        <div class="note-panel search-empty">
          <h4>JavaScript is required for this assessment</h4>
          <p class="muted">This page builds a mixed quiz set in the browser. If scripting is disabled, you can still study the full reference in <a class="text-link" href="../fallacies/">All Fallacies</a>.</p>
        </div>
      </noscript>
    </section>
  `;

  return pageShell({
    title: "Check Yourself: 10-Question Logical Fallacy Assessment | LogFall",
    description: assessmentIndexSeoDescription(records.length),
    prefix: "../",
    currentSection: "check-yourself",
    canonicalPath: "check-yourself/",
    keywords: [
      "logical fallacy assessment",
      "mixed logical fallacy quiz",
      "critical thinking assessment",
      "10 question fallacy quiz",
      `${records.length} fallacy examples`,
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Check Yourself", path: "check-yourself/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Check Yourself",
        url: absoluteUrl("check-yourself/"),
        description: assessmentIndexSeoDescription(records.length),
        publisher: publisherSchema(),
        about: ["logical fallacies", "critical thinking", "assessment"],
      },
      learningResourceSchema({
        name: "Check Yourself: 10-Question Logical Fallacy Assessment",
        path: "check-yourself/",
        description: assessmentIndexSeoDescription(records.length),
        about: ["logical fallacies", "critical thinking", "argument diagnosis"],
        teaches: ["logical fallacy identification", "comparison of nearby fallacies"],
        learningResourceType: ["Assessment", "Quiz"],
        educationalUse: ["assessment", "teaching", "self-study"],
        keywords: ["logical fallacy assessment", "mixed logical fallacy quiz", "10 question fallacy quiz", "critical thinking assessment"],
      }),
    ],
    content,
  });
}

function buildAssessmentPage(record, records) {
  const pedagogy = pedagogyForRecord(record);
  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Check Yourself</a><span>/</span><strong>${escapeHtml(record.name)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Check Yourself</p>
      <h2 class="detail-title">${escapeHtml(record.name)} assessment</h2>
      <p class="detail-deck">
        Work from one example claim, identify the fallacy, explain the exact reasoning slip, and then repair the argument without saying more than the evidence or reasoning has earned.
      </p>
      <div class="meta-grid section-block">
        <div class="note-panel">
          <h4>Fallacy page</h4>
          <p class="muted"><a class="text-link" href="../../fallacies/${escapeHtml(record.slug)}/">Open the full ${escapeHtml(record.name)} reference entry</a>.</p>
        </div>
        <div class="note-panel">
          <h4>Categories</h4>
          <p class="muted">${escapeHtml(record.categories.join(", "))}</p>
        </div>
        <div class="note-panel">
          <h4>Classroom fit</h4>
          <p class="muted">${escapeHtml(pedagogy.classroomLevel)} · ${escapeHtml(pedagogy.difficulty)}</p>
        </div>
      </div>
    </section>

    <section class="detail-section assessment-example-panel">
      <p class="eyebrow">Example claim to diagnose</p>
      <p class="assessment-example-text">${escapeHtml(record.example)}</p>
      <p class="muted assessment-example-note">Use this exact claim throughout the assessment below. The label, explanation, and repair should all stay tied to this wording.</p>
    </section>

    ${renderQuizAndRepairSection(record, records, { includeHeader: false })}
  `;

  return pageShell({
    title: assessmentSeoTitle(record),
    description: assessmentSeoDescription(record),
    prefix: "../../",
    currentSection: "check-yourself",
    canonicalPath: `check-yourself/${record.slug}/`,
    keywords: [
      `${record.name} quiz`,
      `${record.name} assessment`,
      `${seoFallacyKeyword(record)} quiz`,
      ...record.aliases.slice(0, 3),
      "logical fallacy quiz",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Check Yourself", path: "check-yourself/" },
        { name: record.name, path: `check-yourself/${record.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "WebPage",
        name: assessmentSeoTitle(record),
        url: absoluteUrl(`check-yourself/${record.slug}/`),
        description: assessmentSeoDescription(record),
        isPartOf: {
          "@type": "WebSite",
          name: "LogFall",
          url: absoluteUrl(""),
        },
      },
      learningResourceSchema({
        name: `${record.name} assessment`,
        path: `check-yourself/${record.slug}/`,
        description: assessmentSeoDescription(record),
        about: [record.name, ...record.categories],
        teaches: [record.name, "argument repair", "fallacy identification"],
        learningResourceType: ["Quiz", "Assessment"],
        educationalUse: ["assessment", "teaching", "self-study"],
        keywords: [`${record.name} quiz`, `${record.name} assessment`, "logical fallacy quiz"],
      }),
    ],
    content,
  });
}

function buildAllFallaciesPage(records, categories) {
  const difficultyOptions = ["Foundational", "Intermediate", "Advanced"]
    .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
    .join("");
  const classroomOptions = classroomLevels
    .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
    .join("");
  const options = categories
    .map((category) => `<option value="${escapeHtml(category.name)}">${escapeHtml(category.name)}</option>`)
    .join("");

  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>All Fallacies</strong>
    </div>

    <section class="panel search-panel">
      <div class="section-header">
        <div>
          <h2 class="section-title">All fallacies</h2>
          <p class="section-copy">Search by name, alias, definition, example, or case-study wording, then narrow the list by category.</p>
        </div>
      </div>
      <div class="search-row">
        <input class="search-input" type="search" placeholder="Search fallacies, aliases, or keywords..." data-search-input />
        <select class="search-select" data-category-filter>
          <option value="">All categories</option>
          ${options}
        </select>
        <select class="search-select" data-difficulty-filter>
          <option value="">All difficulty levels</option>
          ${difficultyOptions}
        </select>
        <select class="search-select" data-classroom-filter>
          <option value="">All classroom levels</option>
          ${classroomOptions}
        </select>
        <button class="search-reset" type="button" data-search-reset>Clear</button>
      </div>
      <div class="search-meta" data-search-count role="status" aria-live="polite">${records.length} of ${records.length} fallacies shown</div>
      <div class="note-panel search-empty hidden" data-search-empty>
        <h4>No matches yet</h4>
        <p class="muted">Try a broader keyword, clear the current filters, or browse by category to compare nearby mistakes.</p>
      </div>
    </section>

    <section class="section-block">
      <div class="fallacy-grid" data-fallacy-grid>
        ${records.map((record) => renderFallacyCard(record, "../")).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: "All Logical Fallacies: Search Definitions, Examples, and Categories | LogFall",
    description: fallaciesIndexSeoDescription(records.length),
    prefix: "../",
    currentSection: "fallacies",
    canonicalPath: "fallacies/",
    keywords: [
      "logical fallacies list",
      "fallacy index",
      "reasoning errors list",
      "critical thinking glossary",
      `${records.length} logical fallacies`,
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "All Fallacies", path: "fallacies/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "All Logical Fallacies",
        url: absoluteUrl("fallacies/"),
        description: fallaciesIndexSeoDescription(records.length),
        publisher: publisherSchema(),
        mainEntity: {
          "@type": "DefinedTermSet",
          name: "LogFall Fallacies",
          url: absoluteUrl("fallacies/"),
          numberOfItems: records.length,
        },
      },
      learningResourceSchema({
        name: "All Logical Fallacies",
        path: "fallacies/",
        description: fallaciesIndexSeoDescription(records.length),
        about: ["logical fallacies", "critical thinking", "reasoning errors"],
        teaches: ["logical fallacies", "comparison of fallacies", "argument analysis"],
        learningResourceType: ["Glossary", "Reference"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacies list", "fallacy index", "critical thinking glossary"],
      }),
    ],
    content,
  });
}

function buildCategoriesIndexPage(categories) {
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Categories</strong>
    </div>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h2 class="section-title">Taxonomy categories</h2>
          <p class="section-copy">Each category groups fallacies by the main way reasoning goes wrong, not by topic or ideology.</p>
        </div>
      </div>
      <div class="category-grid">
        ${categories
          .map(
            (category) => `<article class="category-card">
              <h3><a href="${category.slug}/">${escapeHtml(category.name)}</a></h3>
              <p class="card-copy">${escapeHtml(category.description)}</p>
              <div class="pill-row"><span class="pill pill-${escapeHtml(category.name)}">${category.count} entries</span></div>
            </article>`,
          )
          .join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: "Logical Fallacy Categories and Taxonomy | LogFall",
    description: "Browse the LogFall taxonomy of logical fallacies by category, from causal and evidential errors to formal and emotional ones.",
    prefix: "../",
    currentSection: "categories",
    canonicalPath: "categories/",
    keywords: [
      "logical fallacy categories",
      "fallacy taxonomy",
      "types of logical fallacies",
      "reasoning error categories",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Categories", path: "categories/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Logical Fallacy Categories",
        url: absoluteUrl("categories/"),
        description:
          "Browse the LogFall taxonomy of logical fallacies by category, from causal and evidential errors to formal and emotional ones.",
        publisher: publisherSchema(),
        hasPart: categories.map((category) => ({
          "@type": "CreativeWork",
          name: category.name,
          url: absoluteUrl(`categories/${category.slug}/`),
          description: category.description,
        })),
      },
      learningResourceSchema({
        name: "Logical Fallacy Categories and Taxonomy",
        path: "categories/",
        description: "Browse the LogFall taxonomy of logical fallacies by category, from causal and evidential errors to formal and emotional ones.",
        about: ["logical fallacies", "taxonomy", "critical thinking"],
        teaches: ["types of logical fallacies", "comparison of reasoning errors"],
        learningResourceType: ["Taxonomy", "Reference"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacy categories", "fallacy taxonomy", "types of logical fallacies"],
      }),
    ],
    content,
  });
}

function buildCategoryPage(category, records) {
  const members = records.filter((record) => record.categories.includes(category.name));
  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Categories</a><span>/</span><strong>${escapeHtml(category.name)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Category</p>
      <h2 class="detail-title">${escapeHtml(category.name)}</h2>
      <p class="detail-deck">${escapeHtml(category.description)}</p>
      <div class="meta-grid section-block">
        <div class="note-panel">
          <h4>Entries</h4>
          <p class="muted">${category.count} fallacies in this category.</p>
        </div>
        <div class="note-panel">
          <h4>Diagnostic prompt</h4>
          <p class="muted">${escapeHtml(diagnosticPrompts[category.name])}</p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="fallacy-grid">
        ${members.map((record) => renderFallacyCard(record, "../../")).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: `${category.name} Logical Fallacies: Definitions and Examples | LogFall`,
    description: categorySeoDescription(category),
    prefix: "../../",
    currentSection: "categories",
    canonicalPath: `categories/${category.slug}/`,
    keywords: [
      `${category.name} logical fallacies`,
      `${category.name.toLowerCase()} fallacies`,
      `${category.name.toLowerCase()} reasoning errors`,
      "critical thinking",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Categories", path: "categories/" },
        { name: category.name, path: `categories/${category.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: `${category.name} Logical Fallacies`,
        url: absoluteUrl(`categories/${category.slug}/`),
        description: categorySeoDescription(category),
        publisher: publisherSchema(),
        mainEntity: {
          "@type": "ItemList",
          numberOfItems: members.length,
          itemListElement: members.map((record, index) => ({
            "@type": "ListItem",
            position: index + 1,
            name: record.name,
            url: absoluteUrl(`fallacies/${record.slug}/`),
          })),
        },
      },
      learningResourceSchema({
        name: `${category.name} Logical Fallacies`,
        path: `categories/${category.slug}/`,
        description: categorySeoDescription(category),
        about: [category.name, "logical fallacies", "critical thinking"],
        teaches: [`${category.name} logical fallacies`, "argument analysis"],
        learningResourceType: ["Category page", "Reference"],
        educationalUse: ["teaching", "self-study"],
        keywords: [`${category.name} logical fallacies`, `${category.name.toLowerCase()} fallacies`, "critical thinking"],
      }),
    ],
    content,
  });
}

function buildDetailPage(record, records, categoryProfiles, posterAssets) {
  const related = relatedFallacies(record, records);
  const prompts = record.categories.map((category) => diagnosticPrompts[category]).filter(Boolean);
  const hasPosterIllustration = Boolean(resolvePosterAssetForRecord(record, posterAssets));
  const referenceMeta = renderReferenceMeta(record, prompts);
  const profileReferenceMarkup = hasPosterIllustration
    ? `<div class="profile-reference">
          <p class="eyebrow">Reference</p>
          ${referenceMeta}
        </div>`
    : "";

  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">All Fallacies</a><span>/</span><strong>${escapeHtml(record.name)}</strong>
    </div>

    <section class="detail-hero${hasPosterIllustration ? " detail-hero-with-illustration" : ""}">
      <article class="detail-section">
        <p class="eyebrow">Fallacy profile</p>
        <h2 class="detail-title">${escapeHtml(record.name)}</h2>
        <p class="detail-deck">${escapeHtml(record.definition)}</p>
        ${renderPills(record.categories)}
        <div class="detail-grid">
          <div class="note-panel">
            <p class="detail-card-label">Definition</p>
            <p class="detail-card-value">${escapeHtml(record.definition)}</p>
          </div>
          <div class="note-panel">
            <p class="detail-card-label">Illustrative example</p>
            <p class="detail-card-value">${escapeHtml(record.example)}</p>
          </div>
        </div>
        ${renderRhetoricGaugeSection(record)}${profileReferenceMarkup}
      </article>

      ${
        hasPosterIllustration
          ? renderPosterIllustration(record, "../../", posterAssets)
          : `<aside class="detail-section">
        <p class="eyebrow">Reference</p>
        <div class="meta-grid">
          <div class="note-panel">
            <h4>Family</h4>
            <p class="muted">${escapeHtml(record.family || "Unspecified")}</p>
          </div>
          ${
            record.aliases.length
              ? `<div class="note-panel">
            <h4>Aliases</h4>
            <p class="muted">${escapeHtml(record.aliases.join(", "))}</p>
          </div>`
              : ""
          }
          <div class="note-panel">
            <h4>Quick check</h4>
            <p class="muted">${escapeHtml(prompts[0] || "Ask what evidence or reasoning step is doing too much work.")}</p>
          </div>
        </div>
      </aside>`
      }
    </section>

    ${
      record.notes
        ? `<section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Why it misleads</h3>
          <p class="section-copy">A fuller explanation of how the fallacy works and why it can look persuasive.</p>
        </div>
      </div>
      <article class="detail-section">
        <p>${escapeHtml(record.notes)}</p>
      </article>
    </section>`
        : ""
    }

    ${renderCaveatSection(record, records)}

    ${renderConfusionSection(record, records, "../../")}

    ${renderArgumentMapSection(record)}

    ${renderRationalityLab(record, categoryProfiles)}

    ${renderAssessmentTeaserSection(record)}

    ${
      prompts.length
        ? `<section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Questions to ask</h3>
          <p class="section-copy">Use these category-based prompts to audit similar arguments.</p>
        </div>
      </div>
      <div class="two-column">
        ${prompts
          .map(
            (prompt, index) => `<div class="note-panel">
              <h4>Prompt ${index + 1}</h4>
              <p class="muted">${escapeHtml(prompt)}</p>
            </div>`,
          )
          .join("")}
      </div>
    </section>`
        : ""
    }

    ${
      record.caseStudies.length
        ? `<section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Case studies</h3>
          <p class="section-copy">Each case study explains why the example fits the fallacy and links back to its source whenever source information is available.</p>
        </div>
      </div>
      <div class="case-list">
        ${record.caseStudies.map((item) => renderCaseStudy(item)).join("")}
      </div>
    </section>`
        : ""
    }

    ${
      related.length
        ? `<section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Related fallacies</h3>
          <p class="section-copy">Nearby entries chosen by shared categories and family resemblance.</p>
        </div>
      </div>
      <div class="fallacy-grid related-grid">
        ${related.map((candidate) => renderFallacyCard(candidate, "../../")).join("")}
      </div>
    </section>`
        : ""
    }
  `;

  return pageShell({
    title: fallacySeoTitle(record),
    description: fallacySeoDescription(record),
    prefix: "../../",
    currentSection: "fallacies",
    canonicalPath: `fallacies/${record.slug}/`,
    ogType: "article",
    extraHeadHtml: cloudflareWebAnalyticsTag,
    keywords: [
      record.name,
      seoFallacyKeyword(record),
      "logical fallacy",
      ...record.aliases.slice(0, 3),
      ...record.categories.map((category) => `${category.toLowerCase()} logical fallacies`),
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "All Fallacies", path: "fallacies/" },
        { name: record.name, path: `fallacies/${record.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "WebPage",
        name: fallacySeoTitle(record),
        url: absoluteUrl(`fallacies/${record.slug}/`),
        description: fallacySeoDescription(record),
        isPartOf: {
          "@type": "WebSite",
          name: "LogFall",
          url: absoluteUrl(""),
        },
        mainEntity: {
          "@type": "DefinedTerm",
          name: record.name,
          description: record.definition,
          url: absoluteUrl(`fallacies/${record.slug}/`),
          alternateName: record.aliases.length ? record.aliases : undefined,
          inDefinedTermSet: absoluteUrl("fallacies/"),
          termCode: record.slug,
        },
      },
      learningResourceSchema({
        name: seoFallacyName(record),
        path: `fallacies/${record.slug}/`,
        description: fallacySeoDescription(record),
        about: [record.name, ...record.categories],
        teaches: [record.name, ...record.aliases],
        learningResourceType: ["Reference", "Practice tool"],
        educationalUse: ["teaching", "self-study"],
        keywords: [
          record.name,
          seoFallacyKeyword(record),
          ...record.aliases.slice(0, 3),
          ...record.categories.map((category) => `${category.toLowerCase()} logical fallacies`),
        ],
      }),
    ],
    content,
  });
}

function build404Page() {
  const content = `
    <section class="detail-section">
      <p class="eyebrow">Not found</p>
      <h2 class="detail-title">This page doesn't exist yet.</h2>
      <p class="detail-deck">Try the full index or browse by category to find the fallacy you were looking for.</p>
      <div class="hero-actions">
        <a class="button button-primary" href="./">Home</a>
        <a class="button button-secondary" href="./fallacies/">All Fallacies</a>
      </div>
    </section>
  `;

  return pageShell({
    title: "404 | LogFall",
    description: "Fallback page for the LogFall static site.",
    prefix: "",
    currentSection: "",
    canonicalPath: "404.html",
    robots: "noindex,follow",
    content,
  });
}

function columnLetter(index) {
  let value = "";
  let current = index;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    current = Math.floor((current - 1) / 26);
  }
  return value;
}

async function buildWorkbook(records, categories, categoryProfiles) {
  const workbook = Workbook.create();

  const overview = workbook.worksheets.add("Overview");
  const overviewRows = [
    ["LogFall Workbook", "Reference data used to generate the LogFall site."],
    ["Record count", records.length],
    ["Category count", categories.length],
    ["Case studies", records.reduce((sum, record) => sum + record.caseStudies.length, 0)],
    ["Gauge scale", "The three gauge columns use 0-100 editorial teaching estimates rather than measured statistics."],
    ["Workbook use", "Use the Fallacies sheet for entry-by-entry editing and the Categories sheet for counts and category notes."],
    ["Copyright", copyrightNotice],
  ];
  overview.getRange(`A1:B${overviewRows.length}`).values = overviewRows;
  overview.getRange("A:A").format.columnWidthPx = 220;
  overview.getRange("B:B").format.columnWidthPx = 620;
  overview.getRange("B:B").format.wrapText = true;

  const fallaciesSheet = workbook.worksheets.add("Fallacies");
  const headers = [
    "Name",
    "Slug",
    "Primary Category",
    "Additional Category 1",
    "Additional Category 2",
    "Original Number",
    "Family",
    "Sub-Category",
    "Sub-Sub-Category",
    "Aliases",
    "Difficulty",
    "Classroom Level",
    "Teaching Domain",
    "Teaching Paths",
    "Common in Rhetoric (0-100)",
    "Easy to Spot (0-100)",
    "Easy to Innocently Commit (0-100)",
    "Often Confused With",
    "Definition",
    "Example",
    "Notes",
    "Caveat",
    "Why This Mistake Matters",
    "Main Reasoning Problem",
    "What to Watch For",
    "Warning Signs",
    "Suggested Tool",
    "What the Reader Does",
    "What the Tool Shows",
    "How to Fix It",
    "Case Study 1",
    "Case Study 2",
    "Case Study 3",
    "Case Study 4",
    "Case Study 5",
    "Editorial Status",
  ];
  const rows = records.map((record) => {
    const pedagogy = pedagogyForRecord(record);
    const gauges = rhetoricGaugesForRecord(record);
    const confusions = confusionCandidates(record, records, 2)
      .map((item) => item.candidate.name)
      .join(", ");
    return [
      record.name,
      record.slug,
      record.categories[0] || "",
      record.categories[1] || "",
      record.categories[2] || "",
      record.originalNumber,
      record.family,
      record.subCategory,
      record.subSubCategory,
      record.aliases.join(", "),
      pedagogy.difficulty,
      pedagogy.classroomLevel,
      pedagogy.domainTag,
      pedagogy.teachingPaths.map((item) => item.title).join(", "),
      gauges.common.value,
      gauges.spot.value,
      gauges.innocent.value,
      confusions,
      record.definition,
      record.example,
      record.notes,
      caveatTextForWorkbook(record, records),
      record.rationalityDanger || "",
      record.mainReasoningProblem || "",
      record.dynamicsToNotice || "",
      record.warningSigns || "",
      record.interactiveMechanic || "",
      record.userAction || "",
      record.feedbackLogic || "",
      record.repairPrompts || "",
      formatCaseStudyCell(record.caseStudies[0]),
      formatCaseStudyCell(record.caseStudies[1]),
      formatCaseStudyCell(record.caseStudies[2]),
      formatCaseStudyCell(record.caseStudies[3]),
      formatCaseStudyCell(record.caseStudies[4]),
      record.editorialStatus,
    ];
  });
  const fallaciesMatrix = [headers, ...rows];
  fallaciesSheet.getRange(`A1:${columnLetter(headers.length)}${fallaciesMatrix.length}`).values =
    fallaciesMatrix;
  fallaciesSheet.freezePanes.freezeRows(1);
  fallaciesSheet.freezePanes.freezeColumns(4);
  fallaciesSheet.getRange("A:A").format.columnWidthPx = 260;
  fallaciesSheet.getRange("B:B").format.columnWidthPx = 220;
  fallaciesSheet.getRange("C:E").format.columnWidthPx = 180;
  fallaciesSheet.getRange("F:F").format.columnWidthPx = 120;
  fallaciesSheet.getRange("G:I").format.columnWidthPx = 180;
  fallaciesSheet.getRange("J:N").format.columnWidthPx = 220;
  fallaciesSheet.getRange("O:Q").format.columnWidthPx = 150;
  fallaciesSheet.getRange("R:R").format.columnWidthPx = 220;
  fallaciesSheet.getRange("S:AD").format.columnWidthPx = 420;
  fallaciesSheet.getRange("AE:AI").format.columnWidthPx = 360;
  fallaciesSheet.getRange("AJ:AJ").format.columnWidthPx = 180;
  fallaciesSheet.getRange(`A1:${columnLetter(headers.length)}1`).format.wrapText = true;
  fallaciesSheet.getRange("J:AJ").format.wrapText = true;

  const categorySheet = workbook.worksheets.add("Categories");
  const categoryRows = [
    ["Category", "Count", "Description", "Main Reasoning Problem", "Why This Kind Matters", "Suggested Tool", "What the Reader Does", "What the Tool Shows"],
    ...categories.map((category) => {
      const profile = categoryProfiles[category.name] || {};
      return [
        category.name,
        category.count,
        category.description,
        profile.distortion || "",
        profile.danger || "",
        profile.mechanic || "",
        profile.user_action || "",
        profile.feedback || "",
      ];
    }),
  ];
  categorySheet.getRange(`A1:H${categoryRows.length}`).values = categoryRows;
  categorySheet.freezePanes.freezeRows(1);
  categorySheet.getRange("A:A").format.columnWidthPx = 180;
  categorySheet.getRange("B:B").format.columnWidthPx = 100;
  categorySheet.getRange("C:H").format.columnWidthPx = 360;
  categorySheet.getRange("C:H").format.wrapText = true;

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(workbookOutPath);
}

async function ensureCleanDist() {
  await fs.mkdir(distRoot, { recursive: true });
  await fs.mkdir(dataOutDir, { recursive: true });
}

async function pruneGeneratedDirectories(parentDir, keepNames) {
  let entries = [];
  try {
    entries = await fs.readdir(parentDir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (keepNames.has(entry.name)) continue;
    await fs.rm(path.join(parentDir, entry.name), { recursive: true, force: true });
  }
}

async function writeText(relativePath, contents) {
  const fullPath = path.join(distRoot, relativePath);
  await fs.mkdir(path.dirname(fullPath), { recursive: true });
  await fs.writeFile(fullPath, contents, "utf8");
}

async function main() {
  await ensureCleanDist();

  const payload = JSON.parse(await fs.readFile(dataPath, "utf8"));
  const records = payload.records.map((record) => ({
    ...record,
    categories: normalizeRecordCategories(record),
  }));
  const categoryProfiles = payload.categoryProfiles || {};
  const categories = payload.categories.map((category) => ({
    ...category,
    description: categoryDescriptions[category.name] || "A reasoning category in the LogFall taxonomy.",
  }));
  const posterAssets = new Set(
    (await fs.readdir(path.join(distRoot, "assets")).catch(() => []))
      .filter((name) => /^fallacy-.*-poster\.(webp|png|jpe?g)$/i.test(name)),
  );

  await pruneGeneratedDirectories(path.join(distRoot, "fallacies"), new Set(records.map((record) => record.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "categories"), new Set(categories.map((category) => category.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "paths"), new Set(teachingPathDefinitions.map((pathDefinition) => pathDefinition.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "check-yourself"), new Set());

  await fs.copyFile(path.join(siteRoot, "styles.css"), path.join(distRoot, "styles.css"));
  await fs.copyFile(path.join(siteRoot, "app.js"), path.join(distRoot, "app.js"));

  await writeText("index.html", buildHomePage(records, categories));
  await writeText("about/index.html", buildAboutPage());
  await writeText("check-yourself/index.html", buildAssessmentIndexPage(records, categories));
  await writeText("prompts/index.html", buildPromptsPage());
  await writeText("fallacies/index.html", buildAllFallaciesPage(records, categories));
  await writeText("categories/index.html", buildCategoriesIndexPage(categories));
  await writeText("paths/index.html", buildTeachingPathsIndexPage(records));
  await writeText("404.html", build404Page());
  const sitemapEntries = [
    { path: "" },
    { path: "about/" },
    { path: "check-yourself/" },
    { path: "prompts/" },
    { path: "fallacies/" },
    { path: "categories/" },
    { path: "paths/" },
    ...categories.map((category) => ({ path: `categories/${category.slug}/` })),
    ...teachingPathDefinitions.map((pathDefinition) => ({ path: `paths/${pathDefinition.slug}/` })),
    ...records.map((record) => ({ path: `fallacies/${record.slug}/` })),
  ];
  await writeText("sitemap.xml", buildSitemap(sitemapEntries));
  await writeText("robots.txt", `User-agent: *\nAllow: /\nSitemap: ${absoluteUrl("sitemap.xml")}\n`);

  for (const category of categories) {
    await writeText(`categories/${category.slug}/index.html`, buildCategoryPage(category, records));
  }

  for (const pathDefinition of teachingPathDefinitions) {
    await writeText(`paths/${pathDefinition.slug}/index.html`, buildTeachingPathPage(pathDefinition, records));
  }

  for (const record of records) {
    await writeText(
      `fallacies/${record.slug}/index.html`,
      buildDetailPage(record, records, categoryProfiles, posterAssets),
    );
  }

  await buildWorkbook(records, categories, categoryProfiles);

  console.log(
    JSON.stringify(
      {
        distRoot,
        pageCount: 8 + categories.length + teachingPathDefinitions.length + records.length,
        recordCount: records.length,
        workbookOutPath,
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
