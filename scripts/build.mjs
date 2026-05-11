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
const buildDate = new Date().toISOString().split("T")[0];
const copyrightNotice = "Copyright © Phil Stilwell";
const cloudflareWebAnalyticsTag =
  `<!-- Cloudflare Web Analytics --><script defer src='https://static.cloudflareinsights.com/beacon.min.js' data-cf-beacon='{"token": "798a30777a6d424f9c4055a02e7bde91"}'></script><!-- End Cloudflare Web Analytics -->`;

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

function pageShell({
  title,
  description,
  prefix,
  content,
  currentSection = "",
  canonicalPath = "",
  ogType = "website",
  robots = "index,follow",
  extraHeadHtml = "",
}) {
  const homeHref = prefix || "./";
  const navItems = [
    { href: homeHref, label: "Home", key: "home" },
    { href: `${prefix}fallacies/`, label: "All Fallacies", key: "fallacies" },
    { href: `${prefix}categories/`, label: "Categories", key: "categories" },
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

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(title)}</title>
    <meta name="description" content="${escapeHtml(description)}" />
    <meta name="robots" content="${escapeHtml(robots)}" />
    <link rel="canonical" href="${escapeHtml(canonicalUrl)}" />
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
    <meta property="og:image:alt" content="LogFall logo" />
    <meta name="twitter:card" content="summary" />
    <meta name="twitter:title" content="${escapeHtml(title)}" />
    <meta name="twitter:description" content="${escapeHtml(description)}" />
    <meta name="twitter:image" content="${escapeHtml(socialImageUrl)}" />
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

function renderQuizAndRepairSection(record, records) {
  const quiz = quizConfigForRecord(record, records);
  const repairChecklist = repairChecklistForRecord(record);
  const repairModel = repairModelForRecord(record);

  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">Check yourself</h3>
        <p class="section-copy">Use this short quiz to identify the fallacy, explain it, and then practice repairing the claim.</p>
      </div>
    </div>
    <div class="two-column quiz-grid">
      <article
        class="detail-section quiz-card"
        data-quiz-widget
        data-quiz-answer="${escapeHtml(quiz.answer)}"
        data-quiz-keywords="${escapeHtml(JSON.stringify(quiz.keywords))}"
        data-quiz-model="${escapeHtml(quiz.model)}"
      >
        <h4>1. Identify the fallacy</h4>
        <p class="muted">Choose the best label for the example on this page.</p>
        <p class="quiz-example"><strong>Example:</strong> ${escapeHtml(record.example)}</p>
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
    title: "Teaching Paths | LogFall",
    description: "Curated routes through LogFall for classrooms, review sessions, and first-time readers.",
    prefix: "../",
    currentSection: "",
    canonicalPath: "paths/",
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
    title: `${pathDefinition.title} | LogFall`,
    description: pathDefinition.description,
    prefix: "../../",
    currentSection: "",
    canonicalPath: `paths/${pathDefinition.slug}/`,
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
    title: "LogFall | Logical Fallacies",
    description:
      "A practical logical fallacies reference with category browsing, clear explanations, case studies, and teaching tools.",
    prefix: "",
    currentSection: "home",
    canonicalPath: "",
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
      <h2 class="detail-title">A teaching-focused reference for reasoning mistakes.</h2>
      <p class="detail-deck">
        LogFall is designed to help readers recognize common reasoning errors, distinguish near neighbors,
        and practice better habits of interpretation, comparison, and repair.
      </p>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>What each page does</h4>
          <p class="muted">Each fallacy page combines a definition, a concrete example, explanatory notes, case studies, related fallacies, and a guided practice tool.</p>
        </div>
        <div class="note-panel">
          <h4>How the categories work</h4>
          <p class="muted">Categories sort fallacies by the main kind of reasoning failure involved, so readers can compare similar mistakes instead of memorizing isolated names.</p>
        </div>
      </div>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>Design goal</h4>
          <p class="muted">Make the taxonomy easy to scan, search, compare, and teach while preserving the recognizable red-and-cyan LogFall identity.</p>
        </div>
        <div class="note-panel">
          <h4>Editorial standard</h4>
          <p class="muted">Entries are strongest when they are plainspoken, logically precise, tied to concrete examples, and explicit about how a better line of reasoning would proceed.</p>
        </div>
      </div>
    </section>
  `;

  return pageShell({
    title: "About | LogFall",
    description: "About LogFall and how to use the site to study logical fallacies.",
    prefix: "../",
    currentSection: "about",
    canonicalPath: "about/",
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
    title: "All Fallacies | LogFall",
    description: "Search and browse the full LogFall index of logical fallacies.",
    prefix: "../",
    currentSection: "fallacies",
    canonicalPath: "fallacies/",
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
    title: "Categories | LogFall",
    description: "Browse the LogFall taxonomy by category.",
    prefix: "../",
    currentSection: "categories",
    canonicalPath: "categories/",
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
    title: `${category.name} | LogFall`,
    description: category.description,
    prefix: "../../",
    currentSection: "categories",
    canonicalPath: `categories/${category.slug}/`,
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
        </div>${profileReferenceMarkup}
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

    ${renderConfusionSection(record, records, "../../")}

    ${renderArgumentMapSection(record)}

    ${renderRationalityLab(record, categoryProfiles)}

    ${renderQuizAndRepairSection(record, records)}

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
    title: `${record.name} | LogFall`,
    description: truncate(record.definition, 150),
    prefix: "../../",
    currentSection: "fallacies",
    canonicalPath: `fallacies/${record.slug}/`,
    ogType: "article",
    extraHeadHtml: cloudflareWebAnalyticsTag,
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
    "Often Confused With",
    "Definition",
    "Example",
    "Notes",
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
      confusions,
      record.definition,
      record.example,
      record.notes,
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
  fallaciesSheet.getRange("J:O").format.columnWidthPx = 220;
  fallaciesSheet.getRange("P:Z").format.columnWidthPx = 420;
  fallaciesSheet.getRange("AA:AE").format.columnWidthPx = 360;
  fallaciesSheet.getRange("AF:AF").format.columnWidthPx = 180;
  fallaciesSheet.getRange(`A1:${columnLetter(headers.length)}1`).format.wrapText = true;
  fallaciesSheet.getRange("J:AF").format.wrapText = true;

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

  await fs.copyFile(path.join(siteRoot, "styles.css"), path.join(distRoot, "styles.css"));
  await fs.copyFile(path.join(siteRoot, "app.js"), path.join(distRoot, "app.js"));

  await writeText("index.html", buildHomePage(records, categories));
  await writeText("about/index.html", buildAboutPage());
  await writeText("fallacies/index.html", buildAllFallaciesPage(records, categories));
  await writeText("categories/index.html", buildCategoriesIndexPage(categories));
  await writeText("paths/index.html", buildTeachingPathsIndexPage(records));
  await writeText("404.html", build404Page());
  const sitemapEntries = [
    { path: "" },
    { path: "about/" },
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
        pageCount: 6 + categories.length + teachingPathDefinitions.length + records.length,
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
