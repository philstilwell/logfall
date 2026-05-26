import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const dataPath = path.join(projectRoot, "data", "fallacies.json");
const byteseismicCrossrefsPath = path.join(projectRoot, "data", "byteseismic_crossrefs.json");
const posterCaptionsPath = path.join(projectRoot, "data", "poster_captions.json");
const siteRoot = path.join(projectRoot, "site");
const distRoot = projectRoot;
const dataOutDir = path.join(distRoot, "data");
const workbookOutPath = path.join(distRoot, "logfall-root-edition.xlsx");
const siteUrl = "https://logfall.com/";
const byteseismicSiteUrl = "https://byteseismic.com/";
const socialImagePath = "assets/logo.jpg";
const socialImageType = "image/jpeg";
const socialImageWidth = 124;
const socialImageHeight = 124;
const buildDate = new Intl.DateTimeFormat("en-CA", {
  timeZone: "America/New_York",
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
}).format(new Date());
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
const featuresSectionLabel = "Fallacy Detective";

let posterCaptionOverrides = {};

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

const mapDimensions = [
  {
    slug: "common",
    label: "Common in today's rhetoric",
    shortLabel: "Common",
    lowLabel: "Rare",
    highLabel: "Frequent",
  },
  {
    slug: "spot",
    label: "Easy to spot",
    shortLabel: "Easy to spot",
    lowLabel: "Hidden",
    highLabel: "Obvious",
  },
  {
    slug: "innocent",
    label: "Easy to innocently commit",
    shortLabel: "Easy to commit",
    lowLabel: "Low risk",
    highLabel: "Easy slip",
  },
  {
    slug: "difficulty",
    label: "Difficulty",
    shortLabel: "Difficulty",
    lowLabel: "Foundational",
    highLabel: "Advanced",
  },
];

const mapCategoryPalette = {
  Formal: "#1f5f9d",
  Mathematical: "#53627b",
  Causal: "#d06a4c",
  Linguistic: "#16889e",
  Conceptual: "#846277",
  Evidential: "#e63b34",
  Perceptual: "#7d8fa8",
  Perspectival: "#8a6f3f",
  Epistemic: "#5267a8",
  Tactical: "#111111",
  Emotional: "#c44e36",
};

const familyDescriptions = {
  "Formal/Structural Fallacy":
    "The argument fails because its internal structure does not validly carry the premises to the conclusion.",
  "Evidential/Methodological Fallacy":
    "The mistake lies in how evidence is gathered, weighed, interpreted, or treated as sufficient.",
  "Causal/Explanatory Fallacy":
    "The error concerns what caused what, what explains what, or how a process is supposed to work.",
  "Statistical/Sampling Fallacy":
    "The reasoning misuses rates, probabilities, samples, distributions, or other quantitative expectations.",
  "Linguistic/Definition Fallacy":
    "The problem is driven by wording, ambiguity, definitions, or verbal framing rather than sound reasoning.",
  "Conceptual/Framing Fallacy":
    "The claim is distorted by bad categories, rigid framing, or confused conceptual boundaries.",
  "Comparison/Generalization Fallacy":
    "The argument draws the wrong lesson from a comparison, stereotype, exception, or generalization.",
  "Relevance/Distraction Fallacy":
    "The move shifts attention away from the real issue and substitutes something rhetorically nearby but logically irrelevant.",
  "Persuasive/Appeal Fallacy":
    "The argument leans on emotional, social, or rhetorical force where evidence or reasoning should do the work.",
};

const familyDiagnosticPrompts = {
  "Formal/Structural Fallacy":
    "If the premises were true, would this form still fail to support the conclusion?",
  "Evidential/Methodological Fallacy":
    "What evidence is missing, cherry-picked, stretched, or treated as stronger than it is?",
  "Causal/Explanatory Fallacy":
    "What causal or explanatory link is being assumed rather than actually shown?",
  "Statistical/Sampling Fallacy":
    "What sample, base rate, probability, or distribution is being mishandled here?",
  "Linguistic/Definition Fallacy":
    "What shift in wording, meaning, or definition is doing the hidden work?",
  "Conceptual/Framing Fallacy":
    "What bad category, rigid frame, or confused boundary is distorting the claim?",
  "Comparison/Generalization Fallacy":
    "What comparison, stereotype, or thin slice of experience is being overextended?",
  "Relevance/Distraction Fallacy":
    "What rhetorically nearby move is distracting from the real issue?",
  "Persuasive/Appeal Fallacy":
    "What emotional, social, or rhetorical pressure is standing in for actual support?",
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

const theoryArticleDefinitions = [
  {
    slug: "fallacy-rebuttals-without-fallacy-naming",
    title: "Fallacy Rebuttals Without Fallacy Naming",
    description:
      "A rigorous guide to answering bad reasoning with clarifying analogies before reaching for a technical label.",
    intro:
      "This article explains why some of the clearest rebuttals do not start by naming the fallacy. Instead, they expose the shape of the mistake with a parallel analogy that makes the misstep obvious before any label is introduced.",
  },
  {
    slug: "teaching-logical-fallacies-a-classroom-process-and-curriculum",
    title: "Teaching Logical Fallacies: A Classroom Process and Curriculum",
    description:
      "A practical teaching process and adaptable curriculum for using logical fallacies in a critical thinking class without collapsing the subject into jargon drills or point-scoring.",
    intro:
      "This article gives teachers a repeatable classroom process, a unit sequence, and a wider curriculum frame for teaching logical fallacies as part of a serious critical thinking course.",
  },
  {
    slug: "teaching-logical-fallacies-with-ai-gems-and-prompted-agents",
    title: "Teaching Logical Fallacies with AI Gems and Prompted Agents",
    description:
      "A classroom model for teaching logical fallacies by having students collaboratively design a Gemini Gem or other pre-prompted agent that identifies, scores, and responds to weak reasoning.",
    intro:
      "This article shows how teachers can use collaborative AI-agent design to teach fallacy recognition, comparison, scoring, rebuttal, and repair without surrendering judgment to the model.",
  },
  {
    slug: "how-to-distinguish-fallacies-from-cognitive-biases",
    title: "How to Distinguish Fallacies from Cognitive Biases",
    description:
      "A practical guide to telling the difference between bad arguments in public form and the background mental habits that often help produce them.",
    intro:
      "This article shows why fallacies and cognitive biases overlap, why they are not the same thing, and how to teach the difference without turning either topic into mush.",
  },
  {
    slug: "when-not-to-call-something-a-fallacy",
    title: "When Not to Call Something a Fallacy",
    description:
      "A restraint-focused guide to false positives, overlabeling, and the discipline of withholding a fallacy charge when the fit is weak.",
    intro:
      "This article argues that a good critical thinker needs a working brake pedal as well as an accelerator: sometimes the wisest move is not to reach for a fallacy label.",
  },
  {
    slug: "near-neighbors-how-to-tell-similar-fallacies-apart",
    title: "Near Neighbors: How to Tell Similar Fallacies Apart",
    description:
      "A comparison guide to the fallacies students and readers most often collapse together, with exact splits and memorable examples.",
    intro:
      "This article is built for the moment when a student says, 'Wait, how is that not just a red herring?' That moment is where real understanding begins.",
  },
  {
    slug: "how-to-repair-a-fallacious-argument",
    title: "How to Repair a Fallacious Argument",
    description:
      "A constructive guide to turning a bad argument into a better one by narrowing claims, adding evidence, clarifying terms, and fixing the inference.",
    intro:
      "This article treats fallacy diagnosis as the beginning of the job rather than the end. Spotting the damage matters, but rebuilding the argument matters more.",
  },
  {
    slug: "formal-informal-causal-and-statistical-kinds-of-reasoning-failure",
    title: "Formal, Informal, Causal, and Statistical Kinds of Reasoning Failure",
    description:
      "An accessible taxonomy of different kinds of reasoning breakdowns and why they should not all be taught or diagnosed in the same way.",
    intro:
      "This article explains why some errors are structural, some are evidential, some are causal, and some are statistical, and why lumping them all together makes students dumber than they need to be.",
  },
  {
    slug: "why-true-conclusions-can-still-have-bad-arguments",
    title: "Why True Conclusions Can Still Have Bad Arguments",
    description:
      "A foundational article on why argument quality and conclusion truth come apart, and why that distinction is central to serious critical thinking.",
    intro:
      "This article explains one of the first shocks of logic class: a claim can be right for embarrassingly bad reasons, and that still matters.",
  },
  {
    slug: "how-probability-and-statistics-clarify-logical-fallacies",
    title: "How Probability and Statistics Clarify Logical Fallacies",
    description:
      "A practical article on how base rates, sample size, uncertainty, regression, and causal design make many fallacies easier to diagnose.",
    intro:
      "This article shows that probability and statistics are not side dishes to fallacy study. In many cases, they are the flashlight.",
  },
  {
    slug: "how-to-use-fallacy-language-without-becoming-insufferable",
    title: "How to Use Fallacy Language Without Becoming Insufferable",
    description:
      "A guide to using fallacy vocabulary with humility, precision, charity, and enough restraint to remain useful in actual conversation.",
    intro:
      "This article is for anyone who wants to think clearly without becoming the conversational equivalent of a smoke alarm taped to a trumpet.",
  },
  {
    slug: "argument-maps-for-common-fallacies",
    title: "Argument Maps for Common Fallacies",
    description:
      "A guide to using argument maps to surface hidden premises, show exactly where support fails, and teach fallacies visually.",
    intro:
      "This article treats argument maps as x-rays for reasoning. They do not replace judgment, but they do make the fractures easier to see.",
  },
  {
    slug: "teaching-fallacies-through-debate-editorials-and-news-analysis",
    title: "Teaching Fallacies Through Debate, Editorials, and News Analysis",
    description:
      "A classroom article on how to teach fallacies through live rhetorical materials rather than canned examples alone.",
    intro:
      "This article focuses on the practical side of teaching: prompts, assignments, rubrics, and activity types built around real debates and public argument.",
  },
  {
    slug: "the-role-of-analogy-in-rational-criticism",
    title: "The Role of Analogy in Rational Criticism",
    description:
      "A wider theory article on why analogy is one of the most powerful tools in criticism, teaching, and argument repair when used carefully.",
    intro:
      "This article widens the lens from fallacy rebuttal alone and asks why analogy, at its best, is one of reason's favorite crowbars.",
  },
  {
    slug: "fallacies-in-the-age-of-algorithmic-media",
    title: "Fallacies in the Age of Algorithmic Media",
    description:
      "An article on how feeds, clipping, outrage incentives, virality, and attention economies change which fallacies thrive in public discourse.",
    intro:
      "This article explores what happens when bad arguments are no longer merely spoken or printed, but optimized for velocity, salience, and engagement.",
  },
];

const featureArticleDefinitions = [
  {
    slug: "how-one-alabama-map-headline-invites-several-fallacies",
    title: "How One Alabama Map Headline Invites Several Fallacies",
    date: "2026-05-26",
    description:
      "A headline-level feature on how compressed legal and political wording can invite single-cause, false-dilemma, correlation-causation, equivocation, and cherry-picking mistakes.",
    intro:
      "This feature takes a current headline and asks a sharper question than 'Do I agree with it?' It asks what reasoning shortcuts a reader can slide into if a short headline is treated like a full explanation.",
  },
];

const theorySourceCatalog = {
  fallaciesSep: {
    title: "Fallacies (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/fallacies/",
  },
  fallaciesIep: {
    title: "Fallacies (Internet Encyclopedia of Philosophy)",
    url: "https://iep.utm.edu/fallacy/",
  },
  criticalThinkingIep: {
    title: "Critical Thinking (Internet Encyclopedia of Philosophy)",
    url: "https://iep.utm.edu/critical-thinking/",
  },
  argumentSep: {
    title: "Argument and Argumentation (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/argument/",
  },
  analogySep: {
    title: "Analogy and Analogical Reasoning (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/reasoning-analogy/",
  },
  inductionSep: {
    title: "The Problem of Induction (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/induction-problem/",
  },
  logicInductiveSep: {
    title: "Inductive Logic (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/logic-inductive/",
  },
  statisticsSep: {
    title: "Philosophy of Statistics (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/statistics/",
  },
  boundedRationalitySep: {
    title: "Bounded Rationality (Stanford Encyclopedia of Philosophy)",
    url: "https://plato.stanford.edu/entries/bounded-rationality/",
  },
  openstaxBiases: {
    title: "Overcoming Cognitive Biases and Engaging in Critical Reflection (OpenStax)",
    url: "https://openstax.org/books/introduction-philosophy/pages/2-2-overcoming-cognitive-biases-and-engaging-in-critical-reflection",
  },
  openstaxArguments: {
    title: "Arguments (OpenStax Introduction to Philosophy)",
    url: "https://openstax.org/books/introduction-philosophy/pages/5-3-arguments",
  },
  openstaxInferences: {
    title: "Types of Inferences (OpenStax Introduction to Philosophy)",
    url: "https://openstax.org/books/introduction-philosophy/pages/5-4-types-of-inferences",
  },
  openstaxFallacies: {
    title: "Informal Fallacies (OpenStax Introduction to Philosophy)",
    url: "https://openstax.org/books/introduction-philosophy/pages/5-5-informal-fallacies",
  },
  argumentMapsStudy: {
    title: "Argument Maps Improve Critical Thinking (Twardy, Teaching Philosophy)",
    url: "https://philpapers.org/archive/TWAAMI.pdf",
  },
  rationaleOnline: {
    title: "Rationale: Teaching Critical Thinking with Argument Maps",
    url: "https://www.rationaleonline.com/",
  },
  waltonCriticalQuestions: {
    title: "Advances in the Theory of Argumentation Schemes and Critical Questions (Informal Logic)",
    url: "https://informallogic.ca/index.php/informal_logic/article/view/485",
  },
  gemsTips: {
    title: "Tips for Creating Custom Gems (Gemini Apps Help)",
    url: "https://support.google.com/gemini/answer/15235603?hl=en",
  },
  gemsUse: {
    title: "Use Gems in Gemini Apps (Gemini Apps Help)",
    url: "https://support.google.com/gemini/answer/15146780",
  },
  pewSocialMediaUse: {
    title: "Americans’ Social Media Use (Pew Research Center, 2024)",
    url: "https://www.pewresearch.org/internet/2024/01/31/americans-social-media-use/",
  },
  pewNewsSocial: {
    title: "Many Americans Find Value in Getting News on Social Media, but Concerns About Inaccuracy Have Risen (Pew Research Center, 2024)",
    url: "https://www.pewresearch.org/short-reads/2024/02/07/many-americans-find-value-in-getting-news-on-social-media-but-concerns-about-inaccuracy-have-risen/",
  },
  pewTechCompanies: {
    title: "How Americans View Big Tech in 2024 (Pew Research Center)",
    url: "https://www.pewresearch.org/internet/2024/04/29/americans-views-of-technology-companies-2/",
  },
  explainableFallacyDetection: {
    title: "Robust and Explainable Identification of Logical Fallacies in Natural Language Arguments (arXiv, 2022)",
    url: "https://arxiv.org/abs/2212.07425",
  },
  logicalFallacyDetection: {
    title: "Logical Fallacy Detection (Findings of ACL EMNLP 2022)",
    url: "https://aclanthology.org/2022.findings-emnlp.532.pdf",
  },
};

const dialogueAssessmentChoices = [
  { key: "left-formal", label: "Left Formal" },
  { key: "left-informal", label: "Left Informal" },
  { key: "none", label: "None" },
  { key: "right-informal", label: "Right Informal" },
  { key: "right-formal", label: "Right Formal" },
];

function buildDialogueAssessmentItem({
  id,
  answerKey,
  fallacyName = "",
  fallacySlug = "",
  explanation,
  turns,
}) {
  if (!Array.isArray(turns) || turns.length !== 6) {
    throw new Error(`Dialogue assessment item "${id}" must contain exactly 6 turns.`);
  }
  return {
    id,
    answerKey,
    fallacyName,
    fallacySlug,
    explanation,
    turns: turns.map((text, index) => ({
      side: index % 2 === 0 ? "left" : "right",
      text,
    })),
  };
}

const dialogueAssessmentBank = [
  buildDialogueAssessmentItem({
    id: "lf-01",
    answerKey: "left-formal",
    fallacyName: "Affirming the consequent",
    fallacySlug: "affirming-the-consequent",
    explanation:
      "The left speaker moves from if the dog got out, the side gate would be open and the gate is open to therefore the dog got out. An open gate could have other causes too.",
    turns: [
      "If the dog got out, the side gate would be open.",
      "Sure, that would be one sign.",
      "The side gate is open right now.",
      "So something may have happened there.",
      "Then the dog got out.",
      "An open gate points somewhere, but it does not prove who went through it.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-02",
    answerKey: "left-formal",
    fallacyName: "Denying the antecedent",
    fallacySlug: "denying-the-antecedent",
    explanation:
      "The left speaker assumes that because Dad did not get his bonus early, the beach trip cannot happen. But that bonus was only one possible way to pay for it, not the only way.",
    turns: [
      "If Dad gets his bonus early, we can go to the beach this weekend.",
      "Right, that would make the trip easy to afford.",
      "He did not get the bonus early.",
      "So that money is not coming in.",
      "Then we cannot go to the beach.",
      "That shuts one door, but why think every way of paying for the trip is gone?",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-03",
    answerKey: "left-formal",
    fallacyName: "Affirming a disjunct",
    fallacySlug: "affirming-a-disjunct",
    explanation:
      "The left speaker treats an either-or claim as if only one option could be true. From either Maya forgot the tickets or Luis forgot them and Maya forgot hers, the speaker jumps to so Luis did not forget his.",
    turns: [
      "Either Maya forgot the tickets or Luis forgot them.",
      "Those are two real possibilities.",
      "Maya definitely forgot hers.",
      "So Maya is in the story.",
      "Then Luis did not forget his.",
      "That only follows if both people could not have forgotten.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-04",
    answerKey: "left-formal",
    fallacyName: "Undistributed middle",
    fallacySlug: "undistributed-middle",
    explanation:
      "The left speaker notes that all school buses are yellow and that a van is yellow, then concludes it must be a school bus. Sharing one broad feature does not prove identity.",
    turns: [
      "All school buses are yellow.",
      "Yes.",
      "That van is yellow.",
      "It is.",
      "Then it must be a school bus.",
      "Lots of things are yellow; that color alone does not settle what it is.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-05",
    answerKey: "left-formal",
    fallacyName: "Affirming the consequent",
    fallacySlug: "affirming-the-consequent",
    explanation:
      "The left speaker infers that the freezer failed from the fact that the ice cream melted, even though melting could have other causes. One expected result is treated as if it proved the cause.",
    turns: [
      "If the freezer stopped working, the ice cream would melt.",
      "That would be one likely result.",
      "The ice cream melted.",
      "So something went wrong.",
      "Then the freezer stopped working.",
      "Maybe, but melted ice cream can come from more than one cause.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-06",
    answerKey: "left-formal",
    fallacyName: "Denying the antecedent",
    fallacySlug: "denying-the-antecedent",
    explanation:
      "The left speaker reasons that because the phone was not left at home, there is no way the group text was missed. But having the phone nearby does not rule out every other way to miss a message.",
    turns: [
      "If I left my phone at home, I would miss the group text.",
      "That would do it.",
      "I did not leave my phone at home.",
      "Okay.",
      "Then I could not have missed the group text.",
      "Keeping your phone with you does not rule out every other way to miss a message.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-07",
    answerKey: "left-formal",
    fallacyName: "Masked man fallacy",
    fallacySlug: "masked-man-fallacy",
    explanation:
      "The left speaker argues from a difference in what is known to a difference in identity. Knowing Coach Miller does not tell you who is inside the mascot suit.",
    turns: [
      "You know Coach Miller.",
      "Of course.",
      "You do not know who is inside the lion mascot suit.",
      "Right.",
      "Then Coach Miller cannot be inside the lion mascot suit.",
      "Knowing Coach Miller does not tell you who is inside the mascot suit.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "lf-08",
    answerKey: "left-formal",
    fallacyName: "Undistributed middle",
    fallacySlug: "undistributed-middle",
    explanation:
      "The left speaker notes that all firefighters wear heavy boots and that an uncle is wearing heavy boots, then concludes he must be a firefighter. The shared middle feature is too broad to prove identity.",
    turns: [
      "All firefighters wear heavy boots.",
      "Sure.",
      "My uncle is wearing heavy boots.",
      "He is.",
      "Then he must be a firefighter.",
      "Boots like that show up in lots of jobs, so the match is too broad.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-01",
    answerKey: "left-informal",
    fallacyName: "Ad hominem",
    fallacySlug: "ad-hominem",
    explanation:
      "The left speaker dismisses Leah's driving advice by attacking her personality and hobbies rather than answering the advice itself. The target is the person, not the support for the claim.",
    turns: [
      "I am not taking Leah's advice about safe driving seriously; she still posts goofy videos all day.",
      "That may say something about her hobbies, but what about her point about using seat belts?",
      "People like that always want attention.",
      "That still does not answer her warning.",
      "If she were a serious person, maybe I would listen.",
      "You are talking about Leah instead of whether the advice is good.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-02",
    answerKey: "left-informal",
    fallacyName: "False dilemma",
    fallacySlug: "false-dilemma",
    explanation:
      "The left speaker compresses the options into buying one car today or staying stuck forever, ignoring other live options such as waiting, saving, or choosing a different car.",
    turns: [
      "Either we buy this car today or we admit we will never get ahead.",
      "Could we keep looking for a cheaper one next week?",
      "Waiting is just another name for giving up.",
      "It could also mean choosing more carefully.",
      "No, it is really buy it now or stay stuck forever.",
      "That leaves out several live options.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-03",
    answerKey: "left-informal",
    fallacyName: "Cherry picking",
    fallacySlug: "cherry-picking",
    explanation:
      "The left speaker points to three strong tomato plants and treats them as proof that the whole garden is thriving. The argument looks stronger only because the weaker evidence is left out.",
    turns: [
      "This garden is doing great; look at these three big tomatoes.",
      "Why those three plants and not the whole yard?",
      "Because those are the clearest ones.",
      "The rest of the garden looks wilted.",
      "The important point is that these three prove the garden is thriving.",
      "They prove three plants look good, not that the whole garden is thriving.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-04",
    answerKey: "left-informal",
    fallacyName: "Appeal to authority",
    fallacySlug: "appeal-to-authority",
    explanation:
      "The left speaker leans on a movie star's endorsement as if that settled whether the skin cream works. Fame may catch attention, but it does not replace evidence.",
    turns: [
      "That skin cream must work. A famous actor said it changed her life.",
      "Her praise is interesting, but what do the ingredients and studies show?",
      "Someone that famous would not recommend junk.",
      "Fame is not the same thing as proof.",
      "Her name settles enough of it for me.",
      "An endorsement is not the same as evidence.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-05",
    answerKey: "left-informal",
    fallacyName: "Red herring",
    fallacySlug: "red-herring",
    explanation:
      "The left speaker is asked about scratching the car and replies with points about helping around the house. Those points may be favorable, but they do not answer the original question.",
    turns: [
      "Before we talk about who scratched the car, remember how much I help around the house.",
      "I appreciate that, but I asked about the scratch on the car.",
      "I am the one who takes out the trash and mows the lawn.",
      "That still does not answer the question.",
      "My point is that I am a helpful person.",
      "Maybe, but I am still asking who scratched the car.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-06",
    answerKey: "left-informal",
    fallacyName: "Hasty generalization",
    fallacySlug: "hasty-generalization",
    explanation:
      "The left speaker generalizes from two rude kids to a broad claim about everyone on Pine Street. The sample is far too small to support that conclusion.",
    turns: [
      "Two kids from Pine Street were rude to me, so people from Pine Street are rude.",
      "That sounds broader than your evidence.",
      "It shows what that street is like.",
      "It shows what two kids were like.",
      "After meeting those two, I do not trust anyone from there.",
      "Two bad encounters are not enough to judge everyone on the street.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-07",
    answerKey: "left-informal",
    fallacyName: "False equivalence",
    fallacySlug: "false-equivalence",
    explanation:
      "The left speaker treats borrowing a hoodie without asking as basically the same as stealing a car because both involve taking something. The comparison flattens a major difference.",
    turns: [
      "Borrowing my hoodie without asking is basically the same as stealing a car. Both are stealing.",
      "Both are wrong, but they are not wrong in the same way.",
      "Taking is taking.",
      "That is too thin a similarity to flatten the difference.",
      "Once you take something that is not yours, the cases are equal.",
      "A hoodie and a car are not the same kind or size of wrongdoing.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "li-08",
    answerKey: "left-informal",
    fallacyName: "Appeal to motive",
    fallacySlug: "appeal-to-motive",
    explanation:
      "The left speaker treats a neighbor's motive as if it settled the playground safety complaint itself. Motive may matter, but it does not answer whether the swing is actually broken.",
    turns: [
      "He only says the playground is unsafe because he wants the neighborhood to vote for his plan.",
      "That motive might be worth noticing, but what about the broken swing he pointed to?",
      "People do not complain unless they want something.",
      "Sometimes they do, but that still would not settle whether the swing is broken.",
      "His motive tells me enough about how seriously to take the complaint.",
      "Motive may matter, but it does not answer the safety question by itself.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-01",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The speakers treat the family tie as a reason for stricter checking, not as an automatic refutation of the claim itself.",
    turns: [
      "The seller is my cousin, so I want another mechanic to look at the car before we trust the pitch.",
      "That seems fair. The family tie gives us a reason to check more carefully, not a reason to say the car is bad automatically.",
      "Exactly. I am not calling the car bad. I just want the facts checked by someone independent.",
      "Then the real question is whether the car still looks good after the inspection.",
      "Yes, and if it does, the car may still be fine.",
      "Good. That keeps the concern tied to evidence instead of turning it into a personal attack.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-02",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The speakers describe a real short-term limit created by the bus schedule rather than pretending that all possible options have vanished forever.",
    turns: [
      "The last bus leaves in ten minutes, so tonight we probably have to choose between running for it or calling a ride.",
      "That sounds like a real time limit, not a claim that those are the only travel options in life.",
      "Right. Other options existed earlier, but at this moment they have dropped away.",
      "Then the narrow choice comes from the clock, not from squeezing the argument.",
      "Exactly. I would be saying something stronger only if I claimed no other options had ever existed.",
      "Then the tight choice comes from the deadline, not from squeezing the options.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-03",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The speakers use the electrician's judgment as a guide while still appealing to the underlying evidence and relevant skill.",
    turns: [
      "I trust the electrician more than us on the wiring issue, but I still want to see what he found.",
      "That seems careful. His skill gives us a reason to start with his view, not a reason to stop asking questions.",
      "Exactly. If his reasons are weak, his job title will not save them.",
      "So the expert is a guide, not a magic stamp.",
      "Right. The goal is to use the right expert and still look at the evidence.",
      "Then the expert is being used for what he actually knows, not as a shortcut around evidence.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-04",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The analogy is limited to one relevant point, and the speaker makes that limit explicit instead of pretending the two cases are identical.",
    turns: [
      "Learning to drive is like learning piano in one way: short practice sessions work better than cramming.",
      "So you are not saying driving and piano are the same thing in every respect.",
      "Not at all. I only mean the steady-practice part carries over.",
      "Then the comparison is narrow and relevant.",
      "Exactly. If I tried to transfer every detail, the comparison would break.",
      "Then the analogy is doing a small job instead of pretending the two things are identical.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-05",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The anecdote is used as a reason to investigate further, not as if it were enough to settle a broad claim.",
    turns: [
      "My bike chain snapped on this brand, so I want to check whether lots of riders had the same problem.",
      "That sounds like using your case as a clue, not as a final verdict.",
      "Exactly. I am raising a question, not proving the whole brand is bad.",
      "Then the next step is to look for wider reports.",
      "Yes. One story can start the search even when it cannot finish it.",
      "Good, as long as the story stays a lead and not the whole case.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-06",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The speakers notice a correlation but openly refuse to treat it as proof of causation without better evidence.",
    turns: [
      "Kids who sleep more often get better grades, but I do not think that alone proves sleep caused the grades.",
      "So you are treating the pattern as a clue, not as a finished explanation.",
      "Exactly. We would still need timing and other evidence.",
      "Then the link is being handled carefully instead of being inflated into cause.",
      "Right. It gives us a reason to study more, not to close the case.",
      "That keeps the uncertainty honest.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-07",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The speakers treat the salesman's motive as a reason for caution while refusing to let motive alone settle the truth of the claim.",
    turns: [
      "The salesman wants to close the deal, so I want to double-check what he says about the truck.",
      "That makes sense. His motive gives us a reason to be cautious, not a reason to call every claim false.",
      "Exactly. If the facts check out, the truck may still be a good buy.",
      "Then you are separating his incentive from the truth of the claim.",
      "Yes. Motive can tell us to look harder without doing the looking for us.",
      "Good. That keeps suspicion from pretending to be proof.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "nf-08",
    answerKey: "none",
    explanation:
      "No fallacy is committed. The speakers insist on judging people in the past by the information they had at the time, which avoids unfair hindsight.",
    turns: [
      "People watching that storm years ago did not have the radar tools we have now, so we should be careful about saying they obviously should have known everything.",
      "That is different from excusing every mistake. It just means using the information they actually had.",
      "Right. Some warning signs were there, but hindsight makes the picture cleaner than it looked at the time.",
      "So the caution is against sneaking our later knowledge back into their moment.",
      "Exactly. The goal is fairness, not automatic pardon.",
      "That makes sense. Judge them by their information, not ours.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-01",
    answerKey: "right-informal",
    fallacyName: "Ad hominem",
    fallacySlug: "ad-hominem",
    explanation:
      "The right speaker dismisses the stop-sign argument by attacking Mia's lifestyle instead of answering the safety point she raised.",
    turns: [
      "Mia says the new stop sign would make this corner safer.",
      "Mia spends all day posting dance videos, so I am not taking her traffic ideas seriously.",
      "That says something about Mia, not yet about the corner.",
      "People like that do not know how streets really work.",
      "But the question was whether cars are speeding through the turn.",
      "Maybe, but I still trust regular drivers more than someone like Mia.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-02",
    answerKey: "right-informal",
    fallacyName: "False dilemma",
    fallacySlug: "false-dilemma",
    explanation:
      "The right speaker compresses the issue into sharing phone passwords or allowing cheating, which leaves out more limited and reasonable options.",
    turns: [
      "I am worried that giving out our phone passwords goes too far.",
      "Then you must want to let cheating go unchecked.",
      "I might want a narrower rule, not no rule.",
      "In practice it is one or the other: full access or no honesty.",
      "That seems to leave out other live options.",
      "In a problem like this, those really are the only two choices.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-03",
    answerKey: "right-informal",
    fallacyName: "Cherry picking",
    fallacySlug: "cherry-picking",
    explanation:
      "The right speaker points to three good weigh-ins and treats them as decisive while ignoring the broader month-long pattern.",
    turns: [
      "I am not sure this diet is really working overall.",
      "Look at the three mornings this week when the scale number dropped.",
      "Why only those mornings and not the full month?",
      "Because those are the clearest proof that it works.",
      "The rest of the month mostly went the other way.",
      "The best mornings tell the real story, so those are the ones I care about.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-04",
    answerKey: "right-informal",
    fallacyName: "Appeal to authority",
    fallacySlug: "appeal-to-authority",
    explanation:
      "The right speaker treats a famous athlete's endorsement as if it settled whether the drink is safe. Prestige is doing evidential work the claim has not earned.",
    turns: [
      "I am still not sure this energy drink is safe.",
      "A famous athlete says he drinks it every day, so that should settle it.",
      "His routine may be interesting, but what do the health warnings say?",
      "Someone at that level would not put bad stuff in his body.",
      "That sounds stronger than the actual evidence.",
      "When a top athlete trusts it, I do not need more proof.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-05",
    answerKey: "right-informal",
    fallacyName: "Red herring",
    fallacySlug: "red-herring",
    explanation:
      "The right speaker answers a question about the missing cake by shifting to how helpful he is around the house. That may be flattering, but it does not answer the issue raised.",
    turns: [
      "Did you eat the cake I was saving for tomorrow?",
      "Before we talk about that, remember how often I help with dinner.",
      "I do remember that, but I asked about the cake.",
      "I am the one who usually washes the dishes too.",
      "That still does not answer the question.",
      "My help around the house matters more to me than one missing cake.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-06",
    answerKey: "right-informal",
    fallacyName: "Hasty generalization",
    fallacySlug: "hasty-generalization",
    explanation:
      "The right speaker leaps from one missed plumbing appointment to a broad claim about all plumbers. The conclusion outruns the evidence.",
    turns: [
      "One plumber missed our appointment, but I am not sure that tells us about all plumbers.",
      "After something like that, it is obvious plumbers are all unreliable.",
      "That sounds broader than one bad appointment can support.",
      "You only need one clear case to see the type.",
      "Not if the claim is about a whole line of work.",
      "One failure like that tells you plenty about the whole bunch.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-07",
    answerKey: "right-informal",
    fallacyName: "False equivalence",
    fallacySlug: "false-equivalence",
    explanation:
      "The right speaker treats forgetting to text back as basically the same as lying to someone's face because both hide information. The comparison erases an important difference.",
    turns: [
      "Forgetting to text back is rude, but it is not the same as lying to your face.",
      "It is basically the same because both leave someone in the dark.",
      "Both may frustrate people, but they are not equal in kind.",
      "Once you hide something, the moral difference disappears.",
      "That flattens an important difference.",
      "If both acts hide information, then they are the same sort of wrong.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "ri-08",
    answerKey: "right-informal",
    fallacyName: "Slippery slope",
    fallacySlug: "slippery-slope",
    explanation:
      "The right speaker projects from one make-up quiz to a flood of grade requests without showing the steps in between. The slide is asserted rather than argued for.",
    turns: [
      "I am okay with letting one student retake a quiz after being sick.",
      "Once you allow one retake, the next step is everyone begging for new grades.",
      "That is a long chain. What makes the later steps follow?",
      "That is just how these things always go once the first exception is made.",
      "I still do not hear the link between this one case and the flood you predict.",
      "Once the first exception stands, the rest is only a matter of time.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-01",
    answerKey: "right-formal",
    fallacyName: "Affirming the consequent",
    fallacySlug: "affirming-the-consequent",
    explanation:
      "The right speaker reasons from the ringing smoke alarm back to one specific cause, even though the alarm could have been set off in other ways too.",
    turns: [
      "If the cake burned, the smoke alarm would ring.",
      "The smoke alarm rang, so the cake burned.",
      "Could the alarm ring for some reason other than a burnt cake?",
      "Maybe, but the condition already points us where we need to go.",
      "It gives one route to the alarm, not the only route.",
      "If the alarm rang, that gives me the answer I need.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-02",
    answerKey: "right-formal",
    fallacyName: "Denying the antecedent",
    fallacySlug: "denying-the-antecedent",
    explanation:
      "The right speaker infers that missing the early start means missing the bus, even though leaving early was only one sufficient way to catch it.",
    turns: [
      "If I leave early, I will catch the bus.",
      "You did not leave early, so you will not catch the bus.",
      "That only rules out one way to catch it.",
      "It rules out the way we counted on.",
      "Yes, but not every possible way.",
      "If the early-leaving condition is gone, the bus-catching result is gone too.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-03",
    answerKey: "right-formal",
    fallacyName: "Affirming a disjunct",
    fallacySlug: "affirming-a-disjunct",
    explanation:
      "The right speaker treats a disjunction as exclusive without showing that only one option could be true.",
    turns: [
      "Either the dog knocked over the trash or the wind did.",
      "The wind was strong, so the dog did not do it.",
      "That only follows if the two possibilities could not both be involved.",
      "I took the either-or to settle that.",
      "But you have not shown that it was exclusive.",
      "If the wind is in, the dog is out.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-04",
    answerKey: "right-formal",
    fallacyName: "Undistributed middle",
    fallacySlug: "undistributed-middle",
    explanation:
      "The right speaker infers that Sam is a lifeguard because both Sam and lifeguards wear red shirts. Sharing a broad feature is not enough.",
    turns: [
      "All lifeguards wear red shirts.",
      "Sam is wearing a red shirt, so Sam must be a lifeguard.",
      "That only shows Sam shares one feature with lifeguards.",
      "It is still a strong clue.",
      "Not if lots of other people wear red shirts too.",
      "The red shirt is the mark that tells us what he is.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-05",
    answerKey: "right-formal",
    fallacyName: "Masked man fallacy",
    fallacySlug: "masked-man-fallacy",
    explanation:
      "The right speaker infers a difference in identity from a difference in what is known. Not recognizing the person in costume does not prove it is not Aunt Rosa.",
    turns: [
      "You know Aunt Rosa.",
      "I do not know who is inside the chicken costume, so it is not Aunt Rosa.",
      "That conclusion assumes knowing Aunt Rosa settles who is inside the costume.",
      "If I cannot tell it is her, why think it could be her?",
      "Not recognizing the costume person is not the same as proving it is not Aunt Rosa.",
      "If the person in the costume is unknown to me, then it cannot be Aunt Rosa.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-06",
    answerKey: "right-formal",
    fallacyName: "Affirming the consequent",
    fallacySlug: "affirming-the-consequent",
    explanation:
      "The right speaker infers water damage from a black screen, even though a black screen can have other causes too.",
    turns: [
      "If the phone got wet, the screen would go black.",
      "The screen went black, so the phone got wet.",
      "Could the screen go black for some reason other than water?",
      "Maybe, but the match still points to water.",
      "A match does not prove it was the only possible cause.",
      "Once the screen goes black, that gives me the reason.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-07",
    answerKey: "right-formal",
    fallacyName: "Denying the antecedent",
    fallacySlug: "denying-the-antecedent",
    explanation:
      "The right speaker treats the absence of a power outage as if it guaranteed the garage door would open. But the original conditional named only one sufficient reason for failure.",
    turns: [
      "If the power is out, the garage door will not open.",
      "The power is not out, so the garage door will open.",
      "That only rules out one reason it might stay shut.",
      "It rules out the reason we named.",
      "Yes, but not every possible reason.",
      "If the power problem is gone, the closed-door result should go with it.",
    ],
  }),
  buildDialogueAssessmentItem({
    id: "rf-08",
    answerKey: "right-formal",
    fallacyName: "Exclusive premises",
    fallacySlug: "exclusive-premises",
    explanation:
      "The right speaker tries to get a positive conclusion from two negative premises. The structure does not support that jump.",
    turns: [
      "No dogs are cats.",
      "And no cats are parrots, so all dogs are parrots.",
      "That conclusion sounds dropped in from somewhere else.",
      "Why? The two premises draw the lines pretty clearly.",
      "They draw two negative lines, but they do not by themselves create that positive conclusion.",
      "Those two exclusions leave only one place for dogs to go.",
    ],
  }),
];

function validateDialogueAssessmentBank(bank) {
  if (!Array.isArray(bank) || bank.length !== 40) {
    throw new Error(`Dialogue assessment bank must contain exactly 40 items; found ${bank?.length || 0}.`);
  }
  const choiceKeys = new Set(dialogueAssessmentChoices.map((choice) => choice.key));
  const counts = new Map(dialogueAssessmentChoices.map((choice) => [choice.key, 0]));
  const ids = new Set();

  for (const item of bank) {
    if (ids.has(item.id)) {
      throw new Error(`Duplicate dialogue assessment item id "${item.id}".`);
    }
    ids.add(item.id);
    if (!choiceKeys.has(item.answerKey)) {
      throw new Error(`Dialogue assessment item "${item.id}" has invalid answer key "${item.answerKey}".`);
    }
    if (!Array.isArray(item.turns) || item.turns.length !== 6) {
      throw new Error(`Dialogue assessment item "${item.id}" must contain exactly 6 turns.`);
    }
    counts.set(item.answerKey, (counts.get(item.answerKey) || 0) + 1);
  }

  for (const choice of dialogueAssessmentChoices) {
    if ((counts.get(choice.key) || 0) !== 8) {
      throw new Error(`Dialogue assessment bank must contain exactly 8 items for "${choice.key}"; found ${counts.get(choice.key) || 0}.`);
    }
  }
}

validateDialogueAssessmentBank(dialogueAssessmentBank);

const analogyResponseOverrides = {
  "Absence of evidence fallacy":
    "That's like checking an empty shelf in one room and announcing the whole archive must still contain the missing file somewhere else. If the search conditions were good enough to leave traces, then repeated failure to find them does count against the claim.",
  "Ad hominem":
    "That's like saying the thermometer must be wrong because you dislike the person holding it. The personal jab does not answer whether the reading itself is accurate.",
  "Appeal to authority":
    "That's like accepting a bridge design because a famous chef approved it. Reputation only helps when the authority is actually in the right lane of expertise and the claim still fits the evidence.",
  "Appeal to emotion":
    "That's like turning up the soundtrack and calling the volume an argument. The feeling may be real, but it still has to be connected to evidence before it can carry the conclusion.",
  "Appeal to fear":
    "That's like yanking the fire alarm in order to win a zoning dispute. Panic can move people, but it does not prove the threatened outcome is actually supported.",
  "Appeal to nature":
    "That's like praising poison ivy for being natural and condemning eyeglasses for being artificial. The label tells you where something came from, not whether it is good, safe, or wise.",
  "Argument from ignorance":
    "That's like saying the locked room must contain treasure because nobody has opened the door yet. Lack of disproof is not the same thing as proof.",
  "Base rate fallacy":
    "That's like hearing a smoke detector beep once and forgetting you are standing in a building full of low batteries. The vivid clue feels decisive only because the background rate was ignored.",
  "Begging the question":
    "That's like using your own house key as proof that you deserve to be inside. The conclusion is being smuggled back in as if it were independent support.",
  "Cherry picking":
    "That's like showing only the sunny frames from a stormy week and calling it the full weather report. The argument looks stronger only because the missing evidence was left offstage.",
  "Correlation is not causation":
    "That's like noticing that umbrellas and wet sidewalks appear together and deciding umbrellas cause rain. The pattern might matter, but the cause still has to be established rather than assumed.",
  "Equivocation":
    "That's like winning a card game by quietly changing what the wild card means halfway through the hand. The wording stays similar while the meaning shifts underneath it.",
  "False analogy":
    "That's like claiming a bicycle should float because a boat also carries people. A shared surface feature does not guarantee the deeper structure is alike where it matters.",
  "False dilemma":
    "That's like insisting a city map has only north and south because east and west complicate the story. The argument sounds decisive only because the live alternatives were erased.",
  "False equivalence":
    "That's like saying a paper cut and a broken leg are basically the same because both are injuries. The comparison uses a thin similarity to flatten an important difference.",
  "Hasty generalization":
    "That's like tasting one spoonful from a burnt corner of the soup and condemning the entire pot. The leap from small sample to broad conclusion outruns what the evidence can support.",
  "Moving the goalpost":
    "That's like promising a race ends at the oak tree and then moving the finish line to the next hill when someone gets there first. The standard changes only to keep the result from counting.",
  "No True Scotsman":
    "That's like redrawing the target after the arrow lands outside the circle. The standard is being revised after the counterexample appears so the original claim can survive untouched.",
  "Post hoc ergo propter hoc":
    "That's like hearing the rooster crow before sunrise and concluding the rooster pulled the sun over the horizon. Mere sequence is being treated as if it already proved causation.",
  "Red herring":
    "That's like answering a fire alarm by criticizing the paint color on the extinguisher. The detour may be vivid, but it is still not an answer to the issue that raised the alarm.",
  "Slippery slope":
    "That's like claiming one library late fee will inevitably end in martial law. A real chain of consequences needs support for each step, not just dread about the final one.",
  "Straw man argument":
    "That's like replacing your opponent's chessboard with a toy checkerboard and then bragging that you defeated their strategy. The rebuttal wins only because it targeted a weaker stand-in instead of the real position.",
  "Survivorship bias":
    "That's like studying only the planes that made it back and then deciding the missing planes must not matter. The lesson is distorted because the silent failures were never counted.",
  "Tu quoque":
    "That's like dismissing a doctor's warning about smoking because the doctor smokes. The hypocrisy may be real, but it does not by itself make the warning false.",
};

const auditedAnalogyResponseOverrides = {
  "Abstraction denial":
    "That's like saying there is no such thing as a traffic jam because there are only individual cars. A higher-level pattern can be real and causally important even though it is made out of lower-level parts.",
  "Abstraction fallacy":
    "That's like treating the speed limit sign as proof that no car can ever go faster. A useful model or trend is being treated as if reality were logically forbidden to depart from it.",
  "Affirmative conclusion from a negative premise":
    "That's like saying, 'This fruit is not rotten. Rotten fruit is unsafe. Therefore this fruit is safe.' Ruling out one bad category does not automatically prove the positive opposite.",
  "Affirming a disjunct":
    "That's like saying a room is either crowded or noisy, and because it is crowded it cannot also be noisy. An ordinary 'or' is being treated as if only one option could be true.",
  "Affirming the consequent":
    "That's like saying, 'If it rained, the sidewalk would be wet. The sidewalk is wet, so it rained.' The observed result could have come from other causes.",
  "All or nothing fallacy":
    "That's like throwing away the whole meal because one dish was oversalted. A flaw in part of the package is being treated as total failure of the whole.",
  "Anecdotal fallacy":
    "That's like judging the climate from one hot afternoon on your porch. A vivid personal case is being treated as stronger than broader and more representative evidence.",
  "Appeal to accomplishment":
    "That's like asking a gold-medal swimmer to sign off on your tax return. Success in one field does not automatically transfer authority to another.",
  "Appeal to consequences":
    "That's like denying the storm warning because boarding the windows would be expensive. The cost or comfort of the conclusion is being mistaken for evidence about whether it is true.",
  "Appeal to flattery":
    "That's like telling jurors they are far too intelligent to acquit, and then counting the compliment as proof. Praise is being used where evidence should be.",
  "Appeal to motive":
    "That's like ignoring the math on a grocery bill because you think the cashier hopes for a tip. Suspected motive does not by itself answer the claim.",
  "Appeal to novelty":
    "That's like buying the newest stove and assuming dinner must taste better. Newness is being treated as if it were evidence of superiority.",
  "Appeal to pity":
    "That's like asking the scale to change your weight because you've had a hard week. Sympathy may be appropriate, but it does not prove the conclusion.",
  "Appeal to poverty":
    "That's like treating a weather forecast as truer because it came from a person in worn shoes. Humble origins do not convert a claim into evidence.",
  "Appeal to probability":
    "That's like buying one lottery ticket and speaking as if winning were practically on the calendar. Mere possibility is being inflated into likelihood or inevitability.",
  "Appeal to ridicule":
    "That's like laughing at the pilot's haircut and calling the joke a refutation of the flight plan. Mockery is being substituted for reasons.",
  "Appeal to spite":
    "That's like voting for a bad bridge plan just because it would annoy the mayor you hate. Resentment toward a target is being used as support for the conclusion.",
  "Appeal to tradition":
    "That's like insisting the classroom clock must stay wrong because it has always been five minutes slow. Longevity alone does not justify the practice.",
  "Appeal to wealth":
    "That's like assuming a billionaire's stethoscope makes a diagnosis better. Money and status are being mistaken for evidence.",
  "Argument from fallacy":
    "That's like spotting a typo in the map legend and concluding the city itself must not exist. A bad argument for a claim does not automatically make the claim false.",
  "Argument from incredulity":
    "That's like saying airplanes cannot fly because you personally cannot picture the aerodynamics. Difficulty imagining a mechanism is not evidence against the thing.",
  "Argument from repetition":
    "That's like hitting refresh on the same rumor and counting each reload as new proof. Familiarity is being mistaken for support.",
  "Argument from silence":
    "That's like taking a professor's missed email as confirmation that your answer key must be right. Non-response is being treated as endorsement.",
  "Argumentum ad baculum":
    "That's like saying, 'Agree that the bridge is safe, or lose your job.' The threat may secure compliance, but it does not add support.",
  "Argumentum ad populum":
    "That's like calling a restaurant healthy because the line is long. Popularity is being used as if it were proof.",
  "Artificial negation":
    "That's like saying that because I do not collect stamps, I must be positively committed to anti-stamp activism. Mere non-belief is being inflated into a strong opposite claim.",
  "Association fallacy":
    "That's like throwing out a sound medical study because it traveled in the same folder as a bad one. The claim is being judged by a nearby association instead of its merits.",
  "Bare assertion fallacy":
    "That's like slamming a stamp on a blank file and calling the case closed. Confidence of assertion is replacing actual support.",
  "Bottom-up condemnation":
    "That's like meeting one rude nurse and deciding the next nurse must be rude too because she belongs to the same profession. A negative group judgment is being pasted onto the individual.",
  "Bottom-up justification":
    "That's like trusting a mechanic you've never met because another mechanic once saved your road trip. A group-level compliment is being used as proof about this particular member.",
  "Broken window fallacy":
    "That's like praising a house fire because it creates work for carpenters. The visible rebuilding is counted while the destroyed value and lost alternatives disappear.",
  "Chronological snobbery":
    "That's like dismissing Euclid because he lacked Wi-Fi. An idea is being rejected for being old rather than for being wrong.",
  "Circular cause and consequence":
    "That's like saying the club is exclusive because no one joins, and no one joins because it is exclusive. The loop is being treated as if it explained or justified itself.",
  "Composition fallacy":
    "That's like saying every brick is light enough to carry, so the whole building must be light enough to carry. What is true of parts need not be true of the whole.",
  "Confidence as a validator":
    "That's like treating the volume of a car horn as evidence that the driver has the right of way. Felt certainty is being mistaken for support.",
  "Conjunction fallacy":
    "That's like saying it is more likely a die rolled a six-and-landed-on-a-Tuesday than simply rolled a six. Adding details makes the story feel richer, not more probable.",
  "Contextomy":
    "That's like quoting 'I approve' from the sentence 'I do not approve' and calling the clip accurate. A sliced excerpt is being treated as if it preserved the original meaning.",
  "Continuum fallacy":
    "That's like saying we cannot call a person bald unless we can name the exact hair that made the difference. A blurry boundary is being mistaken for a useless concept.",
  "Definist fallacy":
    "That's like defining 'healthy' as 'whatever doctors recommend' and then announcing that every recommendation is healthy by definition. A contested issue is being smuggled into a definition.",
  "Demanding a mechanism":
    "That's like refusing to believe the bridge is swaying until you can derive every equation of the wind load. Incomplete mechanism is being used to dismiss strong evidence that the effect is real.",
  "Demanding negative proof":
    "That's like insisting the fire department prove there is no spark anywhere before you will stop calling it arson. The claimant is pushing the burden onto critics instead of supporting the claim.",
  "Denial of the epistemic gradient":
    "That's like forcing every weather forecast into only 'certain' or 'false' while banning words like 'likely' and 'unlikely.' A graded confidence judgment is being crushed into a crude binary.",
  "Denying a remote hypothetical":
    "That's like refusing to test a fire exit because most days the building does not burn. Rare cases can still matter when the principle is supposed to hold universally.",
  "Denying the antecedent":
    "That's like saying, 'If the alarm is set, the house is protected. The alarm is not set, so the house is unprotected.' The conclusion ignores other ways the result could still hold.",
  "Denying the correlative":
    "That's like saying a foul cannot count as a foul unless the game ended because of it. One side of a meaningful contrast is being redefined until the other side has nowhere to apply.",
  "Division fallacy":
    "That's like saying the orchestra is wealthy, so every violinist must be wealthy. What is true of the whole does not automatically distribute to each part.",
  "Ecological fallacy":
    "That's like saying the average household on this street is tall, so the person at number 18 must be tall. Group statistics are being pasted onto the individual.",
  "Empty refutation":
    "That's like shouting 'case closed' before opening the folder. Declaring something debunked is being used in place of showing the flaw.",
  "Epistemic/ontological conflation":
    "That's like saying the comfort of a bedtime lamp proves the monsters under the bed are real. The effects of believing something are being mistaken for evidence that the thing exists.",
  "Equivocation fallacy":
    "That's like hearing a toy car described as 'driving itself' and then concluding it must have a human driver's understanding of traffic. A harmless sense of a word is being stretched into a stronger one.",
  "Etymological fallacy":
    "That's like insisting a computer mouse must be a rodent because that was the older meaning of the word 'mouse.' Historical origin is being treated as present meaning.",
  "Exclusive premises":
    "That's like saying, 'No cats are dogs. Some dogs are not pets. Therefore some pets are not cats.' Two negatives were never enough to build the needed bridge.",
  "Existential fallacy":
    "That's like saying every dragon in my garage is asleep, therefore at least one dragon is in there snoring. A statement about all members of a class is being treated as proof that the class has members.",
  "Fallacy of many questions":
    "That's like asking, 'When did you stop cheating at cards?' and pretending the question itself counts as evidence. The answer is being forced to carry an unproved assumption.",
  "Fallacy of necessity":
    "That's like saying a man in handcuffs cannot possibly ever use his hands again. A condition that is necessary under one description is being turned into a permanent necessity in reality.",
  "False attribution":
    "That's like forging a mechanic's signature on a repair report and then treating the fake signature as proof the brakes are fine. Borrowed credibility from the wrong source is being treated as real support.",
  "False balance":
    "That's like putting one engineer and one flat-earther on a stage and calling the shape of the Earth a 50-50 debate. Equal airtime is being mistaken for equal evidence.",
  "False compromise":
    "That's like saying one person calls noon and another calls midnight, so the truth must be six p.m. The midpoint is being treated as correct just because it sits between extremes.",
  "False surrender":
    "That's like stopping the chess clock the moment your queen is hanging and proposing a draw because 'both sides make good points.' A truce is being used to freeze the score when the evidence turns against one side.",
  "Faulty generalization":
    "That's like testing two cracked umbrellas and concluding umbrellas are useless in storms. The evidence does not reach as far as the conclusion claims.",
  "For the sake of argument denial":
    "That's like refusing to discuss whether the parachute would work because you do not believe anyone will jump. A hypothetical premise is being rejected just because it is not believed to be actual.",
  "Four terms fallacy":
    "That's like using one key labeled 'bank' for a riverbank in the first sentence and a money bank in the second, then acting as if the same lock was opened both times. The argument only appears to have three stable terms.",
  "Gambler's fallacy":
    "That's like expecting the roulette wheel to apologize for last night's reds by paying out black today. Independent events do not keep a memory ledger.",
  "Genetic fallacy":
    "That's like rejecting a sound recipe because you first saw it scribbled on a napkin. Origin is being treated as if it settled present merit.",
  "Historian's fallacy":
    "That's like grading a quarterback's split-second throw with the replay paused and the final score already on the screen. Hindsight is being smuggled backward, so people in the past are judged as if they knew what only later observers learned.",
  "Homunculus fallacy":
    "That's like explaining how a security camera works by saying there is a tiny security guard inside the lens watching the feed. The inner observer only repeats the mystery instead of solving it.",
  "Human standard fallacy":
    "That's like declaring a backyard bird no longer wild because a clipboard now labels it 'domestic.' A human tag or rule is being treated as if it changed the underlying fact.",
  "If-by-whiskey":
    "That's like selling the same drink as medicine to one crowd and as rebellion to another, while pretending the label never moved. The key term keeps changing shape to please both sides.",
  "Illicit major":
    "That's like saying all jazz musicians are artists, no plumbers are jazz musicians, therefore no plumbers are artists. The conclusion distributes a term more widely than the premise ever allowed.",
  "Impotent logical space":
    "That's like drawing a target so large that every dart, including the ones that miss the wall, still counts as a bullseye. The claim is framed so nothing could ever count against it.",
  "Incomplete comparison":
    "That's like calling one suitcase lighter than another without saying whether the first suitcase is empty. 'Better' is being claimed without a complete comparison class.",
  "Inconsistent comparison":
    "That's like crowning one athlete the overall best because she outruns one rival, outjumps another, and outswims a third. Different comparison targets are being stitched together to fake an all-around victory.",
  "Intentional fallacy":
    "That's like ignoring the director's note that the stage gun is a prop when the question is what the scene was meant to communicate. Authorial intention is being thrown away even where it matters to interpretation.",
  "Is-ought problem":
    "That's like noticing that many cars speed and concluding therefore the speed limit should be ignored. A fact about what happens is not yet a reason for what ought to happen.",
  "Judgmental language":
    "That's like spray-painting 'idiot' across the whiteboard and calling the graffiti an argument. Loaded language is doing the steering instead of reasons.",
  "Linearity fallacy":
    "That's like assuming two blankets will make soup boil twice as fast. Complex systems often have thresholds, saturation, and diminishing returns rather than neat straight-line gains.",
  "Luddite fallacy":
    "That's like seeing a washing machine in a laundromat and concluding humanity must be doomed to permanent unemployment. Automating one task is being mistaken for wiping out all useful human work.",
  "Masked man fallacy":
    "That's like saying, 'I know my neighbor, but I don't know who is wearing the mascot costume, so the mascot cannot be my neighbor.' Identity cannot always be substituted transparently inside belief or knowledge contexts.",
  "Misleading vividness":
    "That's like letting one shark-attack headline rewrite your picture of the whole ocean. A dramatic case is being treated as if it were representative frequency.",
  "Motte and bailey fallacy":
    "That's like advertising a castle on open ground, then retreating into a small stone bunker whenever anyone attacks it, and later bragging that the castle itself survived the challenge. Defending the safer fallback does not prove the stronger original claim.",
  "Naturalistic fallacy":
    "That's like saying muddy stream water must be healthier than filtered water because it came straight from nature. 'Natural' is being treated as if it settled what is good or preferable.",
  "Negative proof fallacy":
    "That's like treating an unopened drawer as proof that the missing ring is inside because nobody has shown otherwise. Lack of disproof is being mistaken for support.",
  "Nirvana fallacy":
    "That's like refusing a seatbelt because it cannot guarantee survival in every crash. A useful improvement is being rejected for not being perfection.",
  "Overwhelming exception":
    "That's like writing a speed limit that applies except when you're late, tired, celebrating, worried, or in a hurry. The rule is padded with so many exceptions that it stops guiding anything.",
  "Package-deal fallacy":
    "That's like assuming every professor must drink coffee because 'professor' and 'coffee' often travel together in the stereotype. A bundled stereotype is being treated as a necessary package.",
  "Pathetic fallacy":
    "That's like saying the photocopier is sulking today because it jammed. Human feelings are being projected onto impersonal things and then treated as explanation.",
  "Perfect solution fallacy":
    "That's like refusing a mop because it will not make the floor permanently clean forever. A useful remedy is being discarded because it cannot do everything.",
  "Perfect standard":
    "That's like saying a piano recital was a total failure because one note was flat. A messy range of better and worse cases is being collapsed into perfect or worthless.",
  "Perverted analogy":
    "That's like replying to 'budgeting is like dieting' by objecting that budgets do not contain calories. The analogy is being stretched past its point so it can be mocked instead of understood.",
  "Piggy-back assumption":
    "That's like trusting a stranger's whole suitcase because one sticker on it matches the airport code. Evidence for one small detail is being made to ride shotgun for a larger claim it never proved.",
  "Poisoning the well":
    "That's like smearing the witness before the witness even takes the stand and then treating the smear as a rebuttal to whatever the witness will say. The audience is being primed to dismiss the claim before hearing it fairly.",
  "Proof by example":
    "That's like saying one swan on the lake proves every swan on earth is white. A few examples are being asked to do the work of a universal claim.",
  "Proof by verbosity":
    "That's like piling so much wrapping paper on an empty box that people assume there must be something substantial inside. Sheer volume is being mistaken for substance.",
  "Prosecutor's fallacy":
    "That's like hearing that only one key in a thousand fits the lock and concluding the person holding that key must be the burglar. The rarity of a match is being confused with the probability of guilt.",
  "Psychologist's fallacy":
    "That's like saying everyone must hate cilantro for the same inner reason you do. Your own psychology is being projected outward and treated as diagnosis.",
  "Redeeming illogic with evidence":
    "That's like demanding lab tests before rejecting a square circle. Evidence is being demanded to rescue something that is already incoherent on logical grounds.",
  "Regression fallacy":
    "That's like taking credit for calming the ocean because the wave after the huge one was smaller. A move back toward normal is being credited to the intervention without justification.",
  "Reification":
    "That's like blaming inflation as if it were a person hiding in the basement turning knobs. An abstraction is being dressed up as a concrete agent.",
  "Retrospective determinism":
    "That's like insisting the maze had only one obvious path after you've already seen the exit from above. The actual uncertainty beforehand is being erased by the outcome afterward.",
  "Semantic pixelization":
    "That's like forcing every shade between dawn and noon into only 'night' or 'day.' A graded position is being chopped into unrealistically sharp bins.",
  "Sentimental fallacy":
    "That's like treating a moving movie ending as proof the story must be true. Emotional beauty is being mistaken for evidence or feasibility.",
  "Sharpshooter fallacy":
    "That's like firing bullets at a barn and painting the bullseye afterward around the tightest cluster. A hand-picked pattern is being treated as if it were the tested target all along.",
  "Single cause fallacy":
    "That's like explaining a plane crash by naming one loose bolt and ignoring weather, maintenance, training, and fuel. A complex outcome is being squeezed into one cause.",
  "Special pleading":
    "That's like insisting every runner must stay behind the starting line except your cousin, because his dream matters more. An exception is being requested without a relevant reason.",
  "Spotlight fallacy":
    "That's like judging a whole city by the few blocks the news helicopter keeps circling. The most visible cases are being treated as the whole pattern.",
  "Square logic":
    "That's like braiding a knot out of extension cords and calling the tangle a power grid. The argument sounds intricate, but its pieces no longer form a coherent route from premise to conclusion.",
  "Style over substance fallacy":
    "That's like awarding the race to the loudest sports car commercial instead of the fastest car on the track. Presentation polish is being mistaken for argumentative strength.",
  "Suppressed correlative":
    "That's like defining 'clean' so strictly that nothing can count as clean, which makes 'dirty' useless too. One side of a contrast is stretched until the other side cannot function.",
  "Teleological fallacy":
    "That's like saying the river bends because it wants to reach the sea gracefully. A purpose or end-state is being projected where no such built-in goal has been shown.",
  "Thought-terminating cliché":
    "That's like slapping a 'do not inspect' sticker on a cracked engine and calling the repair finished. A stock phrase is being used to shut down inquiry.",
  "Top-down condemnation":
    "That's like catching one salmonella outbreak at a restaurant and concluding restaurants are crooks. A bad individual case is being used to condemn the whole group.",
  "Top-down faulty generalization":
    "That's like rejecting the rule 'rain makes sidewalks wet' because a covered sidewalk stayed dry. A generalization is being attacked by pretending it promised to hold without relevant scope conditions.",
  "Top-down justification":
    "That's like meeting one honest used-car dealer and concluding used-car dealers are unfairly stereotyped. A positive individual case is being made to justify the whole group.",
  "Track-record reset":
    "That's like handing a serial false alarmist a fresh siren every morning and pretending yesterday's empty evacuations tell you nothing. History of failure is being erased when it should affect default confidence.",
  "Two wrongs make a right":
    "That's like keying your neighbor's car because he scratched yours. One wrong is being treated as if it could justify another.",
  "Undistributed middle":
    "That's like saying all fire trucks are red, my bicycle is red, therefore my bicycle is a fire truck. Sharing a broad category does not make two things identical.",
  "Vague insulators":
    "That's like writing a doctor's note that says the machine failed because the patient's vibes were misaligned. Vague elastic terms are being chosen so the claim cannot be pinned down or tested.",
  "Wishful thinking":
    "That's like reading a bank balance through rose-colored glasses and calling the optimism a deposit. What would be comforting if true is being mistaken for what the evidence supports.",
  "Witness chain":
    "That's like saying a rumor is verified because your friend heard it from someone whose cousin heard it from many others you cannot question. Alleged unseen witnesses are being stacked to mimic corroboration.",
  "Wrong causal direction":
    "That's like seeing firefighters at burning buildings and concluding firefighters cause fires. The association is real, but the arrow of causation is pointing the wrong way.",
  };

const analogyFamilyTemplates = {
  "Formal/Structural Fallacy":
    "That's like assembling the right pieces of furniture in the wrong slots and then acting surprised when the table will not stand.",
  "Evidential/Methodological Fallacy":
    "That's like trying to judge the whole case from one folder left open on the corner of the desk while the rest of the evidence stays unread.",
  "Causal/Explanatory Fallacy":
    "That's like seeing two clocks strike together and deciding one must be winding the other.",
  "Statistical/Sampling Fallacy":
    "That's like judging an entire orchard from the first bruised apple in the basket.",
  "Linguistic/Definition Fallacy":
    "That's like changing the dictionary in the middle of the argument and pretending the old word still means the same thing.",
  "Conceptual/Framing Fallacy":
    "That's like forcing every shade on the color wheel into two paint cans and calling the missing shades unreal.",
  "Comparison/Generalization Fallacy":
    "That's like meeting one reckless cyclist and rewriting the traffic code for everyone on two wheels.",
  "Relevance/Distraction Fallacy":
    "That's like answering a question about the map by complaining about the driver's haircut.",
  "Persuasive/Appeal Fallacy":
    "That's like turning up the orchestra because the script cannot carry the scene on its own.",
};
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

const difficultyCategoryProfiles = {
  Formal: 78,
  Mathematical: 74,
  Causal: 56,
  Linguistic: 66,
  Conceptual: 60,
  Evidential: 48,
  Perceptual: 40,
  Perspectival: 62,
  Epistemic: 68,
  Tactical: 44,
  Emotional: 40,
};

const difficultyFamilyModifiers = {
  "Formal/Structural Fallacy": 6,
  "Statistical/Sampling Fallacy": 5,
  "Linguistic/Definition Fallacy": 4,
  "Conceptual/Framing Fallacy": 3,
  "Causal/Explanatory Fallacy": 2,
  "Comparison/Generalization Fallacy": 1,
  "Evidential/Methodological Fallacy": -1,
  "Relevance/Distraction Fallacy": -4,
  "Persuasive/Appeal Fallacy": -6,
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

function renderCaseStudySummary(summary = "") {
  const cleaned = String(summary || "").trim();
  if (!cleaned) return "";

  const sentences = cleaned.split(/(?<=[.!?])\s+/).filter(Boolean);
  if (sentences.length < 3 || cleaned.length < 260) {
    return `<p class="case-summary">${escapeHtml(cleaned)}</p>`;
  }

  const splitIndex = sentences.length >= 4 ? 2 : 1;
  const first = sentences.slice(0, splitIndex).join(" ").trim();
  const second = sentences.slice(splitIndex).join(" ").trim();

  return [first, second]
    .filter(Boolean)
    .map((paragraph) => `<p class="case-summary">${escapeHtml(paragraph)}</p>`)
    .join("");
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
    ${renderCaseStudySummary(study.summary)}
    ${sourceLine ? `<p class="case-source">${sourceLine}</p>` : ""}
  </blockquote>`;
}

function normalizeByteseismicCrossrefs(entries = []) {
  return new Map(
    entries
      .filter((entry) => entry && entry.fallacy)
      .map((entry) => [
        entry.fallacy,
        {
          category: entry.category || "",
          suggestedRefs: Array.isArray(entry.suggestedRefs)
            ? entry.suggestedRefs
                .filter((ref) => ref && ref.title && ref.path)
                .slice(0, 3)
                .map((ref) => ({
                  title: ref.title,
                  path: ref.path,
                  why: ref.why || "",
                }))
            : [],
        },
      ]),
  );
}

function renderByteseismicCrossrefsSection(record, byteseismicCrossrefs) {
  const entry = byteseismicCrossrefs.get(record.name);
  if (!entry?.suggestedRefs?.length) return "";

  return `<section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Related reading on Byteseismic</h3>
          <p class="section-copy">These companion articles widen the philosophical or methodological frame around this fallacy without interrupting the main lesson on this page.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        ${entry.suggestedRefs
          .map((ref) => {
            const href = absoluteByteseismicUrl(ref.path);
            return `<article class="note-panel">
                <p class="eyebrow">Byteseismic</p>
                <h4><a href="${escapeHtml(href)}">${escapeHtml(ref.title)}</a></h4>
                ${ref.why ? `<p class="muted"><strong>Why it helps:</strong> ${escapeHtml(ensureSentence(ref.why))}</p>` : ""}
              </article>`;
          })
          .join("")}
      </div>
    </section>`;
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
  "Motte and bailey fallacy":
    "Do not use this label whenever someone honestly clarifies, narrows, or abandons an overstatement. The fallacy appears only when the weaker fallback is used as cover and the stronger claim is later resumed as if it had been defended.",
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

function weightedDifficultyBase(record) {
  const weights = [0.56, 0.29, 0.15];
  let total = 0;
  let weightTotal = 0;

  record.categories.forEach((category, index) => {
    const profile = difficultyCategoryProfiles[category];
    if (!profile) return;
    const weight = weights[index] || 0.1;
    total += profile * weight;
    weightTotal += weight;
  });

  return weightTotal ? total / weightTotal : 52;
}

function difficultyScoreForRecord(record) {
  const teachingPathSlugs = new Set(
    teachingPathDefinitions
      .filter((path) => path.names.includes(record.name))
      .map((path) => path.slug),
  );

  let score = weightedDifficultyBase(record);
  score += Math.max(0, (record.categories?.length || 0) - 1) * 4;
  score += record.subCategory ? 3 : 0;
  score += record.subSubCategory ? 2 : 0;
  score += difficultyFamilyModifiers[record.family] || 0;

  const notesLength = (record.notes || "").length;
  if (notesLength > 320) score += 5;
  else if (notesLength > 240) score += 3;
  else if (notesLength > 160) score += 1;

  const aliasCount = Array.isArray(record.aliases) ? record.aliases.length : 0;
  if (aliasCount >= 2) score -= 1;
  if (aliasCount >= 4) score -= 1;

  const caseStudyCount = Array.isArray(record.caseStudies) ? record.caseStudies.length : 0;
  if (caseStudyCount >= 4) score -= 1;

  if (featuredNames.includes(record.name)) score -= 2;
  if (foundationalNames.has(record.name)) score -= 8;
  if (teachingPathSlugs.has("start-here")) score -= 6;
  if (teachingPathSlugs.has("public-debate")) score -= 2;
  if (teachingPathSlugs.has("often-confused")) score += 4;

  return clampNumber(Math.round(score), 18, 92);
}

function difficultyForRecord(record) {
  const score = difficultyScoreForRecord(record);
  if (score <= 44) return "Foundational";
  if (score <= 69) return "Intermediate";
  return "Advanced";
}

function classroomLevelForRecord(record, difficulty = "", difficultyScore = null) {
  const score = difficultyScore ?? difficultyScoreForRecord(record);
  if (score <= 38) return "Middle school+";
  if (score <= 56) return "High school";
  if (score <= 74) return "Intro college";
  return "Advanced undergraduate";
}

const pedagogyCache = new Map();
const rhetoricGaugeCache = new Map();

function pedagogyForRecord(record) {
  if (pedagogyCache.has(record.slug)) return pedagogyCache.get(record.slug);

  const difficultyScore = difficultyScoreForRecord(record);
  const difficulty = difficultyForRecord(record);
  const classroomLevel = classroomLevelForRecord(record, difficulty, difficultyScore);
  const domainTag = domainTagForRecord(record);
  const teachingPaths = teachingPathDefinitions
    .filter((path) => path.names.includes(record.name))
    .map((path) => ({ slug: path.slug, title: path.title }));

  const meta = {
    difficultyScore,
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

function difficultyGaugeNarrative(score) {
  if (score >= 82) {
    return "Best taught after students are already comfortable with slower argument reconstruction and more technical distinctions.";
  }
  if (score >= 68) {
    return "Usually easier once readers already have some practice with evidence, framing, or analytic structure.";
  }
  if (score >= 54) {
    return "Teachable at the high school or intro-college level with a bit of scaffolding and comparison.";
  }
  if (score >= 40) {
    return "Usually accessible fairly early once students have a few clear examples in view.";
  }
  return "One of the easier fallacies to introduce in an early lesson.";
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
  const difficultyScore = pedagogy.difficultyScore;
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
      value: clampNumber(Math.round(common), 15, 95),
      lowLabel: "Rare",
      highLabel: "Frequent",
    },
    spot: {
      title: "Easy to spot",
      value: clampNumber(Math.round(spot), 15, 95),
      lowLabel: "Hidden",
      highLabel: "Obvious",
    },
    innocent: {
      title: "Easy to innocently commit",
      value: clampNumber(Math.round(innocent), 15, 95),
      lowLabel: "Low risk",
      highLabel: "Easy slip",
    },
    difficulty: {
      title: "Difficulty",
      value: difficultyScore,
      lowLabel: "Foundational",
      highLabel: "Advanced",
    },
  };

  gauges.common.band = gaugeBandLabel("common", gauges.common.value);
  gauges.common.summary = commonGaugeNarrative(gauges.common.value);
  gauges.spot.band = gaugeBandLabel("spot", gauges.spot.value);
  gauges.spot.summary = spotGaugeNarrative(gauges.spot.value);
  gauges.innocent.band = gaugeBandLabel("innocent", gauges.innocent.value);
  gauges.innocent.summary = innocentGaugeNarrative(gauges.innocent.value);
  gauges.difficulty.band = difficulty;
  gauges.difficulty.summary = difficultyGaugeNarrative(gauges.difficulty.value);
  gauges.difficulty.extraMarkup = `<div class="teaching-pill-row gauge-pill-row">
      ${pedagogy.classroomTags.map((tag) => `<span class="teaching-pill">${escapeHtml(tag)}</span>`).join("")}
    </div>`;

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

function absoluteByteseismicUrl(relativePath = "") {
  return new URL(relativePath, byteseismicSiteUrl).toString();
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
    { href: `${prefix}map/`, label: "Map", key: "map" },
    { href: `${prefix}features/`, label: featuresSectionLabel, key: "features" },
    { href: `${prefix}check-yourself/`, label: "Check Yourself", key: "check-yourself" },
    { href: `${prefix}assessment/`, label: "Assessment", key: "assessment" },
    { href: `${prefix}prompts/`, label: "Fun AI Prompts", key: "prompts" },
    { href: `${prefix}theory/`, label: "Theory", key: "theory" },
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
            <form class="site-search-form" role="search" action="${prefix}fallacies/" method="get">
              <input
                class="site-search-input"
                type="search"
                name="q"
                placeholder="Search fallacies..."
                aria-label="Search fallacies"
                data-site-search-input
              />
              <button class="site-search-button" type="submit">Search</button>
            </form>
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
    ${gauge.extraMarkup || ""}
  </article>`;
}

function renderRhetoricGaugeSection(record, prefix) {
  const gauges = rhetoricGaugesForRecord(record);
  return `<div class="section-block gauge-section">
    <p class="detail-card-label">Teaching gauges</p>
    <p class="muted gauge-note">These 0-100 gauges are teaching aids for comparing fallacies. They are editorial classroom estimates, not measured statistics. <a class="text-link" href="${prefix}map/">View these on the Map</a>.</p>
    <div class="gauge-grid">
      ${renderRhetoricGaugeCard(gauges.common)}
      ${renderRhetoricGaugeCard(gauges.spot)}
      ${renderRhetoricGaugeCard(gauges.innocent)}
      ${renderRhetoricGaugeCard(gauges.difficulty)}
    </div>
  </div>`;
}

function familyDescriptionForName(family) {
  return familyDescriptions[family] || "This family groups fallacies by the main kind of reasoning mistake they make.";
}

function familyPathForName(family) {
  return `families/${slugify(family)}/`;
}

function familyPromptForName(family) {
  return familyDiagnosticPrompts[family] || "Ask what kind of reasoning mistake is doing most of the damage.";
}

function buildFamilyProfiles(records) {
  const counts = records.reduce((map, record) => {
    const family = record.family || "Unspecified";
    map.set(family, (map.get(family) || 0) + 1);
    return map;
  }, new Map());

  return Object.entries(familyDescriptions).map(([name, description]) => ({
    name,
    description,
    slug: slugify(name),
    count: counts.get(name) || 0,
  }));
}

function familyLabelHtml(family = "") {
  return String(family || "")
    .split("/")
    .map((part) => escapeHtml(part))
    .join('<span class="family-slash">/</span><wbr>');
}

function renderFamilyPanel(record, prefix = "../") {
  const family = record.family || "Unspecified";
  const familyHref = familyDescriptions[family] ? `${prefix}${familyPathForName(family)}` : "";
  const familyLabel = familyHref
    ? `<a class="family-anchor" href="${familyHref}">${familyLabelHtml(family)}</a>`
    : familyLabelHtml(family);
  return `<div class="note-panel">
      <h4>Family</h4>
      <p class="muted"><strong class="family-name">${familyLabel}</strong></p>
      <p class="family-note">${escapeHtml(familyDescriptionForName(family))}</p>
      ${
        familyHref
          ? `<p class="family-link-row"><a class="text-link" href="${familyHref}">View all members of this family</a></p>`
          : ""
      }
    </div>`;
}

function renderFamilyGuide(familyProfiles, familyBasePath = "../families/") {
  const cards = familyProfiles.map(
    (profile) => `<a class="note-panel family-guide-card family-guide-card-link" href="${familyBasePath}${escapeHtml(profile.slug)}/">
        <div class="family-guide-top">
          <h4 class="family-heading">${familyLabelHtml(profile.name)}</h4>
          <span class="family-guide-count">${profile.count}</span>
        </div>
        <p class="muted">${escapeHtml(profile.description)}</p>
        <p class="family-guide-link-label">View all ${profile.count} members</p>
      </a>`,
  );

  return `<section class="panel family-guide-panel">
      <div class="section-header">
        <div>
          <h2 class="section-title">Family guide</h2>
          <p class="section-copy">Families are the broad one-home groupings. Categories are narrower diagnostic tags, so a fallacy can belong to several categories but only one family.</p>
        </div>
      </div>
      <div class="family-guide-grid">
        ${cards.join("")}
      </div>
    </section>`;
}

function renderFallacyCard(record, prefix, options = {}) {
  const pedagogy = pedagogyForRecord(record);
  const aliases = record.aliases.join(" ");
  const caseStudyText = record.caseStudies.map((item) => normalizeCaseStudy(item).summary).join(" ");
  const body = `${record.definition} ${record.example} ${record.notes} ${caseStudyText} ${pedagogy.classroomTags.join(" ")} ${pedagogy.teachingPaths.map((item) => item.title).join(" ")}`;
  const posterAsset = options.posterAssets ? resolvePosterAssetForRecord(record, options.posterAssets) : null;
  const showPoster = Boolean(options.showPoster && posterAsset);
  const posterMarkup = showPoster
    ? `
    <div class="fallacy-card-poster-shell">
      <img
        class="fallacy-card-poster"
        src="${prefix}assets/${posterAsset}"
        alt="${escapeHtml(posterAltTextForRecord(record))}"
        loading="lazy"
      />
    </div>`
    : "";
  return `<article
    class="fallacy-card${showPoster ? " fallacy-card-with-poster" : ""}"
    data-fallacy-card
    data-name="${escapeHtml(record.name)}"
    data-aliases="${escapeHtml(aliases)}"
    data-categories="${escapeHtml(record.categories.join("|"))}"
    data-difficulty="${escapeHtml(pedagogy.difficulty)}"
    data-classroom="${escapeHtml(pedagogy.classroomLevel)}"
    data-body="${escapeHtml(body)}"
  >
    <div class="fallacy-card-main">
      <h3><a href="${prefix}fallacies/${record.slug}/">${escapeHtml(record.name)}</a></h3>
      <p class="card-copy">${escapeHtml(truncate(record.definition, 170))}</p>
      ${renderPills(record.categories)}
      ${renderTeacherPills(record)}
    </div>
    ${posterMarkup}
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
  const override = posterCaptionOverrides[record.slug];
  if (override) return override;

  const example = stripTrailingPunctuation(record.example || "");
  const definition = stripTrailingPunctuation(
    String(record.definition || "")
      .replace(/^Occurs when\s+/i, "")
      .replace(/^This fallacy occurs when\s+/i, ""),
  );

  const parts = [];

  if (example) {
    parts.push(`${example}.`);
  }

  if (definition) {
    parts.push(`The key point is that ${lowercaseFirst(definition)}.`);
  }

  return parts.join(" ");
}

function analogyClaimForRecord(record) {
  return ensureSentence(record.example || record.definition || record.name);
}

function analogyResponseForRecord(record) {
  const override = auditedAnalogyResponseOverrides[record.name] || analogyResponseOverrides[record.name];
  if (override) return ensureSentence(override);

  const opening =
    analogyFamilyTemplates[record.family] ||
    "That's like mistaking the echo in the room for a new piece of evidence.";
  const definition = lowerFirst(definitionCore(record.definition || ""));
  const bridge =
    record.family === "Formal/Structural Fallacy"
      ? `The problem is structural: it treats ${definition} as if the conclusion automatically followed.`
      : record.family === "Evidential/Methodological Fallacy"
        ? `The problem is not just thin support, but treating ${definition} as if the evidence had already done enough work.`
        : record.family === "Causal/Explanatory Fallacy"
          ? `The problem is treating ${definition} as if the missing causal link had already been shown.`
          : record.family === "Statistical/Sampling Fallacy"
            ? `The problem is treating ${definition} as if one sample, rate, or pattern settled the larger question.`
            : record.family === "Linguistic/Definition Fallacy"
              ? `The problem is treating ${definition} as if a shift in wording counted as proof.`
              : record.family === "Conceptual/Framing Fallacy"
                ? `The problem is treating ${definition} as if a bad frame or bad category could replace careful distinctions.`
                : record.family === "Comparison/Generalization Fallacy"
                  ? `The problem is treating ${definition} as if one comparison or one slice of experience could stand in for the whole pattern.`
                  : record.family === "Relevance/Distraction Fallacy"
                    ? `The problem is treating ${definition} as if a nearby distraction had actually answered the original issue.`
                    : record.family === "Persuasive/Appeal Fallacy"
                      ? `The problem is treating ${definition} as if pressure, status, or feeling were evidence.`
                      : `The problem is treating ${definition} as if that alone were enough to carry the conclusion.`;
  return ensureSentence(`${opening} ${bridge}`);
}

function renderAnalogyRebuttalSection(record) {
  const claim = analogyClaimForRecord(record);
  const response = analogyResponseForRecord(record);

  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">That&apos;s like saying...</h3>
        <p class="section-copy">Instead of leading with the label, this analogy answers the shape of the reasoning move directly so the mistake is easier to see in plain language.</p>
      </div>
    </div>
    <div class="analogy-grid">
      <article class="detail-section analogy-card analogy-claim-card">
        <p class="analogy-label">Fallacious claim</p>
        <p class="analogy-text">${escapeHtml(claim)}</p>
      </article>
      <article class="detail-section analogy-card analogy-response-card">
        <p class="analogy-label">That&apos;s like saying...</p>
        <p class="analogy-text">${escapeHtml(response)}</p>
      </article>
    </div>
  </section>`;
}

function renderReferenceMeta(record, prompts) {
  return `<div class="meta-grid reference-meta-grid">
    ${renderFamilyPanel(record, "../../")}
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

function dialogueAssessmentIndexSeoDescription() {
  return "Take a difficult 10-item dialogue assessment that distinguishes left vs right and formal vs informal fallacies, with balanced answer types and no-fallacy control cases.";
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

function theorySeoDescription() {
  return "Read theory articles on how to teach, rebut, compare, and explain logical fallacies without flattening them into slogans.";
}

function theoryArticleSeoDescription(article) {
  return article.description;
}

function featuresSeoDescription() {
  return "Read Fallacy Detective cases that use current headlines and public rhetoric to show how logical fallacies appear in live media language.";
}

function featureArticleSeoDescription(article) {
  return article.description;
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

function renderTheoryArticleCard(article, prefix) {
  return `<article class="category-card theory-article-card">
    <h3><a href="${prefix}theory/${article.slug}/">${escapeHtml(article.title)}</a></h3>
    <p class="card-copy">${escapeHtml(article.description)}</p>
    <p class="theory-article-intro">${escapeHtml(article.intro)}</p>
    <p class="assessment-card-link"><a class="inline-link" href="${prefix}theory/${article.slug}/">Read article</a></p>
  </article>`;
}

function renderFeatureArticleCard(article, prefix) {
  return `<article class="category-card theory-article-card">
    <h3><a href="${prefix}features/${article.slug}/">${escapeHtml(article.title)}</a></h3>
    <p class="feature-card-date">${escapeHtml(article.date || "")}</p>
    <p class="card-copy">${escapeHtml(article.description)}</p>
    <p class="theory-article-intro">${escapeHtml(article.intro)}</p>
    <p class="assessment-card-link"><a class="inline-link" href="${prefix}features/${article.slug}/">Open case</a></p>
  </article>`;
}

function theoryInternalLink(pathname, label) {
  return `<a class="inline-link" href="../../${pathname}">${escapeHtml(label)}</a>`;
}

function theoryExternalLink(url, label) {
  return `<a class="inline-link" href="${escapeHtml(url)}">${escapeHtml(label)}</a>`;
}

function renderTheoryPanels(items, layoutClass = "two-column compact-columns") {
  return `<div class="${layoutClass}">
    ${items
      .map(
        (item) => `<article class="note-panel ${item.cardClass || ""}">
          <h4>${escapeHtml(item.title)}</h4>
          <p class="muted">${item.html}</p>
          ${item.extraHtml || ""}
        </article>`,
      )
      .join("")}
  </div>`;
}

function renderTheorySection(section) {
  const layoutClass =
    section.layout === "grid"
      ? "category-grid theory-family-grid"
      : section.layout === "prompt-grid"
        ? "prompt-grid two-column compact-columns"
        : "two-column compact-columns";
  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">${escapeHtml(section.title)}</h3>
        <p class="section-copy">${escapeHtml(section.copy)}</p>
      </div>
    </div>
    ${section.callout ? `<div class="detail-section theory-callout"><p class="theory-formula">${section.callout}</p></div>` : ""}
    ${section.bodyHtml || renderTheoryPanels(section.items || [], layoutClass)}
  </section>`;
}

function renderTheoryReferencesSection(references) {
  return `<section class="section-block">
    <div class="section-header">
      <div>
        <h3 class="section-title">References and further reading</h3>
        <p class="section-copy">Sources that ground the article or push the discussion further.</p>
      </div>
    </div>
    ${renderTheoryPanels(
      references.map((reference) => ({
        title: reference.title,
        html: `${theoryExternalLink(reference.url, reference.linkLabel || reference.title)}${reference.note ? ` — ${reference.note}` : ""}`,
      })),
      "two-column compact-columns",
    )}
  </section>`;
}

function buildStructuredTheoryArticleContent(article, spec) {
  return `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Theory</a><span>/</span><strong>${escapeHtml(article.title)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Theory article</p>
      <h2 class="detail-title">${escapeHtml(article.title)}</h2>
      <p class="detail-deck">${spec.deck}</p>
      ${spec.introPanels ? renderTheoryPanels(spec.introPanels, "two-column compact-columns section-block") : ""}
    </section>

    ${spec.sections.map((section) => renderTheorySection(section)).join("")}

    <section class="detail-section section-block">
      <p class="eyebrow">Takeaway</p>
      <h3 class="section-title">${escapeHtml(spec.takeaway.title)}</h3>
      <p class="section-copy">${spec.takeaway.html}</p>
    </section>

    ${renderTheoryReferencesSection(spec.references)}
  `;
}

function buildTheoryIndexPage() {
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Theory</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Theory</p>
      <h2 class="detail-title">Related articles on how to teach, compare, and rebut fallacies well.</h2>
      <p class="detail-deck">
        This section collects longer-form pieces that sit behind the LogFall reference pages. The goal is to explain not just what the fallacies are,
        but how to use them responsibly, how to rebut them clearly, and how to avoid turning fallacy language into a blunt rhetorical weapon.
      </p>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Current articles</h3>
          <p class="section-copy">Use these articles for the theory behind analogy-based rebuttals, classroom process, AI-assisted classroom analysis, and fallacy-centered critical thinking instruction.</p>
        </div>
      </div>
      <div class="category-grid theory-article-grid">
        ${theoryArticleDefinitions.map((article) => renderTheoryArticleCard(article, "../")).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: "Theory Articles on Logical Fallacies, Rebuttals, and Pedagogy | LogFall",
    description: theorySeoDescription(),
    prefix: "../",
    currentSection: "theory",
    canonicalPath: "theory/",
    keywords: [
      "logical fallacy theory",
      "fallacy pedagogy",
      "rebutting fallacies",
      "critical thinking articles",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Theory", path: "theory/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Theory",
        url: absoluteUrl("theory/"),
        description: theorySeoDescription(),
        publisher: publisherSchema(),
        hasPart: theoryArticleDefinitions.map((article) => ({
          "@type": "Article",
          headline: article.title,
          url: absoluteUrl(`theory/${article.slug}/`),
          description: article.description,
          author: {
            "@type": "Person",
            name: "Phil Stilwell",
          },
        })),
      },
      learningResourceSchema({
        name: "Theory Articles on Logical Fallacies, Rebuttals, and Pedagogy",
        path: "theory/",
        description: theorySeoDescription(),
        about: ["logical fallacies", "critical thinking pedagogy", "argument rebuttal"],
        teaches: ["how to rebut fallacies", "how to teach fallacies responsibly"],
        learningResourceType: ["Article hub", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacy theory", "fallacy pedagogy", "rebutting fallacies"],
      }),
    ],
    content,
  });
}

function buildFeaturesIndexPage() {
  const orderedFeatures = [...featureArticleDefinitions].sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>${featuresSectionLabel}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">${featuresSectionLabel}</p>
      <h2 class="detail-title">A running index of current headline cases, with the newest investigation first.</h2>
      <p class="detail-deck">
        These Fallacy Detective pages update each week and use live headlines, political language, and public argument to show how readers can slide into fallacious interpretations.
        The goal is not to shout “fallacy” at every headline. It is to slow the reading process down, let visitors try the diagnosis first, notice how their political sympathies pull them toward or away from certain labels, and then reveal the most defensible fallacy lenses.
      </p>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Running index</h3>
          <p class="section-copy">The most recent case appears first. A new case is added each week. Each case starts with a real headline, invites the reader to spot the fallacies before any labels appear, and asks the reader to notice where agreement, resentment, or tribal loyalty may be steering the diagnosis.</p>
        </div>
      </div>
      <div class="category-grid theory-article-grid">
        ${orderedFeatures.map((article) => renderFeatureArticleCard(article, "../")).join("")}
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">How new features are chosen</h3>
          <p class="section-copy">Each week, recent headlines are screened for compressed causal wording, false binaries, and other patterns that can invite overconfident reasoning. The point is not just to catch bad arguments in public. It is also to build the habit of noticing when our own political commitments make us eager to call something a fallacy or reluctant to do so.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <article class="note-panel">
          <h4>What gets flagged</h4>
          <p class="muted">The strongest candidates are usually short, vivid headlines that compress a legal, political, or social dispute into one apparent cause, one stark opposition, or one decisive proof claim.</p>
        </article>
        <article class="note-panel">
          <h4>Why self-control still matters</h4>
          <p class="muted">A strong feature page has to distinguish between a fallacy in the headline itself, a fallacy the headline tempts readers to infer, and a fallacy that a partisan reader is simply too eager or too resistant to see because the issue already feels settled.</p>
        </article>
      </div>
    </section>
  `;

  return pageShell({
    title: `${featuresSectionLabel}: Current Headlines and Logical Fallacies | LogFall`,
    description: featuresSeoDescription(),
    prefix: "../",
    currentSection: "features",
    canonicalPath: "features/",
    keywords: [
      "logical fallacies in headlines",
      "headline fallacies",
      "news headline analysis",
      "current rhetoric analysis",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: featuresSectionLabel, path: "features/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: featuresSectionLabel,
        url: absoluteUrl("features/"),
        description: featuresSeoDescription(),
        publisher: publisherSchema(),
        hasPart: orderedFeatures.map((article) => ({
          "@type": "Article",
          headline: article.title,
          url: absoluteUrl(`features/${article.slug}/`),
          description: article.description,
          author: {
            "@type": "Person",
            name: "Phil Stilwell",
          },
        })),
      },
      learningResourceSchema({
        name: `${featuresSectionLabel}: Current Headlines and Logical Fallacies`,
        path: "features/",
        description: featuresSeoDescription(),
        about: ["logical fallacies", "news headlines", "media rhetoric"],
        teaches: ["how to analyze headline framing", "how to resist headline-driven fallacies"],
        learningResourceType: ["Feature hub", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacies in headlines", "headline fallacies", "news rhetoric analysis"],
      }),
    ],
    content,
  });
}

function buildRebuttalsTheoryArticleContent(article) {
  return `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Theory</a><span>/</span><strong>${escapeHtml(article.title)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Theory article</p>
      <h2 class="detail-title">${escapeHtml(article.title)}</h2>
      <p class="detail-deck">
        A fallacy rebuttal often becomes clearer when it answers the bad reasoning with a sharp analogy before it reaches for a technical label.
        This article explains why that approach can be pedagogically stronger, when it works best, and how to do it without becoming glib or unfair.
      </p>
      <div class="two-column compact-columns section-block">
        <div class="note-panel">
          <h4>What this article is for</h4>
          <p class="muted">
            The aim is to help teachers, students, and careful readers rebut a reasoning mistake in plain language. A good analogy can surface the
            structure of the misstep immediately, especially for readers who do not yet know the vocabulary of formal logic or argument analysis.
          </p>
        </div>
        <div class="note-panel">
          <h4>What this article is not saying</h4>
          <p class="muted">
            This is not a rejection of fallacy names. Technical labels are still useful for indexing, comparing, and teaching. The claim is narrower:
            in many live conversations, an analogy-first rebuttal clarifies the mistake more effectively than a label-first reply.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Why not lead with the label?</h3>
          <p class="section-copy">Names can help, but they also come with common pedagogical costs.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Labels can trigger defensiveness</h4>
          <p class="muted">
            When someone hears, “That’s a straw man” or “That’s ad hominem,” they often process the reply as a status move rather than as an explanation.
            The discussion can harden around who is scoring points instead of what reasoning step actually failed.
          </p>
        </div>
        <div class="note-panel">
          <h4>Analogy makes the structure visible</h4>
          <p class="muted">
            A good rebuttal-by-analogy shows the same inferential shape in a cleaner setting. Once the bad move is seen in a simpler parallel case,
            the original argument often becomes easier to diagnose without further jargon.
          </p>
        </div>
        <div class="note-panel">
          <h4>It teaches transfer, not just vocabulary</h4>
          <p class="muted">
            Students do not truly understand a fallacy when they can merely recite its name. They understand it when they can recognize the same move
            across different subject matters, tones, and political loyalties.
          </p>
        </div>
        <div class="note-panel">
          <h4>It leaves room for the label later</h4>
          <p class="muted">
            Analogy-first does not mean analogy-only. After the reasoning slip is clear, the technical label can still be introduced as a compact way to
            file, compare, and revisit the pattern.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">What a strong rebuttal-by-analogy should do</h3>
          <p class="section-copy">The analogy has to illuminate the reasoning, not merely mock it.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Preserve the logical shape</h4>
          <p class="muted">
            The analogy should mirror the inferential structure of the original move. If the mistake is a leap from sequence to causation, the analogy
            should also expose a sequence-to-causation leap rather than some other weakness.
          </p>
        </div>
        <div class="note-panel">
          <h4>Strip away irrelevant heat</h4>
          <p class="muted">
            A good analogy removes political, moral, tribal, or emotionally loaded framing that might be masking the slip. It turns a contentious case
            into one where the reasoning move can be inspected more calmly.
          </p>
        </div>
        <div class="note-panel">
          <h4>Be concrete enough to bite</h4>
          <p class="muted">
            The best analogies are vivid and memorable without becoming cartoonish. They make the error feel unmistakable, not merely comparable in some
            abstract way.
          </p>
        </div>
        <div class="note-panel">
          <h4>Leave the door open to repair</h4>
          <p class="muted">
            The point is not merely to expose absurdity. A useful rebuttal also helps the speaker see what a fairer, narrower, or better-supported version
            of the claim would have to look like.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">How to build one</h3>
          <p class="section-copy">A repeatable way to turn a fallacy diagnosis into a clarifying analogy.</p>
        </div>
      </div>
      <div class="detail-section theory-callout">
        <p class="theory-formula">
          <strong>Template:</strong> Identify the exact reasoning move, rebuild that same move in a cleaner setting, and then state the point of failure without yet naming the fallacy.
        </p>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>1. Isolate the exact slip</h4>
          <p class="muted">
            Do not start with “This sounds bad.” Start with a precise diagnosis: Is the argument confusing correlation with causation? Smuggling the conclusion
            into its own premises? Treating a small sample as decisive? The more exact the diagnosis, the better the analogy will be.
          </p>
        </div>
        <div class="note-panel">
          <h4>2. Translate the pattern</h4>
          <p class="muted">
            Move the same inferential structure into a more neutral setting: kitchens, maps, thermometers, traffic, libraries, scoreboards, and medical tests
            all work well because they make structure visible without the original ideological baggage.
          </p>
        </div>
        <div class="note-panel">
          <h4>3. Keep the pressure on the logic</h4>
          <p class="muted">
            The analogy should not rely on humiliation, sneering, or exaggerated stupidity. Its force should come from showing that the very same reasoning
            would look obviously weak in the parallel case.
          </p>
        </div>
        <div class="note-panel">
          <h4>4. Point toward the repair</h4>
          <p class="muted">
            Once the analogy has done its work, the next step is often to say what kind of evidence, qualification, or fairer framing would be needed to make
            the original argument stronger.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Family-by-family patterns</h3>
          <p class="section-copy">Different families call for different kinds of analogies.</p>
        </div>
      </div>
      <div class="category-grid theory-family-grid">
        ${Object.entries(familyDescriptions)
          .map(
            ([family, description]) => `<article class="note-panel theory-family-card">
              <h4 class="family-heading">${familyLabelHtml(family)}</h4>
              <p class="muted">${escapeHtml(description)}</p>
              <p class="muted"><strong>Analogy strategy:</strong> ${
                family === "Formal/Structural Fallacy"
                  ? "Show a machine, proof, route, or assembly whose parts are in the wrong order or whose structure cannot carry the claimed result."
                  : family === "Evidential/Methodological Fallacy"
                    ? "Use archives, tests, scoreboards, or sampling scenes that make missing, selected, or over-read evidence visible."
                    : family === "Causal/Explanatory Fallacy"
                      ? "Use timelines, switches, dominoes, clocks, and mechanisms so the missing causal link stands out."
                      : family === "Statistical/Sampling Fallacy"
                        ? "Use baskets, polls, classrooms, and probability settings where the sample or rate can be clearly seen as too thin or distorted."
                        : family === "Linguistic/Definition Fallacy"
                          ? "Use dictionaries, contracts, rules, and game instructions to show how a verbal shift quietly changes the terms."
                          : family === "Conceptual/Framing Fallacy"
                            ? "Use maps, categories, shelves, and color wheels to show how a bad frame erases live distinctions or options."
                            : family === "Comparison/Generalization Fallacy"
                              ? "Use orchards, classrooms, traffic, and everyday stereotypes to show how one case is being stretched into too much."
                              : family === "Relevance/Distraction Fallacy"
                                ? "Use emergencies, interviews, and direct questions where a diversion is obviously not an answer."
                                : "Use theater, volume knobs, applause, titles, and emotional atmospheres to show rhetoric taking the place of support."
              }</p>
            </article>`,
          )
          .join("")}
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Risks and limits</h3>
          <p class="section-copy">The method is powerful, but it can also be misused.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Do not oversimplify the case</h4>
          <p class="muted">
            A rebuttal analogy can become unfair if it quietly removes contextual features that really matter. The cleaner case should preserve the logical
            structure while discarding only what is irrelevant.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not confuse ridicule with clarity</h4>
          <p class="muted">
            Some analogies get a laugh but do not actually map the reasoning well. If the analogy humiliates the speaker without accurately tracking the
            inferential mistake, it teaches aggression more than analysis.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not skip the repair</h4>
          <p class="muted">
            If the analogy only says “this is silly,” it leaves the audience without a better model. The strongest teaching move is to expose the mistake and
            then show what the argument would need in order to become stronger.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not retire the technical vocabulary</h4>
          <p class="muted">
            Names still matter for indexing, searching, and cumulative learning. The point is sequence: very often the understanding should come first and the
            label should come second.
          </p>
        </div>
      </div>
    </section>

    <section class="detail-section section-block">
      <p class="eyebrow">Takeaway</p>
      <h3 class="section-title">Use the analogy to open the door, then use the label to organize the lesson.</h3>
      <p class="section-copy">
        A fallacy name is useful when it compresses prior understanding. It is less useful when it substitutes for explanation. The practical rule is simple:
        if the audience cannot yet see the mistake, start with the analogy. If they can already see it, the label can help them file it, compare it, and
        remember it.
      </p>
    </section>

    ${renderTheoryReferencesSection([
      { ...theorySourceCatalog.analogySep, note: "The core philosophical reference on analogical reasoning and its evaluation." },
      { ...theorySourceCatalog.argumentSep, note: "Useful on analogical argument as a general mode of support." },
      { ...theorySourceCatalog.fallaciesSep, note: "Helpful on modern fallacy theory and the place of informal analysis." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "Useful for moving from label-first criticism to question-guided evaluation." },
    ])}
  `;
}

function buildTeachingCurriculumTheoryArticleContent(article) {
  return `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Theory</a><span>/</span><strong>${escapeHtml(article.title)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Theory article</p>
      <h2 class="detail-title">${escapeHtml(article.title)}</h2>
      <p class="detail-deck">
        Logical fallacies are best taught as one strand within a broader critical thinking course, not as a bag of labels for humiliating opponents.
        This article offers a classroom process, a general curriculum, and a set of teaching habits that help students move from naming fallacies
        to actually reasoning better.
      </p>
      <div class="two-column compact-columns section-block">
        <div class="note-panel">
          <h4>What this article is for</h4>
          <p class="muted">
            The aim is to give instructors a practical way to teach fallacies so that students learn diagnosis, comparison, repair, and self-correction.
            The recommendations work for high school, introductory college, and discussion-based adult education with only modest adjustment.
          </p>
        </div>
        <div class="note-panel">
          <h4>What this article is guarding against</h4>
          <p class="muted">
            A fallacy unit goes wrong when it becomes a vocabulary drill, a debate-club scoring system, or a list of gotchas detached from evidence,
            probability, and argument structure. The curriculum below is built to prevent that narrowing.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Core teaching goals</h3>
          <p class="section-copy">A strong unit trains more than recognition.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>1. Recognition</h4>
          <p class="muted">
            Students should be able to identify a fallacy in live prose and explain exactly where the reasoning misstep occurs rather than merely naming it.
          </p>
        </div>
        <div class="note-panel">
          <h4>2. Comparison</h4>
          <p class="muted">
            Students should learn to separate near neighbors such as <code>Straw man</code>, <code>Red herring</code>, and <code>Ad hominem</code>, or <code>False analogy</code>, <code>False equivalence</code>, and <code>Faulty generalization</code>.
          </p>
        </div>
        <div class="note-panel">
          <h4>3. Repair</h4>
          <p class="muted">
            Students should be able to rewrite a weak argument into a stronger one by narrowing the claim, adding needed evidence, or removing the defective inference.
          </p>
        </div>
        <div class="note-panel">
          <h4>4. Self-audit</h4>
          <p class="muted">
            The deepest goal is inward-facing: students should begin noticing which fallacies they themselves are most tempted to commit under pressure, speed, identity, or emotion.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">The general process</h3>
          <p class="section-copy">A repeatable classroom routine matters more than any one list of examples.</p>
        </div>
      </div>
      <div class="detail-section theory-callout">
        <p class="theory-formula">
          <strong>Best general sequence:</strong> define → exemplify → compare → repair → self-apply.
        </p>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Define</h4>
          <p class="muted">
            Begin with a short definition in plain language and one central question the fallacy raises. Students need an entry point that is conceptually clean before they are buried in examples.
          </p>
        </div>
        <div class="note-panel">
          <h4>Exemplify</h4>
          <p class="muted">
            Use one vivid example and one case study from public rhetoric or ordinary life. The example should be short enough to inspect closely and concrete enough to remember.
          </p>
        </div>
        <div class="note-panel">
          <h4>Compare</h4>
          <p class="muted">
            Do not teach fallacies as isolated cards. Put each one beside the two or three entries students are most likely to confuse it with and ask what the exact split is.
          </p>
        </div>
        <div class="note-panel">
          <h4>Repair</h4>
          <p class="muted">
            Ask students to rewrite the claim so it says only what the evidence or reasoning has earned. This prevents the subject from becoming merely punitive.
          </p>
        </div>
        <div class="note-panel">
          <h4>Self-apply</h4>
          <p class="muted">
            End each lesson by asking where students themselves might commit the fallacy in conversation, writing, or news consumption. This is where the subject begins to change habits rather than just terminology.
          </p>
        </div>
        <div class="note-panel">
          <h4>Revisit</h4>
          <p class="muted">
            Fallacies should reappear across the course rather than vanish after one week. The same mistake looks different in ethics, politics, science, personal conflict, and media analysis.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">A general 8-part curriculum</h3>
          <p class="section-copy">This sequence works well as a unit or as one strand inside a longer critical thinking course.</p>
        </div>
      </div>
      <div class="category-grid theory-family-grid">
        <article class="note-panel theory-family-card">
          <h4>1. Orientation</h4>
          <p class="muted">
            Teach what fallacies are and are not. Emphasize that a fallacy is a reasoning mistake, not a moral stain, and that a bad argument can still defend a true conclusion.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>2. Argument basics</h4>
          <p class="muted">
            Cover claims, premises, conclusions, hidden assumptions, deductive versus inductive support, and the difference between evidence and rhetoric. Students need this structure before labels become useful.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>3. Foundational public-rhetoric fallacies</h4>
          <p class="muted">
            Start with a manageable cluster such as <code>Ad hominem</code>, <code>Straw man</code>, <code>False dilemma</code>, <code>Cherry picking</code>, <code>Red herring</code>, and <code>Tu quoque</code>. These are common, concrete, and easy to revisit.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>4. Evidence and sampling</h4>
          <p class="muted">
            Move into <code>Anecdotal fallacy</code>, <code>Base rate fallacy</code>, <code>Hasty generalization</code>, <code>Survivorship bias</code>, and <code>Spotlight fallacy</code>. Pair these with basic discussion of probability, representativeness, and missing evidence.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>5. Causal reasoning</h4>
          <p class="muted">
            Teach <code>Correlation is not causation</code>, <code>Post hoc ergo propter hoc</code>, <code>Wrong causal direction</code>, <code>Single cause fallacy</code>, and <code>Regression fallacy</code>. These become much clearer when students explicitly map possible alternative explanations.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>6. Language and framing</h4>
          <p class="muted">
            Introduce <code>Equivocation</code>, <code>Contextomy</code>, <code>Thought-terminating cliché</code>, <code>Definist fallacy</code>, and <code>Semantic pixelization</code>. Here students learn how wording itself can carry argumentative distortion.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>7. Formal and structural errors</h4>
          <p class="muted">
            Once students are comfortable with premises and conclusions, add <code>Affirming the consequent</code>, <code>Denying the antecedent</code>, <code>Undistributed middle</code>, and related structural mistakes. These work best with short argument maps rather than long prose.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>8. Integrated application</h4>
          <p class="muted">
            End with real articles, speeches, debates, or student essays. Ask for diagnosis, comparison, repair, and caveat awareness rather than only label matching.
          </p>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">A practical weekly rhythm</h3>
          <p class="section-copy">A simple class structure often works better than elaborate activities.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Opening warm-up</h4>
          <p class="muted">
            Begin with one short claim on the board. Ask students to identify the pressure point in the reasoning before any fallacy names are mentioned.
          </p>
        </div>
        <div class="note-panel">
          <h4>Mini-lesson</h4>
          <p class="muted">
            Teach one primary fallacy and one or two near neighbors. Keep the explanation tight and return quickly to examples.
          </p>
        </div>
        <div class="note-panel">
          <h4>Comparison drill</h4>
          <p class="muted">
            Give students 4 to 6 short passages and ask not only which label fits, but why nearby labels fail. This is usually the real bottleneck in mastery.
          </p>
        </div>
        <div class="note-panel">
          <h4>Repair exercise</h4>
          <p class="muted">
            Require one rewritten claim or paragraph that removes the fallacy while preserving as much of the original point as possible.
          </p>
        </div>
        <div class="note-panel">
          <h4>Reflection</h4>
          <p class="muted">
            End by asking where the fallacy appears in students' own thinking, media habits, or drafting patterns. This keeps the unit from becoming merely outward-facing.
          </p>
        </div>
        <div class="note-panel">
          <h4>Cumulative review</h4>
          <p class="muted">
            Recycle old fallacies continuously. By mid-course, each activity should mix old and new entries so students learn discrimination rather than memorized sequence.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Where fallacies fit in the larger course</h3>
          <p class="section-copy">Fallacies should sit inside a wider rational toolkit.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Connect to cognitive bias</h4>
          <p class="muted">
            Students should see that not every reasoning failure is best described as a fallacy. Some mistakes are better understood as memory bias, salience bias, motivated reasoning, or identity-protective cognition.
          </p>
        </div>
        <div class="note-panel">
          <h4>Connect to probability and statistics</h4>
          <p class="muted">
            Many fallacies only become fully clear when students have some feel for base rates, samples, uncertainty, and causal alternatives. A fallacy curriculum should therefore touch probability and statistics, even at an introductory level.
          </p>
        </div>
        <div class="note-panel">
          <h4>Connect to deductive and inductive structure</h4>
          <p class="muted">
            Students should learn that some fallacies are structural, some evidential, some linguistic, and some rhetorical. That distinction helps them avoid treating every weak argument as the same kind of failure.
          </p>
        </div>
        <div class="note-panel">
          <h4>Connect to argument repair</h4>
          <p class="muted">
            In a mature critical thinking course, diagnosing a fallacy should usually be followed by the question: what would a stronger version of the claim need in order to work?
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Assessment suggestions</h3>
          <p class="section-copy">Test more than label recall.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Identification quizzes</h4>
          <p class="muted">
            Use short mixed sets where the fallacy is not revealed by the page title or grouping. Students should have to diagnose from the claim itself.
          </p>
        </div>
        <div class="note-panel">
          <h4>Near-neighbor explanations</h4>
          <p class="muted">
            Ask why the best label fits better than two close alternatives. This catches shallow recognition very quickly.
          </p>
        </div>
        <div class="note-panel">
          <h4>Repair tasks</h4>
          <p class="muted">
            Grade whether students can produce a fairer, narrower, better-supported version of a flawed argument. This is often a better measure of real understanding than multiple choice alone.
          </p>
        </div>
        <div class="note-panel">
          <h4>Applied analysis</h4>
          <p class="muted">
            Use a short op-ed, advertisement, or debate clip and ask students to quote, classify, justify, and repair. This is the closest analogue to real-world use.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Mistakes teachers should avoid</h3>
          <p class="section-copy">A strong unit prevents common distortions in how fallacies are taught.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Do not teach labels without cases</h4>
          <p class="muted">
            Students remember names poorly when names are taught without live examples, comparisons, and repair exercises.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not reward smugness</h4>
          <p class="muted">
            If students learn that the goal is to catch other people sounding foolish, the subject becomes rhetorically aggressive and intellectually shallow.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not flatten all weak reasoning into fallacy-talk</h4>
          <p class="muted">
            Some weak claims are merely underdeveloped, overstated, vague, or poorly evidenced without fitting a classic fallacy label cleanly.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not isolate the unit from the rest of critical thinking</h4>
          <p class="muted">
            The best teaching keeps fallacy study connected to evidence, probability, argument structure, bias, and intellectual humility.
          </p>
        </div>
      </div>
    </section>

    <section class="detail-section section-block">
      <p class="eyebrow">Takeaway</p>
      <h3 class="section-title">Teach fallacies as habits of diagnosis and repair, not as a museum of labels.</h3>
      <p class="section-copy">
        A good curriculum helps students see recurring reasoning patterns, compare near neighbors, repair weak claims, and turn the tool inward on their own thinking.
        That is how fallacy study becomes part of a genuine critical thinking course rather than a clever side topic.
      </p>
    </section>

    ${renderTheoryReferencesSection([
      { ...theorySourceCatalog.criticalThinkingIep, note: "Strong background on critical thinking as a broader rational practice." },
      { ...theorySourceCatalog.openstaxArguments, note: "Useful on premises, conclusions, and argument structure in an introductory classroom." },
      { ...theorySourceCatalog.openstaxFallacies, note: "Accessible treatment of informal fallacy types for classroom use." },
      { ...theorySourceCatalog.argumentMapsStudy, note: "Helpful if the curriculum includes mapping and structural visualization." },
    ])}
  `;
}

function buildAiGemsTheoryArticleContent(article) {
  const roleAndStancePrompt = escapeHtml(
    [
      "You are a careful logical-fallacy analyst for a critical thinking class.",
      "Your job is not to hunt for labels aggressively, but to identify only the most justified fallacies in the passage.",
      "Prefer specificity over broad accusation.",
      "If a suspected fallacy is weak, borderline, or plausibly explained another way, say so clearly.",
      "Do not moralize. Diagnose the reasoning, quote the relevant wording, explain the dynamics, and propose a fairer response or repair.",
    ].join("\n"),
  );

  const outputSchemaPrompt = escapeHtml(
    [
      "Format every answer with this structure:",
      "",
      "◉ Source Summary",
      "➘ One or two sentences summarizing the article, speech, or debate passage.",
      "",
      "◉ Fallacy Findings",
      "For each fallacy found, use this structure:",
      "➘ Fallacy Name: [most specific label]",
      "➘ Confidence Score (0-4): [number]",
      "➘ Distortion Score (0-4): [number]",
      "➘ Salient Quote: \"[quote only the key lines needed to see the misstep]\"",
      "➘ Why It Fits: [3-5 sentences explaining the reasoning failure]",
      "➘ Caveat: [state what would make this label too strong or misapplied]",
      "➘ LogFall Link: https://logfall.com/fallacies/[slug]/",
      "➘ Response: [answer the misstep in plain language]",
      "➘ Repair: [rewrite the claim in a stronger, fairer form]",
      "",
      "◉ Overall Pattern",
      "➘ Briefly describe what kind of reasoning drift the piece shows overall.",
    ].join("\n"),
  );

  const scoringPrompt = escapeHtml(
    [
      "Scoring rubric:",
      "0 = not present or too weak to justify",
      "1 = faint or merely possible",
      "2 = present but modest",
      "3 = strong and clear",
      "4 = central and unmistakable",
      "",
      "Use two separate scores:",
      "➘ Confidence Score: how sure you are that the label fits",
      "➘ Distortion Score: how much that fallacy affects the passage's overall reasoning",
    ].join("\n"),
  );

  const workflowPrompt = escapeHtml(
    [
      "Passage handling rules:",
      "1. Quote only the lines needed to understand the fallacy.",
      "2. Identify no more than the 3 strongest fallacies unless the user requests a fuller sweep.",
      "3. Compare close alternatives before settling on the final label.",
      "4. Prefer one precise label over several overlapping ones.",
      "5. When no clear fallacy appears, say that directly instead of forcing a diagnosis.",
    ].join("\n"),
  );

  return `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Theory</a><span>/</span><strong>${escapeHtml(article.title)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Theory article</p>
      <h2 class="detail-title">${escapeHtml(article.title)}</h2>
      <p class="detail-deck">
        In the age of AI, one of the best ways to teach logical fallacies is not to have students passively ask a chatbot for answers, but to have them
        collaboratively design a Gemini Gem or other pre-prompted agent that must identify, score, explain, rebut, and repair bad reasoning. When students
        build the analytical frame together, they make their own standards explicit and become much better at seeing both the power and the limits of AI diagnosis.
      </p>
      <div class="two-column compact-columns section-block">
        <div class="note-panel">
          <h4>What this article is for</h4>
          <p class="muted">
            The goal is to help teachers turn AI from a shortcut into a visible thinking tool. The class uses the model as a provisional analyst whose
            output must be judged, revised, and improved by students rather than simply accepted.
          </p>
        </div>
        <div class="note-panel">
          <h4>What this method avoids</h4>
          <p class="muted">
            It avoids the shallow classroom pattern where students paste in a passage, collect labels, and call that critical thinking. The human work here
            lies in building the prompt, auditing the output, tightening the distinctions, and deciding when the model is overreaching.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Why collaborative agent-building works</h3>
          <p class="section-copy">The learning payoff comes from making the criteria public and revisable.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Students externalize their standards</h4>
          <p class="muted">
            The moment students try to write the Gem's instructions, they are forced to decide what counts as good evidence for a fallacy label, what counts
            as a false positive, and how much quotation is needed to make the diagnosis fair.
          </p>
        </div>
        <div class="note-panel">
          <h4>Comparison becomes unavoidable</h4>
          <p class="muted">
            A strong agent prompt has to say how to choose among near neighbors. That means students must clarify the differences among <code>Straw man</code>,
            <code>Red herring</code>, <code>Ad hominem</code>, <code>False equivalence</code>, and other often-confused entries rather than relying on vague familiarity.
          </p>
        </div>
        <div class="note-panel">
          <h4>The model's mistakes become teachable moments</h4>
          <p class="muted">
            When the agent overlabels, misses a caveat, or confuses two fallacies, the class has concrete material to debug. The failure is no longer hidden
            inside a teacher lecture; it becomes a visible reasoning event the whole room can inspect.
          </p>
        </div>
        <div class="note-panel">
          <h4>Students practice rebuttal and repair</h4>
          <p class="muted">
            A good agent does not stop with naming. It also explains the misstep, answers it in plain language, and suggests a stronger rewrite. That keeps
            the exercise pointed toward better reasoning rather than mere accusation.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">A strong classroom process</h3>
          <p class="section-copy">This works as a recurring exercise inside a critical thinking course.</p>
        </div>
      </div>
      <div class="detail-section theory-callout">
        <p class="theory-formula">
          <strong>Best general sequence:</strong> choose a text → build the prompt → run the agent → audit the output → revise the prompt → score and respond.
        </p>
      </div>
      <div class="category-grid theory-family-grid">
        <article class="note-panel theory-family-card">
          <h4>1. Choose a manageable source</h4>
          <p class="muted">
            Use a short editorial, op-ed paragraph, debate exchange, or opening statement. The text should be rich enough to contain real argumentative moves,
            but short enough that the class can still inspect each quoted line carefully.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>2. Build the first prompt as a group</h4>
          <p class="muted">
            Draft the Gem's role, output format, and scoring rules on the board. Ask students what the agent must be prevented from doing, especially in cases
            where weak or ambiguous reasoning might tempt it into overdiagnosis.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>3. Run the agent on the same text</h4>
          <p class="muted">
            Have the class watch one common output rather than scattering immediately into private runs. That shared output gives everyone the same object to critique.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>4. Audit every finding</h4>
          <p class="muted">
            For each fallacy the agent flags, ask: is the quote sufficient, is the label the most precise one, what nearby labels should be ruled out, and what caveat
            should be stated before we accept the diagnosis?
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>5. Revise the prompt</h4>
          <p class="muted">
            Tighten the instructions in response to the model's errors. Over time, the class learns that prompt-writing is really criteria-writing in disguise.
          </p>
        </article>
        <article class="note-panel theory-family-card">
          <h4>6. End with human judgment</h4>
          <p class="muted">
            The final class product should not be "what the model said." It should be a human-vetted set of fallacy diagnoses, scores, rebuttals, and repairs.
          </p>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">What the agent should be asked to do</h3>
          <p class="section-copy">A useful Gem has a narrow mission and a disciplined output.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Identify only the strongest candidates</h4>
          <p class="muted">
            The agent should not carpet-bomb a passage with labels. Ask it to identify only the clearest two or three fallacies unless the user specifically requests a full sweep.
          </p>
        </div>
        <div class="note-panel">
          <h4>Quote enough to make the diagnosis visible</h4>
          <p class="muted">
            Require a short, sufficient quotation for each finding. This keeps the agent from floating free of the text and makes the reasoning misstep inspectable.
          </p>
        </div>
        <div class="note-panel">
          <h4>Explain the dynamics, not just the label</h4>
          <p class="muted">
            A good answer says exactly how the passage moves from evidence to conclusion, where the drift happens, and why that move is too fast, too broad, too selective, or too distracted.
          </p>
        </div>
        <div class="note-panel">
          <h4>Score and respond</h4>
          <p class="muted">
            The agent should assign clear scores and then do something constructive: rebut the misstep in plain language, or repair the claim so it says only what the argument has earned.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Prompt components that produce organized, salient output</h3>
          <p class="section-copy">The best prompts are modular: role, method, format, scoring, and response rules.</p>
        </div>
      </div>
      <div class="prompt-grid two-column compact-columns">
        <article class="note-panel prompt-card">
          <h4>1. Role and stance</h4>
          <p class="muted">
            Tell the model to act like a careful classroom analyst rather than a prosecuting attorney. This single component sharply reduces overlabeling and rhetorical heat.
          </p>
          <pre class="theory-prompt-box">${roleAndStancePrompt}</pre>
        </article>
        <article class="note-panel prompt-card">
          <h4>2. Output schema</h4>
          <p class="muted">
            Strong outputs come from strong formatting constraints. If you want organization, specify headings, quote rules, score labels, and link fields directly.
          </p>
          <pre class="theory-prompt-box">${outputSchemaPrompt}</pre>
        </article>
        <article class="note-panel prompt-card">
          <h4>3. Scoring rules</h4>
          <p class="muted">
            Separate confidence from importance. A model may be very confident that a minor fallacy is present, or only moderately confident that a major one shapes the passage.
          </p>
          <pre class="theory-prompt-box">${scoringPrompt}</pre>
        </article>
        <article class="note-panel prompt-card">
          <h4>4. Passage-handling rules</h4>
          <p class="muted">
            The prompt should say how many fallacies to report, how much to quote, and when to withhold a label. Those guardrails often matter more than the fancy wording at the top.
          </p>
          <pre class="theory-prompt-box">${workflowPrompt}</pre>
        </article>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">A useful scoring model</h3>
          <p class="section-copy">Keep the numbers simple enough for students to compare and debate.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Confidence score</h4>
          <p class="muted">
            This score answers: how well does the passage justify the label? Students can argue over whether the quotation really supports the diagnosis or whether the evidence is too thin.
          </p>
        </div>
        <div class="note-panel">
          <h4>Distortion score</h4>
          <p class="muted">
            This score answers: how much does that fallacy matter to the overall argument? Some fallacies are present but peripheral; others shape the whole reasoning structure.
          </p>
        </div>
        <div class="note-panel">
          <h4>Optional class extension</h4>
          <p class="muted">
            If you want a richer tool, add one more score for <strong>repairability</strong>: how easy is it to salvage the argument without giving up its core point?
          </p>
        </div>
        <div class="note-panel">
          <h4>Keep the human override explicit</h4>
          <p class="muted">
            The class should always be free to lower a score, reject a label, or replace the model's chosen fallacy with a better one. The point is calibration, not obedience.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">What to grade in the classroom</h3>
          <p class="section-copy">Grade the students' reasoning about the AI, not merely the AI's output.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Prompt quality</h4>
          <p class="muted">
            Did the group write clear instructions, good output constraints, and fair caveat rules? A vague prompt usually reveals a vague understanding.
          </p>
        </div>
        <div class="note-panel">
          <h4>Audit quality</h4>
          <p class="muted">
            Did students catch false positives, weak quotations, and sloppy category choices? The audit is often more pedagogically valuable than the first run itself.
          </p>
        </div>
        <div class="note-panel">
          <h4>Rebuttal and repair quality</h4>
          <p class="muted">
            Did the group answer the reasoning mistake clearly and then offer a stronger formulation? This is where the exercise becomes constructive rather than merely classificatory.
          </p>
        </div>
        <div class="note-panel">
          <h4>Reflective accuracy</h4>
          <p class="muted">
            Can students say where the agent helped, where it overreached, and what that reveals about both AI and fallacy instruction? That meta-level reflection is part of the lesson.
          </p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Risks and safeguards</h3>
          <p class="section-copy">AI can sharpen the unit, but only if the classroom remains intellectually in charge.</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <div class="note-panel">
          <h4>Do not let the model become the authority</h4>
          <p class="muted">
            A Gem is a scaffold, not an oracle. Students should always be asked to justify the final judgment independently of the model's wording.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not reward overdiagnosis</h4>
          <p class="muted">
            If students think more labels mean a better answer, the exercise will quickly turn into fallacy inflation. Reward precision, restraint, and clean comparison instead.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not separate fallacies from the wider toolkit</h4>
          <p class="muted">
            Some problems are really about bad statistics, weak causal design, or cognitive bias rather than classic fallacy forms. The agent should be taught to say that when needed.
          </p>
        </div>
        <div class="note-panel">
          <h4>Do not hide the prompt from students</h4>
          <p class="muted">
            The prompt is the curriculum in compressed form. When students can inspect and revise it, they are learning the criteria themselves rather than merely consuming an answer.
          </p>
        </div>
      </div>
    </section>

    <section class="detail-section section-block">
      <p class="eyebrow">Takeaway</p>
      <h3 class="section-title">Treat the Gem as a collaboratively built reasoning instrument, not as a replacement for judgment.</h3>
      <p class="section-copy">
        The strongest classroom use of AI is not passive extraction but collaborative calibration. Students learn logical fallacies more deeply when they must tell
        the agent what to look for, how to quote, how to compare close labels, how to score responsibly, and how to answer a fallacy in plain language once it is found.
      </p>
    </section>

    ${renderTheoryReferencesSection([
      { ...theorySourceCatalog.gemsTips, note: "Google's own guidance on writing clearer, more detailed Gem instructions." },
      { ...theorySourceCatalog.gemsUse, note: "Overview of how Gems operate as repeatable custom instruction agents." },
      { ...theorySourceCatalog.logicalFallacyDetection, note: "Foundational NLP paper on logical fallacy detection as a model task." },
      { ...theorySourceCatalog.explainableFallacyDetection, note: "Useful on explainable, staged fallacy identification in natural-language arguments." },
    ])}
  `;
}

function buildBiasesVsFallaciesTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "The cleanest first distinction is this: a fallacy is usually something you can point to in an argument on the page, while a cognitive bias is usually a tendency in the thinker behind it. They often travel together like bad roommates, but they are not the same thing, and teaching them as if they were the same thing quickly turns a critical thinking class into conceptual soup.",
    introPanels: [
      {
        title: "Quick distinction",
        html: "A fallacy is a defect in reasoning as presented. A bias is a recurring tendency in judgment, attention, memory, or evaluation that can make such defects more likely.",
      },
      {
        title: "Why the confusion happens",
        html: "Biases often help produce fallacies, and fallacies often reveal the fingerprints of bias. But one lives mainly in the structure of the argument, while the other lives mainly in the habits of the mind that produced or accepted it.",
      },
    ],
    sections: [
      {
        title: "A fast rule of thumb",
        copy: "Ask whether you are diagnosing the argument, the thinker, or both.",
        items: [
          {
            title: "Fallacy",
            html: "If you can quote the passage and say, 'Right there, that is where the conclusion outruns the support,' you are usually dealing with a fallacy.",
          },
          {
            title: "Bias",
            html: "If you are explaining why a person was tempted to notice some evidence, ignore other evidence, or cling to a preferred story, you are usually talking about bias.",
          },
          {
            title: "Both at once",
            html: `A writer may show confirmation bias and commit ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")} in the same paragraph. The bias helps explain the selection; the fallacy describes the argumentative result.`,
          },
          {
            title: "Neither term solves everything",
            html: "Some bad reasoning is just careless, vague, poorly sourced, or underdeveloped. Not every weak sentence deserves an honorary medical degree in either fallacies or biases.",
          },
        ],
      },
      {
        title: "Worked contrasts",
        copy: "The best way to see the difference is to compare close-looking cases.",
        layout: "grid",
        items: [
          {
            title: "Confirmation bias vs. Cherry picking",
            html: `Confirmation bias is the tendency to notice and privilege evidence that supports a preferred view. ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")} is the public argumentative move where only the friendly evidence is shown, as if the missing evidence had politely excused itself from the room.`,
          },
          {
            title: "Availability bias vs. Anecdotal fallacy",
            html: `Availability bias makes vivid cases feel more representative than they are. ${theoryInternalLink("fallacies/anecdotal-fallacy/", "Anecdotal fallacy")} happens when that vivid case is then used as if it carried the evidential weight of a broader sample.`,
          },
          {
            title: "Status quo bias vs. Appeal to tradition",
            html: `Status quo bias is a preference for keeping things as they are because change feels risky or costly. ${theoryInternalLink("fallacies/appeal-to-tradition/", "Appeal to tradition")} is the argument that oldness or inherited practice itself counts as support.`,
          },
          {
            title: "Motivated reasoning vs. No True Scotsman",
            html: `Motivated reasoning is the broader habit of protecting a valued identity or conclusion. ${theoryInternalLink("fallacies/no-true-scotsman/", "No True Scotsman")} is one tidy argumentative way of doing that: redefine the category after the counterexample arrives so the preferred claim survives untouched.`,
          },
        ],
      },
      {
        title: "How to teach the distinction",
        copy: "Students usually learn this faster when they are forced to speak in two voices.",
        items: [
          {
            title: "Voice one: argument diagnosis",
            html: "Have students point to the exact sentence where the reasoning fails and explain the failure without mentioning psychology at all.",
          },
          {
            title: "Voice two: bias hypothesis",
            html: "Then ask what background tendency might have made the mistake tempting: salience, identity-protection, status-quo comfort, overconfidence, or something else.",
          },
          {
            title: "Keep the burden of proof separate",
            html: "The fallacy claim should be judged by the text. The bias claim should be treated more cautiously, since students rarely have full access to the writer's inner machinery.",
          },
          {
            title: "End with self-application",
            html: "The class becomes much less smug and much more useful when students ask which biases most often feed the fallacies they themselves commit under speed, irritation, or team loyalty.",
          },
        ],
      },
    ],
    takeaway: {
      title: "Fallacies are usually the visible crack in the argument; biases are often the hidden pressure behind the wall.",
      html: "Teach students to separate the public reasoning failure from the private or background tendency that may have helped produce it. When they can do both, their diagnosis becomes sharper and their humility improves at roughly the same rate.",
    },
    references: [
      { ...theorySourceCatalog.fallaciesSep, note: "Useful on the argument conception of fallacies and on how biases can feed fallacious reasoning." },
      { ...theorySourceCatalog.boundedRationalitySep, note: "Helpful on heuristics, bias traditions, and the standards used to judge human reasoning." },
      { ...theorySourceCatalog.openstaxBiases, note: "Accessible teaching material on cognitive biases and reflective correction." },
      { ...theorySourceCatalog.criticalThinkingIep, note: "Good bridge source on critical thinking, fallacies, and rational evaluation." },
    ],
  });
}

function buildWhenNotToCallFallacyTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "One of the underrated virtues in critical thinking is restraint. If every weak claim becomes a fallacy charge, the vocabulary loses precision, the conversation gets hotter than it gets clearer, and the critic starts sounding like a vending machine that only dispenses the word 'fallacy.' This article is about not doing that.",
    introPanels: [
      {
        title: "The central principle",
        html: "Do not call something a fallacy unless you can quote the relevant move, explain the exact reasoning failure, and say why nearby labels fit less well.",
      },
      {
        title: "Why overlabeling is costly",
        html: "False positives make the critic look lazy, make students distrust the vocabulary, and blur the difference between genuine reasoning failure and mere disagreement.",
      },
    ],
    sections: [
      {
        title: "Four things that are not yet a fallacy",
        copy: "A fallacy label usually needs more than irritation, falsity, or suspicion.",
        items: [
          {
            title: "A false conclusion",
            html: "A claim can be false without the argument for it matching any specific named fallacy. Sometimes it is just wrong, unsupported, or underexplained.",
          },
          {
            title: "A rude tone",
            html: `A speaker may be arrogant, sarcastic, or exhausting while still not committing ${theoryInternalLink("fallacies/ad-hominem/", "Ad hominem")}. Bad manners are not automatically bad logic.`,
          },
          {
            title: "A weak source",
            html: `Questioning a witness, publication, or data source is not automatically fallacious. It becomes a problem only when source talk replaces relevant engagement or pretends to settle what the evidence itself still leaves open.`,
          },
          {
            title: "A claim you simply dislike",
            html: "Disagreement is not diagnosis. The fact that a conclusion strikes you as implausible, politically noxious, or aesthetically hideous does not yet tell you what went wrong in the reasoning.",
          },
        ],
      },
      {
        title: "Common false positives",
        copy: "These are the classroom classics for overreach.",
        layout: "grid",
        items: [
          {
            title: "Not every source challenge is Ad hominem",
            html: `If the source's credibility is directly relevant, then asking whether the source is informed, biased, paid, or methodologically careful may be perfectly fair. ${theoryInternalLink("fallacies/ad-hominem/", "Ad hominem")} begins when the personal fact is doing argumentative work it has not earned.`,
          },
          {
            title: "Not every hard choice is a False dilemma",
            html: `Sometimes the live institutional options really are narrow. ${theoryInternalLink("fallacies/false-dilemma/", "False dilemma")} requires that real alternatives are being erased or hidden, not merely that the situation is unpleasantly constrained.`,
          },
          {
            title: "Not every comparison is a False analogy",
            html: `An analogy is allowed to be imperfect. To refute it, you need a relevant disanalogy, not just a theatrical announcement that two things are not identical because, astonishingly, the universe contains more than one object.`,
          },
          {
            title: "Not every emotional sentence is Appeal to emotion",
            html: `Emotion becomes fallacious when it tries to do the evidential work. Strong feeling can be completely appropriate if it is tied to facts, harms, stakes, or testimony rather than replacing them.`,
          },
        ],
      },
      {
        title: "Questions to ask before you label",
        copy: "These questions slow the critic down in the right way.",
        items: [
          {
            title: "What is the exact conclusion?",
            html: "If you cannot state the conclusion clearly, you probably are not yet ready to classify the failure clearly either.",
          },
          {
            title: "What quotation carries the problem?",
            html: "The charge should be tied to a line, not to a vibe. If the best evidence you have is the general mood of the passage, keep digging.",
          },
          {
            title: "What is the strongest nearby alternative?",
            html: `Ask whether the passage is better classified as ${theoryInternalLink("fallacies/red-herring/", "Red herring")} rather than ${theoryInternalLink("fallacies/straw-man-argument/", "Straw man argument")}, or as ${theoryInternalLink("fallacies/appeal-to-motive/", "Appeal to motive")} rather than ${theoryInternalLink("fallacies/ad-hominem/", "Ad hominem")}.`,
          },
          {
            title: "Could a caveat save the claim?",
            html: "If one missing sentence of qualification would fix the problem, say so. A good critic should know the difference between a wreck and a bent fender.",
          },
        ],
      },
    ],
    takeaway: {
      title: "A fallacy label is a precision tool, not a confetti cannon.",
      html: "Use it when it clarifies a specific reasoning defect, not when it merely expresses impatience. The critic who can withhold a label well usually applies one better too.",
    },
    references: [
      { ...theorySourceCatalog.fallaciesIep, note: "Especially useful for the reminder that a fallacy charge carries a burden of justification." },
      { ...theorySourceCatalog.fallaciesSep, note: "Good on the diversity of fallacy conceptions and why the category is more complex than textbook folklore suggests." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "Helpful on using critical questions and dialogical standards instead of reflex labeling." },
      { ...theorySourceCatalog.openstaxArguments, note: "Useful for students who need a clean distinction between claims, reasons, and conclusions before diagnosis begins." },
    ],
  });
}

function buildNearNeighborsTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "Most students do not fail because they know nothing. They fail because two or three labels feel close enough that they blur together. This article is about sharpening those boundaries. If fallacy study is birdwatching, these are the species that keep getting mistaken for one another through smudged binoculars.",
    introPanels: [
      {
        title: "Why confusion is normal",
        html: "Close fallacies often share topic, tone, or surface symptoms. The split usually appears only when you ask exactly where the support goes off the rails.",
      },
      {
        title: "The classroom trick",
        html: "Always compare at least two nearby alternatives before settling on a label. Students who only learn to match names to examples become confident too early and accurate too late.",
      },
    ],
    sections: [
      {
        title: "The most useful comparison clusters",
        copy: "These are the pairs and trios most worth drilling repeatedly.",
        layout: "grid",
        items: [
          {
            title: "Straw man vs. Red herring",
            html: `${theoryInternalLink("fallacies/straw-man-argument/", "Straw man argument")} misrepresents the opponent's view so it becomes easier to attack. ${theoryInternalLink("fallacies/red-herring/", "Red herring")} changes the subject or shifts the attention away from the real issue.`,
          },
          {
            title: "Ad hominem vs. Appeal to motive",
            html: `${theoryInternalLink("fallacies/ad-hominem/", "Ad hominem")} attacks the person in a way that is supposed to discredit the argument. ${theoryInternalLink("fallacies/appeal-to-motive/", "Appeal to motive")} focuses more specifically on alleged intent or hidden incentives as if that settled the reasoning itself.`,
          },
          {
            title: "False analogy vs. False equivalence",
            html: `${theoryInternalLink("fallacies/false-analogy/", "False analogy")} stretches a comparison beyond what relevant similarities can bear. ${theoryInternalLink("fallacies/false-equivalence/", "False equivalence")} flattens meaningful differences and treats two things as if they carry equal weight, guilt, or evidential standing.`,
          },
          {
            title: "Hasty generalization vs. Anecdotal fallacy",
            html: `${theoryInternalLink("fallacies/hasty-generalization/", "Hasty generalization")} moves from too little evidence to too broad a conclusion. ${theoryInternalLink("fallacies/anecdotal-fallacy/", "Anecdotal fallacy")} relies especially on vivid personal cases or isolated stories as though they could do the work of broader evidence.`,
          },
          {
            title: "Correlation is not causation vs. Post hoc",
            html: `${theoryInternalLink("fallacies/correlation-is-not-causation/", "Correlation is not causation")} warns that co-variation alone is insufficient for causal inference. ${theoryInternalLink("fallacies/post-hoc-ergo-propter-hoc/", "Post hoc ergo propter hoc")} is the more specific leap from sequence in time to causation.`,
          },
          {
            title: "Begging the question vs. Bare assertion",
            html: `${theoryInternalLink("fallacies/begging-the-question/", "Begging the question")} smuggles the conclusion into the support. ${theoryInternalLink("fallacies/bare-assertion-fallacy/", "Bare assertion fallacy")} just states the claim repeatedly or confidently without meaningful support at all.`,
          },
        ],
      },
      {
        title: "Questions that force the exact split",
        copy: "Each cluster has a diagnostic question that does disproportionate work.",
        items: [
          {
            title: "Did the speaker distort the view or merely dodge it?",
            html: "That question splits straw man from red herring surprisingly well.",
          },
          {
            title: "Is the comparison overloaded or is the difference erased?",
            html: "That question helps separate false analogy from false equivalence.",
          },
          {
            title: "Is the evidence too thin, or is it vivid in a misleading way?",
            html: "That question helps distinguish hasty generalization from anecdotal fallacy.",
          },
          {
            title: "Is the leap from timing or from co-occurrence?",
            html: "That question often reveals whether post hoc or correlation-not-causation is the more exact label.",
          },
        ],
      },
      {
        title: "A teaching habit worth keeping",
        copy: "Never let the winning label walk into class alone.",
        items: [
          {
            title: "Teach labels in clusters",
            html: "Introduce a fallacy alongside its nearest rivals, not as an isolated specimen pinned to velvet.",
          },
          {
            title: "Require negative explanation",
            html: "Ask students not only why the chosen label fits, but why two close labels do not fit.",
          },
          {
            title: "Use short quoted passages",
            html: "The shorter the passage, the harder it is for students to hide behind tone or context fog.",
          },
          {
            title: "Return to the same cluster later",
            html: "The second and third comparison usually matter more than the first. Discrimination is learned by recurrence.",
          },
        ],
      },
    ],
    takeaway: {
      title: "The real test is not whether a student can name a fallacy, but whether they can separate it from its nearest impostors.",
      html: "If the exact split is clear, the label becomes sturdy. If the split is fuzzy, the label is probably being carried around like a decorative license plate.",
    },
    references: [
      { ...theorySourceCatalog.fallaciesSep, note: "Strong background on the range of formal and informal fallacy treatments." },
      { ...theorySourceCatalog.fallaciesIep, note: "Useful for surveying many labels and keeping comparison grounded in standard descriptions." },
      { ...theorySourceCatalog.openstaxFallacies, note: "Helpful teaching taxonomy for relevance, weak induction, assumption, and diversion." },
      { ...theorySourceCatalog.argumentSep, note: "Valuable for broader context on argument types and analogical forms of support." },
    ],
  });
}

function buildRepairTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "Spotting a fallacy is not the finish line. If all you can do is name the wound, you are halfway to being medically interesting but not yet medically useful. Repair asks a harder and better question: what would this argument need to become fairer, narrower, and more honestly supported?",
    introPanels: [
      {
        title: "The constructive mindset",
        html: "Repair starts by preserving as much of the speaker's core concern as possible. The aim is not to leave the argument in a crater, but to rebuild it on beams that can actually hold weight.",
      },
      {
        title: "Why students need this",
        html: "Without repair work, fallacy study becomes mostly punitive. With repair work, students learn what better reasoning sounds like in practice.",
      },
    ],
    sections: [
      {
        title: "The basic repair sequence",
        copy: "Most repair tasks can be handled with a small, repeatable procedure.",
        callout:
          "<strong>Repair pattern:</strong> preserve the live concern → identify the exact overreach → narrow the claim → add the missing support → retest the conclusion.",
        items: [
          {
            title: "Preserve the live concern",
            html: "Start by asking what the speaker is worried about, trying to protect, or hoping to establish. The concern is often more salvageable than the argument that currently carries it.",
          },
          {
            title: "Find the exact overreach",
            html: "Do not just say 'this goes too far.' Say how: too little evidence, missing mechanism, false choice, verbal shift, or irrelevant attack.",
          },
          {
            title: "Narrow the conclusion",
            html: "A large fraction of repair work is simply replacing a sweeping claim with one the available support can honestly bear.",
          },
          {
            title: "Add what is missing",
            html: "Sometimes the claim can survive if more data, a causal mechanism, a comparison class, or a caveat is supplied. That difference matters.",
          },
        ],
      },
      {
        title: "Common repair moves",
        copy: "Different fallacies call for different kinds of reconstruction.",
        layout: "grid",
        items: [
          {
            title: "For False dilemma",
            html: `Add the missing live options, then restate the argument more modestly. See ${theoryInternalLink("fallacies/false-dilemma/", "False dilemma")} for the classic compressed-choice pattern.`,
          },
          {
            title: "For Cherry picking",
            html: `Restore the missing evidence set, then ask whether the conclusion survives when the full record is back on the table. That is the repair demanded by ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")}.`,
          },
          {
            title: "For Correlation mistakes",
            html: `State the correlation more carefully, then add causal alternatives, time-order questions, or mechanism requirements before drawing stronger lessons. That is the repair space around ${theoryInternalLink("fallacies/correlation-is-not-causation/", "Correlation is not causation")}.`,
          },
          {
            title: "For Ad hominem",
            html: `Strip away the personal shot and restate the evidential or logical objection directly. If no objection remains, the original move was mostly theater with a necktie.`,
          },
          {
            title: "For Hasty generalization",
            html: `Shrink the scope of the claim or enlarge the evidence base. Often the repair is as simple as changing 'people are' into 'in this small sample, several people were.'`,
          },
          {
            title: "For Equivocation",
            html: `Fix the key term, define it once, and keep that meaning stable. Many slippery arguments become embarrassingly ordinary once the word is nailed to the floor.`,
          },
        ],
      },
      {
        title: "Mini-examples of repair",
        copy: "A repaired argument should sound stronger, not merely safer.",
        items: [
          {
            title: "From False dilemma to honest tradeoff",
            html: "Bad version: 'Either we ban AI in class or students stop thinking.' Better version: 'Some AI uses can short-circuit student thinking, so the class needs specific rules about when AI may be used and how the work must still show human reasoning.'",
          },
          {
            title: "From Anecdote to qualified concern",
            html: "Bad version: 'My cousin was harmed by this treatment, so the treatment is unsafe.' Better version: 'This case raises a safety concern that needs broader outcome data and comparative rates before stronger conclusions are drawn.'",
          },
          {
            title: "From Ad hominem to evidence challenge",
            html: "Bad version: 'Ignore his inflation argument; he is a billionaire.' Better version: 'His argument relies on a selective data window and does not address wage, housing, or regional price variation.'",
          },
          {
            title: "From Slippery slope to contingent warning",
            html: `Bad version: 'If we permit X, we will inevitably end at disaster Z.' Better version: 'If X is adopted without guardrails A, B, and C, it may create pressures that move policy toward Y or Z.' The difference is the difference between prophecy and analysis.`,
          },
        ],
      },
    ],
    takeaway: {
      title: "Repair is where critical thinking becomes generous without becoming soft.",
      html: "A repaired argument says only what its support has earned, but it still tries to preserve what was worth saying in the first place. That is a much better classroom habit than turning every fallacy into roadkill for applause.",
    },
    references: [
      { ...theorySourceCatalog.argumentSep, note: "Good background on argument structure and different kinds of support." },
      { ...theorySourceCatalog.criticalThinkingIep, note: "Helpful on evaluation, reconstruction, and the norms of critical thinking." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "Useful for turning diagnosis into the kinds of questions that guide repair." },
      { ...theorySourceCatalog.openstaxInferences, note: "Accessible for students who need a quick refresher on deductive, inductive, and abductive support." },
    ],
  });
}

function buildKindsOfReasoningFailureTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "People often talk as if a logical fallacy were a single species of badness. It is not. Some mistakes are structural, some are evidential, some are causal, and some are statistical. Throwing them all into one bucket is a bit like teaching students that bones, blood sugar, and broken eyeglasses are all just 'health problems' and then acting surprised when treatment becomes vague.",
    introPanels: [
      {
        title: "Why the distinctions matter",
        html: "Different reasoning failures call for different questions, different teaching moves, and different repairs. A causal blunder is not corrected the same way as a formal invalidity or a sampling mistake.",
      },
      {
        title: "Traditional and site-specific taxonomies",
        html: `Traditional logic often contrasts formal and informal fallacies. LogFall keeps that insight but also uses more teaching-friendly families such as evidential, causal, statistical, linguistic, and relevance-based failures.`,
      },
    ],
    sections: [
      {
        title: "Four large kinds of failure",
        copy: "These are not the only categories, but they are the most pedagogically useful first cuts.",
        layout: "grid",
        items: [
          {
            title: "Formal or structural",
            html: `The support fails because the conclusion does not follow from the shape of the argument. Cases like ${theoryInternalLink("fallacies/affirming-the-consequent/", "Affirming the consequent")} and ${theoryInternalLink("fallacies/denying-the-antecedent/", "Denying the antecedent")} belong here.`,
          },
          {
            title: "Evidential",
            html: `The speaker misuses, selects, or overreads the evidence. ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")} and ${theoryInternalLink("fallacies/absence-of-evidence-fallacy/", "Absence of evidence fallacy")} are classic cases.`,
          },
          {
            title: "Causal or explanatory",
            html: `The problem concerns what caused what, what mechanism is missing, or what explanation is too thin. ${theoryInternalLink("fallacies/post-hoc-ergo-propter-hoc/", "Post hoc ergo propter hoc")} lives here, as does ${theoryInternalLink("fallacies/wrong-causal-direction/", "Wrong causal direction")}.`,
          },
          {
            title: "Statistical",
            html: `The failure lies in rates, samples, distributions, uncertainty, or comparison classes. ${theoryInternalLink("fallacies/base-rate-fallacy/", "Base rate fallacy")} and ${theoryInternalLink("fallacies/survivorship-bias/", "Survivorship bias")} are good examples.`,
          },
        ],
      },
      {
        title: "What each kind needs from the teacher",
        copy: "A good class changes the tool to match the failure.",
        items: [
          {
            title: "Formal errors need structure made visible",
            html: "Students need short premises, short conclusions, and often a map or symbolic skeleton. Long political prose is a terrible first home for teaching invalid form.",
          },
          {
            title: "Evidential errors need comparison sets",
            html: "Ask what evidence is missing, selected, exaggerated, or treated as sufficient. The cure here is often not formal notation but a fuller record.",
          },
          {
            title: "Causal errors need alternative explanations",
            html: "Students should be trained to ask about sequence, mechanism, reverse causation, third variables, and counterfactual alternatives.",
          },
          {
            title: "Statistical errors need numerical humility",
            html: "Base rates, sample size, variance, regression, and uncertainty language usually matter more here than rhetoric does. The math does not have to be fancy; it just has to be present.",
          },
        ],
      },
      {
        title: "Common category mistakes",
        copy: "Misclassification produces weak teaching and weak criticism.",
        items: [
          {
            title: "Treating every informal mistake as 'emotional'",
            html: "Some arguments are vivid or heated, but their deepest problem may be sampling, category confusion, or causal overreach rather than emotional pressure.",
          },
          {
            title: "Treating every causal mistake as a formal mistake",
            html: "A causal argument may be invalid in some abstract rendering, but that often hides the more teachable point: the mechanism or evidence is inadequate.",
          },
          {
            title: "Forgetting that one passage can host multiple kinds",
            html: `A speech can mix ${theoryInternalLink("fallacies/ad-hominem/", "Ad hominem")} with ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")} and a statistical overreach in the same paragraph. Bad reasoning is perfectly capable of multitasking.`,
          },
          {
            title: "Using family labels as if they were verdicts",
            html: "A family label should point students toward the right diagnostic questions. It should not replace the more specific diagnosis.",
          },
        ],
      },
    ],
    takeaway: {
      title: "Different reasoning failures need different lenses.",
      html: "If you teach every bad argument with the same voice and the same tool, students will learn the vocabulary while missing the craft. The categories matter because the remedies matter.",
    },
    references: [
      { ...theorySourceCatalog.fallaciesSep, note: "Strong background on formal and informal distinctions and on modern fallacy theory." },
      { ...theorySourceCatalog.openstaxFallacies, note: "Useful teaching taxonomy for relevance, weak induction, unwarranted assumption, and diversion." },
      { ...theorySourceCatalog.logicInductiveSep, note: "Helpful on strong and weak inductive support, especially where probability enters the picture." },
      { ...theorySourceCatalog.statisticsSep, note: "Useful for the statistical side of evidence, inference, and methodological reasoning." },
    ],
  });
}

function buildTrueConclusionsBadArgumentsTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "One of the first ideas that irritates beginners in a useful way is that a conclusion can be true while the argument for it is still bad. Logic does not ask only, 'Did you land on the right answer?' It also asks, 'Did you get there by a route that actually supports the answer?' Guessing the right password is not the same as knowing it.",
    introPanels: [
      {
        title: "Truth and support come apart",
        html: "A conclusion can be true because reality cooperates, coincidence intervenes, or the speaker got lucky. None of that magically upgrades weak support into strong support.",
      },
      {
        title: "Why this matters",
        html: "If students collapse truth into good argument, they become easy prey for rhetoric that reaches a welcome conclusion by terrible means.",
      },
    ],
    sections: [
      {
        title: "How the split works",
        copy: "The quickest way to understand the point is through cases.",
        items: [
          {
            title: "Lucky guess",
            html: "Suppose someone says, 'The bridge is unsafe because the architect has ugly shoes,' and it turns out the bridge really is unsafe. The conclusion happens to be true; the support is still nonsense.",
          },
          {
            title: "Right conclusion, wrong route",
            html: `A debater might correctly suspect fraud, corruption, or methodological weakness, yet defend that suspicion with ${theoryInternalLink("fallacies/bare-assertion-fallacy/", "Bare assertion")} or ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")}. Being accidentally right is not the same as arguing well.`,
          },
          {
            title: "The seductive shortcut",
            html: "People often forgive a bad argument when they already like the conclusion. That is understandable, but it is one of the shortest roads from clear thinking to tribal thinking.",
          },
          {
            title: "The reverse case also matters",
            html: "A conclusion can be false even when the argument for it is carefully structured, because one or more premises are false. Good reasoning is not a miracle cure for bad starting materials.",
          },
        ],
      },
      {
        title: "Examples students remember",
        copy: "Vivid examples do a lot of the philosophical lifting here.",
        layout: "grid",
        items: [
          {
            title: "The broken clock",
            html: "A stopped clock gives the right time twice a day. Nobody concludes from that that the clock is reliable. The same charity should not be extended to arguments simply because they happen to land on a true conclusion.",
          },
          {
            title: "The bad map that still gets you home",
            html: "If a map has the river in the wrong place, the roads mislabeled, and the scale warped, but you still make it to the bakery, you do not frame the map and teach cartography from it.",
          },
          {
            title: "The suspicious prosecutor",
            html: "A prosecutor may correctly suspect guilt while still presenting irrelevant, prejudicial, or insufficient reasoning. Courts, at least in principle, care about both truth and support.",
          },
          {
            title: "The student's uncanny hunch",
            html: "A student may sense that an op-ed is weak while giving the wrong diagnosis. The hunch may be useful, but the course still needs the diagnosis to become precise.",
          },
        ],
      },
      {
        title: "What this changes in the classroom",
        copy: "Once the split is clear, the whole subject becomes cleaner.",
        items: [
          {
            title: "Students stop treating fallacy labels as ideological weapons",
            html: "The focus shifts from 'Which side is right?' to 'What support has been earned?' That is a healthier axis for critical thinking.",
          },
          {
            title: "Repair becomes possible",
            html: "If the conclusion may still be salvageable, then the class can ask what would count as better support rather than treating the entire position as radioactive waste.",
          },
          {
            title: "Humility increases",
            html: "Students realize they too may sometimes believe true things for weak reasons. That realization is annoying in exactly the right educational way.",
          },
          {
            title: "Evidence regains its dignity",
            html: "The class learns that desirable outcomes, moral alignment, or ideological familiarity do not substitute for warranted support.",
          },
        ],
      },
    ],
    takeaway: {
      title: "A true conclusion does not retroactively bless a bad argument.",
      html: "Teach students to separate truth, validity, strength, and justification. Once that habit sticks, they become much harder to impress with arguments that happen to be correct by accident, intuition, or luck.",
    },
    references: [
      { ...theorySourceCatalog.argumentSep, note: "Useful on the nature of arguments, conclusions, and support relations." },
      { ...theorySourceCatalog.fallaciesSep, note: "Helpful for keeping fallacy diagnosis tied to argumentative failure rather than to mere falsity." },
      { ...theorySourceCatalog.openstaxArguments, note: "Accessible on arguments, premises, and conclusions." },
      { ...theorySourceCatalog.openstaxInferences, note: "Useful for the difference between deductive and inductive support." },
    ],
  });
}

function buildProbabilityStatisticsTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "A surprising number of fallacies become easier to see the moment probability and statistics enter the room. Base rates, sample sizes, regression, uncertainty, and causal alternatives do not make public argument glamorous, but they do make it less likely to wander around wearing someone else's shoes and calling them evidence.",
    introPanels: [
      {
        title: "The central claim",
        html: "Many notorious fallacies are really badly handled uncertainty problems in disguise. The logic is weak because the numbers, samples, or comparison classes are weak.",
      },
      {
        title: "The teaching upside",
        html: "You do not need advanced mathematics to improve fallacy diagnosis. You need a few durable habits: ask about rates, ask about samples, ask about alternatives, and stop pretending that confidence is a substitute for proportion.",
      },
    ],
    sections: [
      {
        title: "Five statistical habits that clarify fallacies",
        copy: "These habits pay rent across dozens of cases.",
        layout: "grid",
        items: [
          {
            title: "Check the base rate",
            html: `This is the cure for ${theoryInternalLink("fallacies/base-rate-fallacy/", "Base rate fallacy")} and a quiet assistant in many medical, legal, and policy arguments.`,
          },
          {
            title: "Ask how large and how representative the sample is",
            html: `That habit exposes ${theoryInternalLink("fallacies/hasty-generalization/", "Hasty generalization")}, ${theoryInternalLink("fallacies/anecdotal-fallacy/", "Anecdotal fallacy")}, and many forms of pundit confidence dressed as evidence.`,
          },
          {
            title: "Ask what the missing comparison class is",
            html: "Many claims look forceful only because they are not being compared with the wider field in which they belong.",
          },
          {
            title: "Expect regression and noise",
            html: `Without that expectation, people invent dramatic explanations for normal fluctuation and end up in ${theoryInternalLink("fallacies/regression-fallacy/", "Regression fallacy")} territory.`,
          },
          {
            title: "Separate correlation from cause",
            html: `This does not just guard against ${theoryInternalLink("fallacies/correlation-is-not-causation/", "Correlation is not causation")}; it also disciplines explanations more broadly by forcing alternative causes back into view.`,
          },
        ],
      },
      {
        title: "Fallacy families that become clearer with statistical thinking",
        copy: "The numbers do not solve everything, but they do reveal a lot.",
        items: [
          {
            title: "Sampling failures",
            html: `Arguments based on small, skewed, or unusually vivid cases often feel persuasive because human attention is not a random sample generator. That is why ${theoryInternalLink("fallacies/survivorship-bias/", "Survivorship bias")} and ${theoryInternalLink("fallacies/spotlight-fallacy/", "Spotlight fallacy")} are classroom gold.`,
          },
          {
            title: "Causal overconfidence",
            html: `Once students learn to ask about confounders, reverse causation, and regression to the mean, several causal fallacies stop looking like deep mysteries and start looking like premature announcements.`,
          },
          {
            title: "Probability-free certainty",
            html: `Overstated confidence often hides behind emotionally charged rhetoric. Statistical literacy is one way of putting uncertainty back into a conversation that has illegally evicted it.`,
          },
          {
            title: "Policy theater",
            html: "Public arguments often compare raw counts where rates are needed, cite outliers where distributions matter, and treat one datapoint like a choir. Statistical habits make those moves much easier to resist.",
          },
        ],
      },
      {
        title: "Simple classroom moves",
        copy: "You can bring statistics into a fallacy unit without turning the class into a spreadsheet cult.",
        items: [
          {
            title: "Ask for the denominator",
            html: "When a student presents a dramatic number, ask: out of how many? That one question exposes a remarkable amount of nonsense.",
          },
          {
            title: "Make them compare two formulations",
            html: "For example: 'Three people I know had side effects' versus 'three out of ten thousand patients had side effects.' Same numerator, very different reasoning atmosphere.",
          },
          {
            title: "Require uncertainty language",
            html: "Push students to choose among words like suggests, indicates, raises concern, supports strongly, or does not yet justify. Precision in modality is half of intellectual adulthood.",
          },
          {
            title: "Map the causal alternatives",
            html: "Before accepting a cause claim, have students name at least two rival explanations. This is cheaper than a semester of statistical inference and often pedagogically better.",
          },
        ],
      },
    ],
    takeaway: {
      title: "Probability and statistics do not replace fallacy study; they sharpen it.",
      html: "The student who asks about rates, samples, uncertainty, and causal alternatives is already harder to fool. In that sense, statistical literacy is one of logic's most useful sidekicks.",
    },
    references: [
      { ...theorySourceCatalog.statisticsSep, note: "Strong philosophical background on statistical inference and evidence." },
      { ...theorySourceCatalog.inductionSep, note: "Useful on induction, probability, and the logic of projecting beyond the data." },
      { ...theorySourceCatalog.logicInductiveSep, note: "Helpful on strong and weak inductive support." },
      { ...theorySourceCatalog.openstaxInferences, note: "Accessible on deductive, inductive, and abductive reasoning." },
    ],
  });
}

function buildInsufferableFallacyTalkTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "Fallacy language can make a person sound sharper than they are, faster than they deserve, and more unbearable than anyone asked for. There is a specific species of critical thinker who can identify a straw man from thirty yards away but cannot identify the social consequences of saying so like a hall monitor with a philosophy minor. This article is a small public service.",
    introPanels: [
      {
        title: "The problem",
        html: "A good vocabulary can become a bad personality if it is used mainly for scoring points, displaying superiority, or avoiding the patient work of explanation.",
      },
      {
        title: "The better aim",
        html: "Use fallacy language to clarify, compare, repair, and self-correct. If it mainly makes the other person feel clubbed, the technique is probably being misused.",
      },
    ],
    sections: [
      {
        title: "Better habits of fallacy talk",
        copy: "These habits make the vocabulary more humane and more accurate at the same time.",
        layout: "grid",
        items: [
          {
            title: "Quote before you classify",
            html: "Point to the line that carries the problem. Diagnosis anchored in text feels less like posturing and more like analysis.",
          },
          {
            title: "Explain before you Latinize",
            html: "Plain language should do the first job. The technical label can arrive after the reasoning slip is already visible.",
          },
          {
            title: "Offer a repair",
            html: "If you can say how the argument could be made stronger, you sound like a collaborator in truth-seeking rather than a referee who enjoys the whistle too much.",
          },
          {
            title: "Use caveats",
            html: "Say when the label is strong, when it is tentative, and what would change your mind. Precision is more impressive than swagger, though admittedly less cinematic.",
          },
          {
            title: "Turn the tool inward",
            html: "The best antidote to smugness is remembering how many of these moves you yourself can commit when tired, rushed, angry, or pleased with your own paragraph.",
          },
          {
            title: "Know when not to label",
            html: "Sometimes the better move is simply: 'This needs more evidence,' 'That term is doing too much work,' or 'I think two different issues are being run together here.'",
          },
        ],
      },
      {
        title: "What bad fallacy talk often sounds like",
        copy: "You will know the tone when you hear it, but it helps to name the pattern.",
        items: [
          {
            title: "The drive-by label",
            html: "Someone says 'straw man' or 'ad hominem' and then vanishes like a morally disappointed bat. No explanation, no quotation, no repair, just airborne self-satisfaction.",
          },
          {
            title: "The encyclopedia dump",
            html: "A student names six possible fallacies for one sentence, as if uncertainty were best handled by unloading the whole museum at once.",
          },
          {
            title: "The ideological boomerang",
            html: "Labels are applied lavishly to enemies and timidly to allies. At that point the vocabulary has stopped functioning as analysis and started moonlighting as tribal decoration.",
          },
          {
            title: "The conversation derailment",
            html: "Sometimes the fallacy label itself becomes a relevance problem because it interrupts the substantive point instead of clarifying it. Irony, as always, works overtime.",
          },
        ],
      },
      {
        title: "Better replacement phrases",
        copy: "A classroom should train students in wording that opens rather than closes inquiry.",
        items: [
          {
            title: "Instead of 'That's a fallacy,' try",
            html: "'I think the conclusion is moving faster than the evidence here.'",
          },
          {
            title: "Instead of 'That's ad hominem,' try",
            html: "'That targets the person more than the actual support for the claim.'",
          },
          {
            title: "Instead of 'False dilemma,' try",
            html: "'Are those really the only live options, or have some alternatives been left out?'",
          },
          {
            title: "Instead of 'Begging the question,' try",
            html: "'It sounds as if the conclusion is already built into the support.'",
          },
          {
            title: "Instead of 'That's a straw man,' try",
            html: "'I think that rephrases the position into a weaker version than the one actually being defended.'",
          },
          {
            title: "Instead of 'That's a red herring,' try",
            html: "'That may be interesting, but I don't yet see how it answers the original point under dispute.'",
          },
          {
            title: "Instead of 'That's cherry picking,' try",
            html: "'Those examples may matter, but what happens when we bring the missing evidence back into the picture too?'",
          },
          {
            title: "Instead of 'That's hasty generalization,' try",
            html: "'That sounds broader than the sample can really support as it stands.'",
          },
          {
            title: "Instead of 'That's appeal to authority,' try",
            html: "'Expertise may be relevant here, but we still need to know whether the evidence and reasoning actually support the claim.'",
          },
          {
            title: "Instead of 'That's slippery slope,' try",
            html: "'Can you show the steps that would connect this first move to the later outcome, rather than just assuming the slide?'",
          },
          {
            title: "Instead of 'Correlation is not causation,' try",
            html: "'That pattern is interesting, but what rules out coincidence, reverse direction, or a third factor?'",
          },
          {
            title: "Instead of 'That's equivocation,' try",
            html: "'I think the key term may be shifting meaning as the argument moves forward.'",
          },
        ],
      },
    ],
    takeaway: {
      title: "Good fallacy language should make reasoning clearer, not make the speaker harder to sit next to.",
      html: "If the vocabulary increases explanation, repair, charity, and self-audit, keep it. If it mainly increases superiority theater, turn the volume down until the logic is audible again.",
    },
    references: [
      { ...theorySourceCatalog.fallaciesIep, note: "Especially good on the burden carried by any fallacy accusation." },
      { ...theorySourceCatalog.criticalThinkingIep, note: "Useful on critical evaluation and responsible reasoning practices." },
      { ...theorySourceCatalog.fallaciesSep, note: "Helpful on the diversity of fallacy conceptions and why careless labeling is easy." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "A reminder that evaluation often works better as questioning than as slogan." },
    ],
  });
}

function buildArgumentMapsTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "Argument maps help because they force reasons to stand where everyone can see them. Once premises, assumptions, and conclusions are drawn out, some fallacies stop looking mysterious and start looking like plumbing failures. Water is being routed somewhere it cannot actually go, and the map makes that harder to miss.",
    introPanels: [
      {
        title: "What a map does well",
        html: "It slows the argument down, separates claims from support, and reveals hidden premises that prose lets sneak around in dark glasses.",
      },
      {
        title: "What a map does not do",
        html: "It does not replace judgment. A beautiful map can still represent a weak argument, but at least the weakness becomes inspectable rather than atmospheric.",
      },
    ],
    sections: [
      {
        title: "The basic parts of an argument map",
        copy: "Students do not need elaborate software to learn the essentials.",
        items: [
          {
            title: "Conclusion",
            html: "What is the claim being supported? If that cannot be stated clearly, the map will wobble before the fallacy even arrives.",
          },
          {
            title: "Stated premises",
            html: "These are the reasons explicitly offered in the passage. They should be written in simple sentence form, not copied as a tangled paragraph vine.",
          },
          {
            title: "Hidden assumptions",
            html: "These are the unspoken bridges the argument needs in order to move from the premises to the conclusion. Fallacies often live here rent-free.",
          },
          {
            title: "Failure point",
            html: "Mark the exact arrow, premise, or assumption where the support breaks. This is where the diagnosis becomes more than a label.",
          },
        ],
      },
      {
        title: "Common fallacies become easier to see on a map",
        copy: "Different maps expose different failure points.",
        layout: "grid",
        items: [
          {
            title: "Straw man",
            html: `A map shows that the rebutted claim is not actually the opponent's original conclusion. The attacked node is a replacement mannequin, not the real person. See ${theoryInternalLink("fallacies/straw-man-argument/", "Straw man argument")}.`,
          },
          {
            title: "False dilemma",
            html: `The map reveals a hidden premise that only two options exist. Once that premise is written down, students can test it instead of letting it pass in formalwear. See ${theoryInternalLink("fallacies/false-dilemma/", "False dilemma")}.`,
          },
          {
            title: "Affirming the consequent",
            html: `The structure becomes stark: If P then Q; Q; therefore P. In prose this may feel plausible. On a map it begins to look like a ladder missing a rung while still expecting applause. See ${theoryInternalLink("fallacies/affirming-the-consequent/", "Affirming the consequent")}.`,
          },
          {
            title: "Correlation errors",
            html: `A map lets you write the observed association separately from the stronger causal conclusion, making the unsupported jump visible. See ${theoryInternalLink("fallacies/correlation-is-not-causation/", "Correlation is not causation")}.`,
          },
          {
            title: "Begging the question",
            html: `The map often reveals that the supposed support contains the conclusion in paraphrased clothing. Once drawn, the circle looks less majestic and more dizzy.`,
          },
          {
            title: "Red herring",
            html: `The map shows that a new branch of material is being developed that never reconnects to the original conclusion under dispute. It is a side quest dressed as progress.`,
          },
        ],
      },
      {
        title: "How to use maps in class",
        copy: "Mapping works best when it is lightweight and frequent.",
        items: [
          {
            title: "Map short passages first",
            html: "One sentence or one paragraph is enough. Long editorials can wait until students know how to separate branches and assumptions without panic.",
          },
          {
            title: "Ask for competing maps",
            html: "Different students will sometimes reconstruct the same argument differently. That disagreement is not a bug; it is the seminar finally doing something interesting.",
          },
          {
            title: "Pair mapping with repair",
            html: "Once the failure point is marked, ask students to redraw the map so the conclusion is supported properly or reduced to a narrower, honest form.",
          },
          {
            title: "Use maps to compare near neighbors",
            html: "Two arguments can sound alike but differ structurally. A map makes the difference between misrepresentation, diversion, and weak support easier to explain.",
          },
        ],
      },
    ],
    takeaway: {
      title: "Argument maps do not replace logic; they make logic visible enough to teach.",
      html: "If students can draw where the support is supposed to go, they become far better at seeing where it actually fails. That is why maps are such useful companions to fallacy study.",
    },
    references: [
      { ...theorySourceCatalog.argumentMapsStudy, note: "A classic argument for the pedagogical value of argument mapping." },
      { ...theorySourceCatalog.rationaleOnline, note: "Practical teaching resource on argument mapping and classroom use." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "Useful for pairing maps with critical questions and scheme evaluation." },
      { ...theorySourceCatalog.argumentSep, note: "Helpful for the general structure of arguments and analogical argument forms." },
    ],
  });
}

function buildDebateEditorialNewsTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "If students learn fallacies only from toy examples, they may perform well on worksheets while missing the moves completely in real life. Editorials, debate exchanges, campaign statements, interviews, and news-adjacent commentary solve that problem because they contain the mess that real reasoning always brings along: speed, framing, emotion, selective quoting, audience targeting, and strategic omission.",
    introPanels: [
      {
        title: "Why live materials matter",
        html: "Students need to see fallacies where they actually live: in public argument shaped by time pressure, persuasion, identity, and incomplete evidence.",
      },
      {
        title: "Why they also need guardrails",
        html: "Real materials are richer but also noisier. That means the class needs clear quoting rules, comparison rules, and a strong norm against partisan point-scoring masquerading as analysis.",
      },
    ],
    sections: [
      {
        title: "A strong assignment pattern",
        copy: "The same basic sequence works across debates, editorials, and news analysis.",
        callout:
          "<strong>Recommended workflow:</strong> select a short passage → quote the key lines → identify the best label → rule out two nearby labels → rebut or repair the claim.",
        items: [
          {
            title: "Keep the source short",
            html: "A paragraph, exchange, or clip transcript is often better than a full article because it forces close attention to the specific move under inspection.",
          },
          {
            title: "Require sufficient quotation",
            html: "Students should quote enough of the original to make the misstep visible. Otherwise the analysis floats free of the source and becomes a vibe review.",
          },
          {
            title: "Make comparison mandatory",
            html: "Insist that students name at least one nearby label and explain why it fits less well. This is where superficial confidence usually goes to be corrected.",
          },
          {
            title: "End with repair",
            html: "Ask students to rewrite the claim or paragraph into a stronger form. That one step changes the assignment from taxidermy into instruction.",
          },
        ],
      },
      {
        title: "Good activity types",
        copy: "Not every classroom task needs to look like a quiz.",
        layout: "grid",
        items: [
          {
            title: "Debate clip diagnosis",
            html: "Use a short debate exchange and ask students to mark the exact line where the shift or overreach occurs.",
          },
          {
            title: "Editorial autopsy",
            html: "Take one op-ed paragraph and have students identify the claim, the support, the hidden assumption, and the likely fallacy if any.",
          },
          {
            title: "Headline vs. body comparison",
            html: "Compare the headline's implication to what the article body actually supports. This is a fertile site for contextomy, overreach, and false balance.",
          },
          {
            title: "Two-source comparison",
            html: "Give students two differently framed reports on the same issue and ask which reasoning moves are shared, which differ, and where each overstates the case.",
          },
          {
            title: "Live repair workshop",
            html: "Project a fallacious paragraph and have groups compete to produce the strongest repaired version rather than the fastest label.",
          },
          {
            title: "Caveat drill",
            html: "Give students a plausible label and make them state the strongest caveat against applying it too quickly. This teaches restraint as a classroom skill, not as a mood.",
          },
        ],
      },
      {
        title: "What to grade",
        copy: "Rubrics should reward reasoning quality, not performative certainty.",
        items: [
          {
            title: "Accuracy of the quoted evidence",
            html: "Did the student quote the right lines, or did they gesture vaguely at the source and hope everyone would be too polite to notice?",
          },
          {
            title: "Precision of the label",
            html: "Does the chosen label fit better than the nearby alternatives? If not, the answer is incomplete even if it sounds fluent.",
          },
          {
            title: "Depth of explanation",
            html: "A strong explanation tells how the reasoning moves from premise to conclusion and where the support fails, not just that it fails.",
          },
          {
            title: "Quality of the repair",
            html: "Can the student salvage the central concern and rewrite the argument into a stronger, narrower, fairer form?",
          },
        ],
      },
    ],
    takeaway: {
      title: "Real materials make fallacy study messier, and that is exactly why they make it better.",
      html: "If students can quote, compare, diagnose, and repair reasoning in debate clips, editorials, and public rhetoric, they are much closer to genuine critical thinking than if they can only recognize museum-grade textbook specimens.",
    },
    references: [
      { ...theorySourceCatalog.openstaxFallacies, note: "Useful starting point for course-level examples and categories." },
      { ...theorySourceCatalog.criticalThinkingIep, note: "Helpful on evaluation standards and rational appraisal." },
      { ...theorySourceCatalog.pewNewsSocial, note: "Useful context for how many students encounter argument through social and news feeds." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "Good for turning public-source analysis into structured evaluation rather than free-floating opinion." },
    ],
  });
}

function buildAnalogyRoleTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "Analogy is one of reason's most useful devices because it lets us test a pattern in a cleaner setting. It can illuminate, compare, restrain, criticize, and occasionally rescue an argument from its own melodrama. It can also mislead badly when the comparison is superficial. That is why analogy deserves both trust and supervision.",
    introPanels: [
      {
        title: "Why analogy matters",
        html: "Analogy helps us see structure by carrying it into a less noisy case. That is why it is so powerful in teaching, criticism, law, science, and ordinary conversation.",
      },
      {
        title: "Why it needs discipline",
        html: "A flashy analogy can seduce people into thinking two cases are similar where the relevant structure is not. The goal is not similarity in general, but similarity where the reasoning lives.",
      },
    ],
    sections: [
      {
        title: "What analogy can do in criticism",
        copy: "Its uses are broader than merely naming bad arguments.",
        items: [
          {
            title: "Expose structure",
            html: "A parallel case can reveal a hidden leap, missing mechanism, or erased alternative more quickly than a technical lecture can.",
          },
          {
            title: "Lower rhetorical temperature",
            html: "By moving the same reasoning form into a neutral setting, analogy can reduce defensiveness and make the issue inspectable rather than tribal.",
          },
          {
            title: "Test proportionality",
            html: "Analogies are excellent for asking whether the response, comparison, or inference is wildly out of scale with the facts.",
          },
          {
            title: "Support repair",
            html: "A strong analogy can show not only why the current move fails, but what a more careful version would have to keep or surrender.",
          },
        ],
      },
      {
        title: "What makes an analogy strong",
        copy: "The key test is relevance, not decorative cleverness.",
        layout: "grid",
        items: [
          {
            title: "Relevant similarity",
            html: "The shared features have to matter to the conclusion. Two things can be alike in color, chronology, or drama while differing exactly where the argument needs them to align.",
          },
          {
            title: "Visible disanalogies",
            html: `A good critic should state the important differences too. Otherwise the analogy begins auditioning for ${theoryInternalLink("fallacies/false-analogy/", "False analogy")}.`,
          },
          {
            title: "A modest conclusion",
            html: "The weaker and more disciplined the analogical conclusion, the stronger the move tends to be. When analogy tries to do too much, it usually tears a muscle.",
          },
          {
            title: "Structural over surface similarity",
            html: "The best analogies preserve the relations that matter rather than merely sharing a few eye-catching details. Costume similarity is not conceptual similarity.",
          },
        ],
      },
      {
        title: "How analogy can go wrong",
        copy: "Because analogy is powerful, its failures are worth teaching explicitly.",
        items: [
          {
            title: "False analogy",
            html: `This is the direct failure mode: the comparison looks helpful until the missing relevant similarity becomes obvious. See ${theoryInternalLink("fallacies/false-analogy/", "False analogy")}.`,
          },
          {
            title: "Decorative analogy",
            html: "The comparison is vivid and quotable but adds no genuine inferential support. It is basically rhetorical garnish wearing a fake ID.",
          },
          {
            title: "Heat-seeking analogy",
            html: "Some comparisons are designed less to clarify than to inflame. These are often memorable for all the wrong reasons and distort judgment more than they help it.",
          },
          {
            title: "Overextended analogy",
            html: "A decent initial comparison is stretched far beyond the point where its structure can responsibly carry the conclusion.",
          },
        ],
      },
    ],
    takeaway: {
      title: "Analogy is one of criticism's best tools when it preserves structure and one of its worst when it merely performs resemblance.",
      html: "The right question is never 'Are these things identical?' The right question is 'Are they similar in the respects that matter for the conclusion at issue?' That is where rational criticism either sharpens or wanders off in costume.",
    },
    references: [
      { ...theorySourceCatalog.analogySep, note: "The major reference point for evaluating analogical reasoning and its criteria." },
      { ...theorySourceCatalog.argumentSep, note: "Useful on analogical arguments as a general form of support." },
      { ...theorySourceCatalog.fallaciesSep, note: "Helpful for linking analogy to fallacy theory and informal evaluation." },
      { ...theorySourceCatalog.waltonCriticalQuestions, note: "Useful for the larger scheme-and-question approach to evaluating informal reasoning." },
    ],
  });
}

function buildAlgorithmicMediaTheoryArticleContent(article) {
  return buildStructuredTheoryArticleContent(article, {
    deck:
      "Algorithmic media changes the ecology of bad reasoning. Arguments now compete not just for truth or coherence, but for clicks, retention, shares, clips, outrage, and emotional memorability. Under those incentives, some fallacies become especially fit. They are not always the most rational; they are simply the most reproductively successful in the feed.",
    introPanels: [
      {
        title: "The medium matters",
        html: "A clipped headline, a short video segment, and a feed optimized for engagement do not merely transport reasoning. They reshape what kinds of reasoning get noticed, repeated, and rewarded.",
      },
      {
        title: "The classroom implication",
        html: "Students should not learn fallacies as though they belong only to formal debate or old-fashioned essays. The feed has its own favorite distortions, and they need to be taught in that habitat.",
      },
    ],
    sections: [
      {
        title: "Fallacies algorithmic media tends to favor",
        copy: "Not because they are good, but because they are sticky.",
        layout: "grid",
        items: [
          {
            title: "Cherry picking",
            html: `Clips, screenshots, and isolated anecdotes are almost purpose-built for ${theoryInternalLink("fallacies/cherry-picking/", "Cherry picking")}. The feed loves fragments because fragments travel faster than context.`,
          },
          {
            title: "Contextomy",
            html: `Short clips invite ${theoryInternalLink("fallacies/contextomy/", "Contextomy")}: extract the line, remove the setting, and let the audience infer the rest with confidence that would be comic if it were not so efficient.`,
          },
          {
            title: "False dilemma",
            html: `Binary framing performs well online because it is easy to caption, easy to sort by tribe, and easy to share. That makes ${theoryInternalLink("fallacies/false-dilemma/", "False dilemma")} unusually vigorous in the wild.`,
          },
          {
            title: "Appeal to emotion",
            html: `Content optimized for fear, disgust, or triumph predictably rewards ${theoryInternalLink("fallacies/appeal-to-emotion/", "Appeal to emotion")} and its close relatives.`,
          },
          {
            title: "False balance",
            html: `When platforms flatten expertise and treat every take as just another tile in the feed, ${theoryInternalLink("fallacies/false-balance/", "False balance")} can start to feel like fairness instead of distortion.`,
          },
          {
            title: "Misleading vividness",
            html: `Highly memorable single cases can dominate public attention and create the impression that salience itself is evidence. That is exactly the climate in which ${theoryInternalLink("fallacies/misleading-vividness/", "Misleading vividness")} thrives.`,
          },
        ],
      },
      {
        title: "Why these fallacies travel well",
        copy: "The answer usually lies in incentive design rather than in abstract logic alone.",
        items: [
          {
            title: "Speed beats completeness",
            html: "Short, vivid, morally legible fragments travel farther than nuanced evidential comparison. That is not a bug in the feed; it is often the business model with nicer shoes.",
          },
          {
            title: "Emotion boosts memory",
            html: "If the content is frightening, enraging, or identity-affirming, it is easier to recall and easier to share, even when the reasoning is brittle.",
          },
          {
            title: "Audiences meet arguments out of order",
            html: "People often see reactions before sources, clips before full interviews, and conclusions before premises. That disorder creates lovely conditions for confusion and overreach.",
          },
          {
            title: "Correction is slower than distortion",
            html: "A misleading clip can travel widely before a careful reconstruction catches up. By then the feed has usually moved on to fresh mischief.",
          },
        ],
      },
      {
        title: "What teachers can do with this",
        copy: "A modern critical thinking class should treat the feed as a case environment, not as background wallpaper.",
        items: [
          {
            title: "Use clip-to-context exercises",
            html: "Show the short clip first, then the full source. Ask students which fallacies became easier or harder to detect once the missing context was restored.",
          },
          {
            title: "Compare headline, body, and repost commentary",
            html: "The reasoning often mutates across those layers. That mutation itself is an excellent teaching object.",
          },
          {
            title: "Ask what the platform is rewarding",
            html: "Students should learn to ask not only whether a claim is weak, but what feature of the environment makes that weakness profitable.",
          },
          {
            title: "Teach self-defense, not just external critique",
            html: "The feed does not merely contain fallacies; it trains appetites for them. Students should learn where their own attention is easiest to hijack.",
          },
        ],
      },
    ],
    takeaway: {
      title: "Algorithmic media does not create fallacies from nothing, but it does cultivate the ones that are fastest, loudest, and easiest to share.",
      html: "Teaching fallacies today means teaching the conditions that amplify them. Otherwise students may become good at naming old textbook specimens while walking straight past the live ones glowing in their own pockets.",
    },
    references: [
      { ...theorySourceCatalog.pewSocialMediaUse, note: "Useful baseline on how deeply social media is woven into ordinary information habits." },
      { ...theorySourceCatalog.pewNewsSocial, note: "Helpful on the value people find in social news feeds and the growing concern about inaccuracy." },
      { ...theorySourceCatalog.pewTechCompanies, note: "Useful context on public concerns about large technology platforms and influence." },
      { ...theorySourceCatalog.fallaciesSep, note: "Helpful for connecting media-specific cases back to general fallacy theory." },
    ],
  });
}

function renderFeatureReveal({ title, fallacyLabel, fallacySlug, bodyHtml }) {
  const referenceHtml = fallacySlug
    ? `<p class="feature-diagnosis-meta"><strong>Closest LogFall entry:</strong> <a class="inline-link" href="../../fallacies/${fallacySlug}/">${escapeHtml(fallacyLabel)}</a></p>`
    : "";
  return `<details class="feature-reveal">
    <summary>${escapeHtml(title)}</summary>
    <div class="feature-reveal-body">
      <p>${bodyHtml}</p>
      ${referenceHtml}
    </div>
  </details>`;
}

function buildAlabamaMapHeadlineFeatureContent(article) {
  const reveals = [
    renderFeatureReveal({
      title: "Reveal 1: Single cause fallacy",
      fallacyLabel: "Single cause fallacy",
      fallacySlug: "single-cause-fallacy",
      bodyHtml:
        `The headline can nudge a reader toward a one-variable story: the map was blocked <em>because</em> it had only one majority-Black district. That is a tempting shortcut, but it flattens a larger legal argument into one visible feature. A fuller reading has to ask what broader pattern of district design, vote dilution claims, remedial history, and court reasoning sat behind the ruling before one descriptor is treated as the whole cause.`,
    }),
    renderFeatureReveal({
      title: "Reveal 2: Correlation is not causation",
      fallacyLabel: "Correlation is not causation",
      fallacySlug: "correlation-is-not-causation",
      bodyHtml:
        `In Alabama politics, race and party voting behavior often overlap in the same districts. A quick reaction can therefore jump from <em>majority-Black</em> to <em>Democratic</em> and then to a clean causal story about what the court was really doing. But overlap by itself does not settle whether race, party, legal standards, or some combination of them is carrying the explanatory weight in this specific ruling.`,
    }),
    renderFeatureReveal({
      title: "Reveal 3: Equivocation",
      fallacyLabel: "Equivocation",
      fallacySlug: "equivocation",
      bodyHtml:
        `The phrase <em>majority-Black district</em> can do several jobs at once: it can identify a district demographically, suggest a voting pattern, or sound like an explanation of the court’s intervention. Sliding between those meanings creates an illusion of clarity. A label that describes a district is not automatically the same thing as the reason a court accepted or rejected a map.`,
    }),
    renderFeatureReveal({
      title: "Reveal 4: False dilemma",
      fallacyLabel: "False dilemma",
      fallacySlug: "false-dilemma",
      bodyHtml:
        `Public reactions often harden into an either-or: either the case was really about race or it was really about party. That framing is cleaner than the underlying reality. Real redistricting disputes can involve overlapping race-and-party evidence, plus a separate legal question about what the governing standard actually requires, so forcing the issue into only two boxes can hide the most important third and fourth factors.`,
    }),
    renderFeatureReveal({
      title: "Reveal 5: Cherry picking",
      fallacyLabel: "Cherry picking",
      fallacySlug: "cherry-picking",
      bodyHtml:
        `A reader can seize on the most vivid phrase in the headline — <em>1 majority-Black district</em> — and let that stand in for the whole case. That leaves out the wider record: prior litigation, the court’s explanation, the state’s remedial attempts, and the structure of the voting-rights analysis. Once one salient phrase is treated as the whole dispute, the evidence base has already been thinned.`,
    }),
  ].join("");
  const teacherGuidePanels = renderTheoryPanels(
    [
      {
        title: "Recommended level and timing",
        html: "Best for upper high school, intro college, civics, rhetoric, media literacy, or critical thinking classes. A strong first pass is 20 to 30 minutes; a fuller discussion-and-writing version works well in 40 to 55 minutes.",
      },
      {
        title: "Teaching goal",
        html: "Help students distinguish between <em>what a headline says</em>, <em>what it tempts them to infer</em>, and <em>what the fuller article or legal context may actually justify</em>.",
      },
      {
        title: "Before the reveal",
        html: "Show only the headline first. Ask students to write two short sentences: one stating exactly what the headline explicitly says, and one naming the strongest additional claim they feel pulled to infer from it. Then have them add a third line: <em>What about this issue makes me want that inference to be true or false?</em>",
      },
      {
        title: "Core distinction to teach",
        html: "Keep pressing the difference between a demographic description, a voting pattern, a causal explanation, and a legal standard. Much of the confusion in this case comes from treating those as if they were interchangeable.",
      },
      {
        title: "Self-knowledge checkpoint",
        html: "Ask students whether they felt a stronger urge to call the headline misleading because they dislike the outlet or the politics around race, or a stronger urge to excuse it because they agree with its likely audience. That pause is part of the lesson, not a side note.",
      },
      {
        title: "Best discussion sequence",
        html: "Start with the single-cause issue, move to the race-party overlap problem, then ask where students felt political pull in one direction or the other, and only then ask whether the class has started forcing the case into a false either-or. After that introduce equivocation and cherry picking.",
      },
      {
        title: "Where students often slip",
        html: "Students may overcorrect and say the headline itself is automatically dishonest, or undercorrect because criticism of the wording feels like criticism of their side. The better lesson is more careful: a headline can be factually compact yet still invite bad inferences if readers treat a descriptive phrase as the whole causal story.",
      },
    ],
    "two-column compact-columns",
  );

  return `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">${featuresSectionLabel}</a><span>/</span><strong>${escapeHtml(article.title)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">${featuresSectionLabel}</p>
      <h2 class="detail-title">${escapeHtml(article.title)}</h2>
      <p class="feature-page-date">${escapeHtml(article.date || "")}</p>
      <p class="detail-deck">
        A short headline does not have room to carry a full court opinion on its back. The important question is not only whether the wording is fair, but
        what reasoning slips a reader can make if the compressed phrasing is treated as a full explanation of the case.
      </p>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Spot it first</h3>
          <p class="section-copy">Read the headline, pause, and see whether you can name the likely reasoning mistakes before opening the reveal boxes below. Just as important, notice whether your politics make you quicker to accuse the headline of distortion or quicker to defend it.</p>
        </div>
      </div>
      <div class="detail-section feature-headline-shell">
        <p class="eyebrow">Headline under review</p>
        <h3 class="feature-headline">“Court blocks Alabama from using congressional map with 1 majority-Black district”</h3>
        <p class="feature-source-line">
          Source:
          <a class="inline-link" href="https://www.cbsnews.com/news/court-alabama-congressional-district-midterms/">CBS News — Melissa Quinn — May 26, 2026</a>
        </p>
        <ul class="feature-question-list">
          <li>What exactly does the headline state, and what does it merely invite the reader to infer?</li>
          <li>Does it tempt you to explain a legally messy ruling with just one visible variable?</li>
          <li>Are race, party voting behavior, and legal standards being collapsed into one cause?</li>
          <li>Would a more careful reader distinguish between a demographic description and an explanation of the ruling?</li>
          <li>If the wording already feels obviously fair or obviously deceptive to you, what prior political assumption is helping produce that certainty?</li>
        </ul>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">A cleaner rewrite</h3>
          <p class="section-copy">Here is a version that keeps the news value while doing less hidden explanatory work.</p>
        </div>
      </div>
      <div class="detail-section feature-headline-shell">
        <p class="eyebrow">Possible rewrite</p>
        <h3 class="feature-headline">“Court blocks Alabama from using current congressional map in dispute over Black voter representation”</h3>
        <p class="feature-source-line">Why this is cleaner: it still states the court action and the subject of the dispute, but it avoids implying that one visible district characteristic by itself fully explains the ruling.</p>
      </div>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Reveal the likely fallacies</h3>
          <p class="section-copy">These are not all guaranteed to be in the headline itself as a deliberate move. They are the main reasoning traps the headline can trigger in a fast reader or partisan commentator, especially when agreement or hostility is already in place before the analysis begins.</p>
        </div>
      </div>
      ${reveals}
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">A more careful reading</h3>
          <p class="section-copy">The safer takeaway is not “the map was blocked because it was Black” and not “the map was blocked only because it was Democratic.” The safer takeaway is that a compressed headline about a race-and-representation case can invite both kinds of oversimplification if the reader turns one descriptive phrase into the whole causal story. Part of the discipline here is logical, and part is personal: can you keep your political loyalties from doing the reading for you?</p>
        </div>
      </div>
      <div class="two-column compact-columns">
        <article class="note-panel">
          <h4>What the headline does well</h4>
          <p class="muted">It identifies the immediate public-facing fact that will matter to many readers: the map had just one majority-Black district and the court stopped it from being used. That makes it newsworthy and legible quickly.</p>
        </article>
        <article class="note-panel">
          <h4>What the reader still has to supply carefully</h4>
          <p class="muted">The causal explanation. A headline can describe the surface of a dispute without settling why the court ruled, what legal standard was applied, or how race and party overlap in the underlying geography.</p>
        </article>
      </div>
    </section>

    <section class="detail-section section-block">
      <p class="eyebrow">Teaching note</p>
      <h3 class="section-title">Why this works as a weekly feature model</h3>
      <p class="section-copy">This is the kind of headline that works well for recurring fallacy analysis: short, vivid, politically charged, and just compressed enough to invite overconfident explanations. The point of the feature is not to sneer at headline writers. It is to help readers notice when their own reasoning outruns what the words on the page actually establish, and when political identity makes that overreach easier to miss.</p>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h3 class="section-title">Teacher’s guide</h3>
          <p class="section-copy">Use this weekly case as a structured classroom exercise rather than a one-shot diagnosis. The best outcome is not that students memorize labels, but that they learn to separate description, inference, and explanation more carefully while also noticing how agreement, anger, resentment, or tribal loyalty tug the diagnosis around.</p>
        </div>
      </div>
      <div class="detail-section theory-callout">
        <p class="theory-formula"><strong>Suggested classroom rhythm:</strong> headline only → private diagnosis → private bias check → pair discussion → reveal one fallacy at a time → compare with the cleaner rewrite → short written reflection on both logic and self-control.</p>
      </div>
      ${teacherGuidePanels}
      <div class="two-column compact-columns section-block">
        <article class="note-panel">
          <h4>Useful board prompt</h4>
          <p class="muted">Write these four columns on the board: <strong>explicitly stated</strong>, <strong>strongly implied</strong>, <strong>not yet justified</strong>, and <strong>my side wants this to mean</strong>. Then have the class place pieces of their interpretation into the right column before any fallacy labels appear.</p>
        </article>
        <article class="note-panel">
          <h4>Short writing task</h4>
          <p class="muted">Ask students to rewrite the headline in one sentence and then explain, in three to five sentences, which reasoning trap they were most tempted by, what political instinct strengthened that temptation, and how their rewrite avoids both problems.</p>
        </article>
      </div>
    </section>

    ${renderTheoryReferencesSection([
      {
        title: "CBS News: Court blocks Alabama from using congressional map with 1 majority-Black district",
        url: "https://www.cbsnews.com/news/court-alabama-congressional-district-midterms/",
        note: "The specific headline used for this first feature page.",
      },
      {
        title: "AP: Federal court blocks Alabama plan for new congressional districts that could help Republicans",
        url: "https://apnews.com/article/b67125657b36e9b915ea9bc5d587d08c",
        note: "Useful comparison because it frames the same ruling differently.",
      },
      {
        title: "All About Redistricting: Singleton v. Allen",
        url: "https://redistricting.lls.edu/case/singleton-v-allen/",
        note: "Helpful background on the larger case context and timeline.",
      },
    ])}
  `;
}

function buildFeatureArticlePage(article) {
  const featureArticleBuilders = {
    "how-one-alabama-map-headline-invites-several-fallacies": buildAlabamaMapHeadlineFeatureContent,
  };
  const content = (featureArticleBuilders[article.slug] || buildAlabamaMapHeadlineFeatureContent)(article);

  return pageShell({
    title: `${article.title} | LogFall ${featuresSectionLabel}`,
    description: featureArticleSeoDescription(article),
    prefix: "../../",
    currentSection: "features",
    canonicalPath: `features/${article.slug}/`,
    ogType: "article",
    keywords: [
      article.title,
      "logical fallacies in headlines",
      "headline analysis",
      "news rhetoric",
      "current events",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: featuresSectionLabel, path: "features/" },
        { name: article.title, path: `features/${article.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "Article",
        headline: article.title,
        description: featureArticleSeoDescription(article),
        url: absoluteUrl(`features/${article.slug}/`),
        author: {
          "@type": "Person",
          name: "Phil Stilwell",
        },
        publisher: publisherSchema(),
        dateModified: buildDate,
      },
      learningResourceSchema({
        name: article.title,
        path: `features/${article.slug}/`,
        description: featureArticleSeoDescription(article),
        about: ["logical fallacies", "news headlines", "media rhetoric"],
        teaches: ["how to test headline inferences", "how to slow down causal and framing mistakes in current events"],
        learningResourceType: ["Article", "Feature", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: [article.title, "headline fallacies", "news rhetoric", "critical thinking"],
      }),
    ],
    content,
  });
}

function buildTheoryArticlePage(article) {
  const theoryArticleBuilders = {
    "fallacy-rebuttals-without-fallacy-naming": buildRebuttalsTheoryArticleContent,
    "teaching-logical-fallacies-a-classroom-process-and-curriculum": buildTeachingCurriculumTheoryArticleContent,
    "teaching-logical-fallacies-with-ai-gems-and-prompted-agents": buildAiGemsTheoryArticleContent,
    "how-to-distinguish-fallacies-from-cognitive-biases": buildBiasesVsFallaciesTheoryArticleContent,
    "when-not-to-call-something-a-fallacy": buildWhenNotToCallFallacyTheoryArticleContent,
    "near-neighbors-how-to-tell-similar-fallacies-apart": buildNearNeighborsTheoryArticleContent,
    "how-to-repair-a-fallacious-argument": buildRepairTheoryArticleContent,
    "formal-informal-causal-and-statistical-kinds-of-reasoning-failure": buildKindsOfReasoningFailureTheoryArticleContent,
    "why-true-conclusions-can-still-have-bad-arguments": buildTrueConclusionsBadArgumentsTheoryArticleContent,
    "how-probability-and-statistics-clarify-logical-fallacies": buildProbabilityStatisticsTheoryArticleContent,
    "how-to-use-fallacy-language-without-becoming-insufferable": buildInsufferableFallacyTalkTheoryArticleContent,
    "argument-maps-for-common-fallacies": buildArgumentMapsTheoryArticleContent,
    "teaching-fallacies-through-debate-editorials-and-news-analysis": buildDebateEditorialNewsTheoryArticleContent,
    "the-role-of-analogy-in-rational-criticism": buildAnalogyRoleTheoryArticleContent,
    "fallacies-in-the-age-of-algorithmic-media": buildAlgorithmicMediaTheoryArticleContent,
  };
  const content = (theoryArticleBuilders[article.slug] || buildRebuttalsTheoryArticleContent)(article);

  return pageShell({
    title: `${article.title} | LogFall Theory`,
    description: theoryArticleSeoDescription(article),
    prefix: "../../",
    currentSection: "theory",
    canonicalPath: `theory/${article.slug}/`,
    ogType: "article",
    keywords: [
      article.title,
      "logical fallacies",
      "critical thinking",
      "argument analysis",
      "reasoning pedagogy",
      "theory article",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Theory", path: "theory/" },
        { name: article.title, path: `theory/${article.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "Article",
        headline: article.title,
        description: theoryArticleSeoDescription(article),
        url: absoluteUrl(`theory/${article.slug}/`),
        author: {
          "@type": "Person",
          name: "Phil Stilwell",
        },
        publisher: publisherSchema(),
        dateModified: buildDate,
      },
      learningResourceSchema({
        name: article.title,
        path: `theory/${article.slug}/`,
        description: theoryArticleSeoDescription(article),
        about: ["logical fallacies", "critical thinking", "argument analysis"],
        teaches: ["how to analyze reasoning failures", "how to compare, rebut, and repair arguments"],
        learningResourceType: ["Article", "Teaching resource"],
        educationalUse: ["teaching", "self-study"],
        keywords: [article.title, "logical fallacies", "critical thinking", "argument analysis", "reasoning pedagogy"],
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
      <h2 class="detail-title">A mixed 10-question assessment.</h2>
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

function buildDialogueAssessmentIndexPage() {
  const bankPayload = dialogueAssessmentBank.map((item) => ({
    ...item,
    fallacyUrl: item.fallacySlug ? `../fallacies/${item.fallacySlug}/` : "",
    imagePath: `../assets/assessment-dialogues/${item.id}.png`,
  }));

  const answerGuide = dialogueAssessmentChoices
    .map((choice) => {
      const [side, kind = ""] = choice.label.split(" ");
      if (choice.key === "none") {
        return `
          <div class="dialogue-answer-chip">
            <span class="dialogue-answer-chip-side">No speaker</span>
            <span class="dialogue-answer-chip-kind">None</span>
          </div>`;
      }
      return `
        <div class="dialogue-answer-chip">
          <span class="dialogue-answer-chip-side">${escapeHtml(side)}</span>
          <span class="dialogue-answer-chip-kind">${escapeHtml(kind)}</span>
        </div>`;
    })
    .join("");

  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Assessment</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Assessment</p>
      <h2 class="detail-title">A challenging 10-dialogue test of where the fallacy is, if it is there at all.</h2>
      <p class="detail-deck">
        Each set draws from a bank of 40 difficult six-turn dialogues. In every item, there is at most one fallacy or no fallacy at all.
        Your task is to decide whether the mistake is on the left, on the right, or nowhere, and whether it is formal or informal.
      </p>
      <div class="note-panel assessment-primer-panel">
        <div class="assessment-primer-grid">
          <div>
            <h4>How this works</h4>
            <p class="muted">Each speaker gets three turns. Read the whole exchange before deciding. Some items are traps for overdiagnosis, so <strong>None</strong> is a real answer when the reasoning holds up.</p>
          </div>
          <div>
            <h4>Answer structure</h4>
            <p class="muted">Every 10-item set is balanced by design, with 2 items in each answer class.</p>
          </div>
        </div>
        <div class="dialogue-answer-guide assessment-primer-guide">
          ${answerGuide}
        </div>
      </div>
    </section>

    <section class="panel search-panel assessment-runner-panel dialogue-assessment-panel" data-dialogue-assessment-shell data-assessment-size="10">
      <div class="section-header">
        <div>
          <h3 class="section-title">Take the dialogue assessment</h3>
          <p class="section-copy">These items are deliberately subtle. They test whether you can locate a fallacy precisely, classify its type, and resist inventing one when the exchange is actually sound. Formal fallacies go wrong in the structure of the reasoning, while informal fallacies go wrong through relevance, evidence, wording, or framing.</p>
        </div>
      </div>
      <div class="assessment-toolbar">
        <button class="button button-primary button-compact" type="button" data-dialogue-assessment-new>Load another set</button>
        <a class="button button-secondary button-compact" href="../output/pdf/logfall-dialogue-assessment.pdf">Download 40-item PDF</a>
        <a class="button button-secondary button-compact" href="../fallacies/">Study the full reference</a>
      </div>
      <div class="assessment-items" data-dialogue-assessment-items></div>
      <div class="assessment-actions">
        <button class="button button-primary" type="button" data-dialogue-assessment-grade>Grade this assessment</button>
      </div>
      <div class="detail-section assessment-results hidden" data-dialogue-assessment-results role="status" aria-live="polite"></div>
      <script id="dialogue-assessment-bank" type="application/json">${safeJsonForScript(bankPayload)}</script>
      <noscript>
        <div class="note-panel search-empty">
          <h4>JavaScript is required for this assessment</h4>
          <p class="muted">This page builds a balanced dialogue set in the browser. If scripting is disabled, you can still study the full reference in <a class="text-link" href="../fallacies/">All Fallacies</a>.</p>
        </div>
      </noscript>
    </section>
  `;

  return pageShell({
    title: "Assessment: difficult dialogue-based logical fallacy test | LogFall",
    description: dialogueAssessmentIndexSeoDescription(),
    prefix: "../",
    currentSection: "assessment",
    canonicalPath: "assessment/",
    keywords: [
      "logical fallacy assessment",
      "dialogue fallacy assessment",
      "formal vs informal fallacy test",
      "critical thinking assessment",
      "difficult fallacy quiz",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Assessment", path: "assessment/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Assessment",
        url: absoluteUrl("assessment/"),
        description: dialogueAssessmentIndexSeoDescription(),
        publisher: publisherSchema(),
        about: ["logical fallacies", "formal fallacies", "informal fallacies", "dialogue analysis"],
      },
      learningResourceSchema({
        name: "Assessment",
        path: "assessment/",
        description: dialogueAssessmentIndexSeoDescription(),
        about: ["logical fallacies", "formal fallacies", "informal fallacies", "dialogue analysis"],
        teaches: ["fallacy identification", "speaker-side diagnosis", "formal vs informal classification"],
        learningResourceType: ["Assessment", "Quiz", "Dialogue exercise"],
        educationalUse: ["assessment", "teaching", "self-study"],
        keywords: ["dialogue fallacy assessment", "formal and informal fallacy quiz", "critical thinking assessment"],
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

function mapCompactDefinition(record) {
  const definition = String(record.definition || record.notes || "").trim();
  if (!definition) return "A reasoning pattern in the LogFall fallacy library.";
  return definition.length > 250 ? `${definition.slice(0, 247).trimEnd()}...` : definition;
}

function buildMapData(records, categories) {
  const points = records
    .filter((record) => record && record.name && record.slug)
    .map((record) => {
      const gauges = rhetoricGaugesForRecord(record);
      const primaryCategory = (record.categories || [])[0] || "Uncategorized";
      const pedagogy = pedagogyForRecord(record);
      const values = Object.fromEntries(
        mapDimensions.map((dimension) => [dimension.slug, gauges[dimension.slug].value]),
      );
      const dimensionNotes = Object.fromEntries(
        mapDimensions.map((dimension) => [
          dimension.slug,
          {
            band: gauges[dimension.slug].band,
            summary: gauges[dimension.slug].summary,
          },
        ]),
      );

      return {
        slug: record.slug,
        name: record.name,
        href: `../fallacies/${record.slug}/`,
        definition: mapCompactDefinition(record),
        categories: record.categories || [],
        categorySlugs: (record.categories || []).map((category) => slugify(category)),
        category: primaryCategory,
        categorySlug: slugify(primaryCategory),
        color: mapCategoryPalette[primaryCategory] || "#53627b",
        values,
        dimensionNotes,
        classroomTags: pedagogy.classroomTags,
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  return {
    generatedAt: new Date().toISOString(),
    defaultAxes: { x: "common", y: "difficulty" },
    dimensions: mapDimensions,
    categories: categories.map((category) => ({
      name: category.name,
      slug: category.slug,
      color: mapCategoryPalette[category.name] || "#53627b",
      count: category.count,
      description: category.description,
    })),
    points,
  };
}

function buildMapPage(records, categories) {
  const mapData = buildMapData(records, categories);
  const description =
    "Use the interactive LogFall map to compare logical fallacies by commonness, visibility, accidental risk, and teaching difficulty.";
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Map</strong>
    </div>

    <section class="hero fallacy-map-hero">
      <div class="hero-panel fallacy-map-hero-panel">
        <p class="eyebrow">Interactive map</p>
        <h2 class="hero-title">Plot fallacies by classroom traits.</h2>
        <p class="hero-lead">Choose any two teaching gauges, filter by category, and inspect how a fallacy behaves before opening the full LogFall entry.</p>
        <div class="hero-actions">
          <a class="button primary" href="#fallacy-map">Explore the map</a>
          <a class="button ghost" href="../fallacies/">Browse all fallacies</a>
        </div>
      </div>
      <aside class="note-panel fallacy-map-primer" aria-label="Available dimensions">
        <p class="eyebrow">Four dimensions</p>
        <h3>Use the map as a comparison lens, not a measurement instrument.</h3>
        <p>The dots use the same editorial 0-100 teaching gauges already shown on LogFall entries.</p>
        <div class="fallacy-map-dimension-list">
          ${mapDimensions
            .map(
              (dimension) =>
                `<span>${escapeHtml(dimension.label)} <small>${escapeHtml(dimension.lowLabel)} to ${escapeHtml(dimension.highLabel)}</small></span>`,
            )
            .join("")}
        </div>
      </aside>
    </section>

    <section class="detail-section fallacy-map-section" id="fallacy-map" data-logfall-map-shell>
      <div class="fallacy-map-heading">
        <div>
          <p class="eyebrow">Scatter plot</p>
          <h2>Interactive fallacy map</h2>
          <p>Hover, focus, or click a dot to inspect it. Choose any two dimensions, isolate one category, or find a specific fallacy.</p>
        </div>
        <span class="fallacy-map-count" data-map-count>Loading fallacies...</span>
      </div>

      <div class="fallacy-map-controls" aria-label="Map controls">
        <label>
          <span>Search</span>
          <input type="search" data-map-search placeholder="Find a fallacy..." />
        </label>
        <label>
          <span>X axis</span>
          <select data-map-x-axis></select>
        </label>
        <label>
          <span>Y axis</span>
          <select data-map-y-axis></select>
        </label>
        <label>
          <span>Category</span>
          <select data-map-category></select>
        </label>
        <button type="button" class="fallacy-map-reset" data-map-reset>Reset map</button>
      </div>

      <div class="fallacy-map-layout">
        <div class="fallacy-map-card">
          <svg class="fallacy-map-plot" data-map-plot role="img" aria-label="Interactive scatter plot of logical fallacies"></svg>
          <div class="fallacy-map-legend" data-map-legend aria-label="Category legend"></div>
        </div>
        <aside class="fallacy-map-detail" data-map-detail aria-live="polite"></aside>
      </div>
    </section>
  `;

  return pageShell({
    title: "Interactive Logical Fallacy Map | LogFall",
    description,
    prefix: "../",
    currentSection: "map",
    canonicalPath: "map/",
    keywords: [
      "interactive logical fallacy map",
      "logical fallacy scatter plot",
      "fallacy comparison tool",
      "critical thinking map",
    ],
    socialImageAlt: "Interactive LogFall map",
    extraHeadHtml: `
    <link rel="stylesheet" href="../map/logfall-map.css" />
    <script type="application/json" id="logfall-map-data">${safeJsonForScript(mapData)}</script>
    <script defer src="../map/logfall-map.js"></script>
    ${cloudflareWebAnalyticsTag}
    `,
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Map", path: "map/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Interactive Logical Fallacy Map",
        url: absoluteUrl("map/"),
        description,
        publisher: publisherSchema(),
        hasPart: mapData.points.map((point) => ({
          "@type": "DefinedTerm",
          name: point.name,
          url: absoluteUrl(`fallacies/${point.slug}/`),
          description: point.definition,
        })),
      },
      learningResourceSchema({
        name: "Interactive Logical Fallacy Map",
        path: "map/",
        description,
        about: ["logical fallacies", "critical thinking", "argument comparison"],
        teaches: [
          "how common fallacies are in modern rhetoric",
          "which fallacies are easier to spot",
          "which fallacies are easier to commit accidentally",
        ],
        learningResourceType: ["Interactive map", "Reference", "Visualization"],
        educationalUse: ["teaching", "self-study", "comparison"],
        keywords: [
          "interactive logical fallacy map",
          "fallacy comparison tool",
          "critical thinking visualization",
        ],
      }),
    ],
    content,
  });
}

function buildAllFallaciesPage(records, categories, posterAssets) {
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

    ${renderFamilyGuide(buildFamilyProfiles(records))}

    <section class="section-block">
      <div class="fallacy-grid" data-fallacy-grid>
        ${records.map((record) => renderFallacyCard(record, "../", { showPoster: true, posterAssets })).join("")}
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

function buildFamiliesIndexPage(familyProfiles) {
  const content = `
    <div class="breadcrumbs">
      <a href="../">Home</a><span>/</span><strong>Families</strong>
    </div>

    ${renderFamilyGuide(familyProfiles, "")}
  `;

  return pageShell({
    title: "Logical Fallacy Families and Member Lists | LogFall",
    description: "Browse LogFall families of logical fallacies, from formal and causal mistakes to evidential, linguistic, and persuasive ones.",
    prefix: "../",
    currentSection: "fallacies",
    canonicalPath: "families/",
    keywords: [
      "logical fallacy families",
      "fallacy families",
      "types of fallacy families",
      "reasoning error families",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Families", path: "families/" },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: "Logical Fallacy Families",
        url: absoluteUrl("families/"),
        description:
          "Browse LogFall families of logical fallacies, from formal and causal mistakes to evidential, linguistic, and persuasive ones.",
        publisher: publisherSchema(),
        hasPart: familyProfiles.map((profile) => ({
          "@type": "CreativeWork",
          name: profile.name,
          url: absoluteUrl(`families/${profile.slug}/`),
          description: profile.description,
        })),
      },
      learningResourceSchema({
        name: "Logical Fallacy Families",
        path: "families/",
        description:
          "Browse LogFall families of logical fallacies, from formal and causal mistakes to evidential, linguistic, and persuasive ones.",
        about: ["logical fallacies", "fallacy families", "critical thinking"],
        teaches: ["families of logical fallacies", "comparison of reasoning mistakes"],
        learningResourceType: ["Taxonomy", "Reference"],
        educationalUse: ["teaching", "self-study"],
        keywords: ["logical fallacy families", "fallacy families", "reasoning error families"],
      }),
    ],
    content,
  });
}

function buildFamilyPage(familyProfile, records, posterAssets) {
  const members = records.filter((record) => record.family === familyProfile.name);
  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Families</a><span>/</span><strong>${escapeHtml(familyProfile.name)}</strong>
    </div>

    <section class="detail-section">
      <p class="eyebrow">Family</p>
      <h2 class="detail-title">${escapeHtml(familyProfile.name)}</h2>
      <p class="detail-deck">${escapeHtml(familyProfile.description)}</p>
      <div class="meta-grid section-block">
        <div class="note-panel">
          <h4>Entries</h4>
          <p class="muted">${familyProfile.count} fallacies in this family.</p>
        </div>
        <div class="note-panel">
          <h4>Quick family question</h4>
          <p class="muted">${escapeHtml(familyPromptForName(familyProfile.name))}</p>
        </div>
        <div class="note-panel">
          <h4>Family vs. category</h4>
          <p class="muted">A family is the broad umbrella that gives a fallacy its main home. Categories are the narrower diagnostic tags, so the same fallacy can appear in multiple categories while still belonging to one family.</p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="fallacy-grid">
        ${members.map((record) => renderFallacyCard(record, "../../", { showPoster: true, posterAssets })).join("")}
      </div>
    </section>
  `;

  return pageShell({
    title: `${familyProfile.name}: Logical Fallacy Family | LogFall`,
    description: `Browse all ${familyProfile.count} members of the ${familyProfile.name} family, with definitions, examples, and related reasoning mistakes.`,
    prefix: "../../",
    currentSection: "fallacies",
    canonicalPath: `families/${familyProfile.slug}/`,
    keywords: [
      `${familyProfile.name} logical fallacies`,
      `${familyProfile.name.toLowerCase()} family`,
      "logical fallacy family",
      "critical thinking",
    ],
    structuredData: [
      breadcrumbSchema([
        { name: "Home", path: "" },
        { name: "Families", path: "families/" },
        { name: familyProfile.name, path: `families/${familyProfile.slug}/` },
      ]),
      {
        "@context": "https://schema.org",
        "@type": "CollectionPage",
        name: `${familyProfile.name} Logical Fallacies`,
        url: absoluteUrl(`families/${familyProfile.slug}/`),
        description: `Browse all ${familyProfile.count} members of the ${familyProfile.name} family, with definitions, examples, and related reasoning mistakes.`,
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
        name: `${familyProfile.name} Logical Fallacies`,
        path: `families/${familyProfile.slug}/`,
        description: `Browse all ${familyProfile.count} members of the ${familyProfile.name} family, with definitions, examples, and related reasoning mistakes.`,
        about: [familyProfile.name, "logical fallacies", "critical thinking"],
        teaches: [`${familyProfile.name} logical fallacies`, "comparison of reasoning mistakes"],
        learningResourceType: ["Family page", "Reference"],
        educationalUse: ["teaching", "self-study"],
        keywords: [`${familyProfile.name} logical fallacies`, `${familyProfile.name.toLowerCase()} family`, "critical thinking"],
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
          <p class="section-copy">Categories are diagnostic tags for the main way reasoning goes wrong, not topic labels or ideologies. Unlike families, they are not exclusive, so one fallacy can sit in several categories at once.</p>
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

function buildCategoryPage(category, records, posterAssets) {
  const members = records.filter((record) => record.categories.includes(category.name));
  const featuredMembers = members
    .slice(0, 5)
    .map(
      (record) =>
        `<a class="path-link-chip" href="../../fallacies/${record.slug}/">${escapeHtml(record.name)}</a>`,
    )
    .join("");
  const overflowCount = Math.max(members.length - 5, 0);
  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">Categories</a><span>/</span><strong>${escapeHtml(category.name)}</strong>
    </div>

    <section class="detail-section category-hero-panel">
      <p class="eyebrow">Category</p>
      <h2 class="detail-title">${escapeHtml(category.name)}</h2>
      <p class="detail-deck">${escapeHtml(category.description)}</p>
      <div class="category-primer-grid section-block">
        <div class="note-panel category-primer-main">
          <div class="category-primer-top">
            <div>
              <h4>How to spot it</h4>
              <p class="category-primer-stat">${members.length} fallacies in this category.</p>
            </div>
            <span class="category-count-badge" aria-label="${members.length} entries">${members.length}</span>
          </div>
          <p class="muted category-primer-question">${escapeHtml(diagnosticPrompts[category.name])}</p>
          <div class="category-member-preview">
            <p class="category-member-label">Examples in this category</p>
            <div class="path-link-row">${featuredMembers}</div>
            ${overflowCount ? `<p class="category-member-note">+${overflowCount} more listed below.</p>` : ""}
          </div>
        </div>
        <div class="note-panel category-primer-side">
          <h4>Category vs. family</h4>
          <p class="muted">A category is a diagnostic lens, so a fallacy may appear in more than one category. A family is the broader umbrella that gives the fallacy its single main home.</p>
        </div>
      </div>
    </section>

    <section class="section-block">
      <div class="fallacy-grid">
        ${members.map((record) => renderFallacyCard(record, "../../", { showPoster: true, posterAssets })).join("")}
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

function buildDetailPage(record, records, categoryProfiles, posterAssets, byteseismicCrossrefs) {
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
        ${renderRhetoricGaugeSection(record, "../../")}${profileReferenceMarkup}
      </article>

      ${
        hasPosterIllustration
          ? renderPosterIllustration(record, "../../", posterAssets)
          : `<aside class="detail-section">
        <p class="eyebrow">Reference</p>
        <div class="meta-grid reference-meta-grid">
          ${renderFamilyPanel(record, "../../")}
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

    ${renderAnalogyRebuttalSection(record)}

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

    ${renderByteseismicCrossrefsSection(record, byteseismicCrossrefs)}

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
    ["Gauge scale", "The four gauge columns use 0-100 editorial teaching estimates rather than measured statistics."],
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
    "Difficulty Gauge (0-100)",
    "Often Confused With",
    "Definition",
    "Example",
    "Notes",
    "Caveat",
    "That's Like Saying Claim",
    "That's Like Saying Response",
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
      gauges.difficulty.value,
      confusions,
      record.definition,
      record.example,
      record.notes,
      caveatTextForWorkbook(record, records),
      analogyClaimForRecord(record),
      analogyResponseForRecord(record),
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
  fallaciesSheet.getRange("O:R").format.columnWidthPx = 150;
  fallaciesSheet.getRange("S:S").format.columnWidthPx = 220;
  fallaciesSheet.getRange("T:Y").format.columnWidthPx = 420;
  fallaciesSheet.getRange("Z:AG").format.columnWidthPx = 360;
  fallaciesSheet.getRange("AH:AL").format.columnWidthPx = 360;
  fallaciesSheet.getRange("AM:AM").format.columnWidthPx = 180;
  fallaciesSheet.getRange(`A1:${columnLetter(headers.length)}1`).format.wrapText = true;
  fallaciesSheet.getRange(`J:${columnLetter(headers.length)}`).format.wrapText = true;

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
  const byteseismicCrossrefsPayload = JSON.parse(
    await fs.readFile(byteseismicCrossrefsPath, "utf8").catch(() => "[]"),
  );
  const posterCaptionsPayload = JSON.parse(
    await fs.readFile(posterCaptionsPath, "utf8").catch(() => '{"captions": {}}'),
  );
  posterCaptionOverrides = posterCaptionsPayload.captions || {};
  const records = payload.records.map((record) => ({
    ...record,
    categories: normalizeRecordCategories(record),
  }));
  const byteseismicCrossrefs = normalizeByteseismicCrossrefs(byteseismicCrossrefsPayload);
  const categoryProfiles = payload.categoryProfiles || {};
  const categories = payload.categories.map((category) => ({
    ...category,
    description: categoryDescriptions[category.name] || "A reasoning category in the LogFall taxonomy.",
  }));
  const familyProfiles = buildFamilyProfiles(records);
  const posterAssets = new Set(
    (await fs.readdir(path.join(distRoot, "assets")).catch(() => []))
      .filter((name) => /^fallacy-.*-poster\.(webp|png|jpe?g)$/i.test(name)),
  );

  await pruneGeneratedDirectories(path.join(distRoot, "fallacies"), new Set(records.map((record) => record.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "categories"), new Set(categories.map((category) => category.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "families"), new Set(familyProfiles.map((profile) => profile.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "paths"), new Set(teachingPathDefinitions.map((pathDefinition) => pathDefinition.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "features"), new Set(featureArticleDefinitions.map((article) => article.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "theory"), new Set(theoryArticleDefinitions.map((article) => article.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "check-yourself"), new Set());

  await fs.copyFile(path.join(siteRoot, "styles.css"), path.join(distRoot, "styles.css"));
  await fs.copyFile(path.join(siteRoot, "app.js"), path.join(distRoot, "app.js"));
  await fs.mkdir(path.join(distRoot, "map"), { recursive: true });
  await fs.copyFile(path.join(siteRoot, "logfall-map.css"), path.join(distRoot, "map", "logfall-map.css"));
  await fs.copyFile(path.join(siteRoot, "logfall-map.js"), path.join(distRoot, "map", "logfall-map.js"));

  await writeText("index.html", buildHomePage(records, categories));
  await writeText("about/index.html", buildAboutPage());
  await writeText("check-yourself/index.html", buildAssessmentIndexPage(records, categories));
  await writeText("assessment/index.html", buildDialogueAssessmentIndexPage());
  await writeText("map/index.html", buildMapPage(records, categories));
  await writeText("features/index.html", buildFeaturesIndexPage());
  await writeText("prompts/index.html", buildPromptsPage());
  await writeText("theory/index.html", buildTheoryIndexPage());
  await writeText("fallacies/index.html", buildAllFallaciesPage(records, categories, posterAssets));
  await writeText("categories/index.html", buildCategoriesIndexPage(categories));
  await writeText("families/index.html", buildFamiliesIndexPage(familyProfiles));
  await writeText("paths/index.html", buildTeachingPathsIndexPage(records));
  await writeText("404.html", build404Page());
  const sitemapEntries = [
    { path: "" },
    { path: "about/" },
    { path: "check-yourself/" },
    { path: "assessment/" },
    { path: "map/" },
    { path: "features/" },
    { path: "prompts/" },
    { path: "theory/" },
    { path: "fallacies/" },
    { path: "categories/" },
    { path: "families/" },
    { path: "paths/" },
    ...categories.map((category) => ({ path: `categories/${category.slug}/` })),
    ...familyProfiles.map((profile) => ({ path: `families/${profile.slug}/` })),
    ...teachingPathDefinitions.map((pathDefinition) => ({ path: `paths/${pathDefinition.slug}/` })),
    ...featureArticleDefinitions.map((article) => ({ path: `features/${article.slug}/` })),
    ...theoryArticleDefinitions.map((article) => ({ path: `theory/${article.slug}/` })),
    ...records.map((record) => ({ path: `fallacies/${record.slug}/` })),
  ];
  await writeText("sitemap.xml", buildSitemap(sitemapEntries));
  await writeText("robots.txt", `User-agent: *\nAllow: /\nSitemap: ${absoluteUrl("sitemap.xml")}\n`);

  for (const category of categories) {
    await writeText(`categories/${category.slug}/index.html`, buildCategoryPage(category, records, posterAssets));
  }

  for (const familyProfile of familyProfiles) {
    await writeText(`families/${familyProfile.slug}/index.html`, buildFamilyPage(familyProfile, records, posterAssets));
  }

  for (const pathDefinition of teachingPathDefinitions) {
    await writeText(`paths/${pathDefinition.slug}/index.html`, buildTeachingPathPage(pathDefinition, records));
  }

  for (const article of featureArticleDefinitions) {
    await writeText(`features/${article.slug}/index.html`, buildFeatureArticlePage(article));
  }

  for (const article of theoryArticleDefinitions) {
    await writeText(`theory/${article.slug}/index.html`, buildTheoryArticlePage(article));
  }

  for (const record of records) {
    await writeText(
      `fallacies/${record.slug}/index.html`,
      buildDetailPage(record, records, categoryProfiles, posterAssets, byteseismicCrossrefs),
    );
  }

  await buildWorkbook(records, categories, categoryProfiles);

  console.log(
    JSON.stringify(
      {
        distRoot,
        pageCount:
          11 + categories.length + familyProfiles.length + teachingPathDefinitions.length + featureArticleDefinitions.length + theoryArticleDefinitions.length + records.length,
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
