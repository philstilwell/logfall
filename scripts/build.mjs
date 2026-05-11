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
const siteUrl = "https://philstilwell.github.io/logfall/";
const socialImagePath = "assets/logo.jpg";
const buildDate = new Date().toISOString().split("T")[0];
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

function sortCategories(categories) {
  const order = [
    "Formal",
    "Mathematical",
    "Causal",
    "Linguistic",
    "Conceptual",
    "Evidential",
    "Perceptual",
    "Perspectival",
    "Epistemic",
    "Tactical",
    "Emotional",
  ];
  return [...categories].sort((a, b) => order.indexOf(a) - order.indexOf(b));
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
  showSpreadsheetNav = true,
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
  if (showSpreadsheetNav) {
    navItems.push({ href: `${prefix}logfall-root-edition.xlsx`, label: "Spreadsheet", key: "spreadsheet" });
  }

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
                <p class="brand-subtitle">A living logical-fallacies reference rebuilt for clearer explanations, faster browsing, and steadily better examples.</p>
              </div>
            </div>
          </div>
          <nav class="top-nav" aria-label="Primary">${nav}</nav>
        </div>
      </header>
      <main class="page-wrap">${content}</main>
      <footer class="footer">
        <div class="footer-inner">
          Built from the ROOT tab source workbook, filtered to entries that existed on the legacy WordPress LogFall site.
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

function renderFallacyCard(record, prefix) {
  const aliases = record.aliases.join(" ");
  const body = `${record.definition} ${record.notes}`;
  return `<article
    class="fallacy-card"
    data-fallacy-card
    data-name="${escapeHtml(record.name)}"
    data-aliases="${escapeHtml(aliases)}"
    data-categories="${escapeHtml(record.categories.join("|"))}"
    data-body="${escapeHtml(body)}"
  >
    <h3><a href="${prefix}fallacies/${record.slug}/">${escapeHtml(record.name)}</a></h3>
    <p class="card-copy">${escapeHtml(truncate(record.definition, 170))}</p>
    ${renderPills(record.categories)}
  </article>`;
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

function buildHomePage(records, categories) {
  const featured = featuredNames
    .map((name) => records.find((record) => record.name === name))
    .filter(Boolean);

  const content = `
    <section class="hero">
      <div class="hero-panel">
        <h2 class="hero-title">Logical fallacies, reorganized for modern readers.</h2>
        <p class="hero-lead">
          This GitHub Pages edition keeps the spirit of the original LogFall project while making the taxonomy easier to browse,
          compare, search, and revise. It is built from the ROOT tab source, limited to fallacies that appeared on the legacy site,
          and structured for ongoing editorial improvement.
        </p>
        <div class="hero-actions">
          <a class="button button-primary" href="fallacies/">Browse All Fallacies</a>
          <a class="button button-secondary" href="categories/">Explore Categories</a>
          <a class="button button-secondary" href="logfall-root-edition.xlsx">Download Workbook</a>
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
            <span class="stat-value">1</span>
            <span class="stat-label">Unified source workbook</span>
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
            <p class="muted">Each fallacy page brings together a concise definition, a concrete example, explanatory notes, case studies, and nearby entries.</p>
          </div>
          <div class="note-panel" style="margin-top:12px;">
            <h4>3. Use the workbook</h4>
            <p class="muted">The downloadable spreadsheet mirrors the site data so the project can be revised in a structured way.</p>
          </div>
        </div>
      </aside>
    </section>

    <section class="section-block">
      <div class="section-header">
        <div>
          <h2 class="section-title">Browse the taxonomy</h2>
          <p class="section-copy">The original project grouped fallacies by the kind of reasoning failure involved. This edition keeps that structure and makes the overlap clearer.</p>
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
      "A rebuilt logical fallacies reference with category browsing, clearer explanations, and modernized examples.",
    prefix: "",
    currentSection: "home",
    showSpreadsheetNav: false,
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
      <p class="eyebrow">About this edition</p>
      <h2 class="detail-title">A cleaner static edition of LogFall.</h2>
      <p class="detail-deck">
        This version is designed for GitHub Pages. It preserves the original logo and taxonomy approach,
        but replaces WordPress-era browsing friction with lightweight static pages, clearer navigation,
        and a synchronized workbook for structured revision.
      </p>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>Source rule</h4>
          <p class="muted">The ROOT tab provides the base material, which can then be cleaned or rewritten for clarity, accuracy, and more useful examples.</p>
        </div>
        <div class="note-panel">
          <h4>Inclusion rule</h4>
          <p class="muted">The new site includes only fallacies that also appeared on the legacy WordPress LogFall site.</p>
        </div>
      </div>
      <div class="two-column section-block">
        <div class="note-panel">
          <h4>Design goal</h4>
          <p class="muted">Make the taxonomy easier to scan, search, and compare while preserving the recognizable red-and-cyan LogFall identity.</p>
        </div>
        <div class="note-panel">
          <h4>Editorial direction</h4>
          <p class="muted">This edition is being actively rewritten, starting with the most-used entries and replacing weaker examples with clearer modern case studies.</p>
        </div>
      </div>
    </section>
  `;

  return pageShell({
    title: "About | LogFall",
    description: "About the rebuilt LogFall GitHub Pages edition and its source rules.",
    prefix: "../",
    currentSection: "about",
    canonicalPath: "about/",
    content,
  });
}

function buildAllFallaciesPage(records, categories) {
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
          <p class="section-copy">Search by name, alias, definition text, or filter the list by category.</p>
        </div>
      </div>
      <div class="search-row">
        <input class="search-input" type="search" placeholder="Search fallacies, aliases, or keywords..." data-search-input />
        <select class="search-select" data-category-filter>
          <option value="">All categories</option>
          ${options}
        </select>
      </div>
      <div class="search-meta" data-search-count>${records.length} fallacies shown</div>
    </section>

    <section class="section-block">
      <div class="fallacy-grid">
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

function buildDetailPage(record, records) {
  const related = relatedFallacies(record, records);
  const prompts = record.categories.map((category) => diagnosticPrompts[category]).filter(Boolean);

  const content = `
    <div class="breadcrumbs">
      <a href="../../">Home</a><span>/</span><a href="../">All Fallacies</a><span>/</span><strong>${escapeHtml(record.name)}</strong>
    </div>

    <section class="detail-hero">
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
      </article>

      <aside class="detail-section">
        <p class="eyebrow">Reference</p>
        <div class="meta-grid">
          <div class="note-panel">
            <h4>Legacy number</h4>
            <p class="muted">${escapeHtml(record.originalNumber || "Not listed")}</p>
          </div>
          <div class="note-panel">
            <h4>Family</h4>
            <p class="muted">${escapeHtml(record.family || "Unspecified")}</p>
          </div>
          <div class="note-panel">
            <h4>Aliases</h4>
            <p class="muted">${escapeHtml(record.aliases.length ? record.aliases.join(", ") : "None listed")}</p>
          </div>
          <div class="note-panel">
            <h4>Quick check</h4>
            <p class="muted">${escapeHtml(prompts[0] || "Ask what evidence or reasoning step is doing too much work.")}</p>
          </div>
        </div>
      </aside>
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
          <p class="section-copy">Recent or representative examples that help stress-test the pattern in practice.</p>
        </div>
      </div>
      <div class="case-list">
        ${record.caseStudies.map((item) => `<blockquote class="case-item">${escapeHtml(item)}</blockquote>`).join("")}
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
      <div class="fallacy-grid">
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

async function buildWorkbook(records, categories) {
  const workbook = Workbook.create();

  const overview = workbook.worksheets.add("Overview");
  const overviewRows = [
    ["LogFall Workbook", "Rebuilt from the ROOT tab source workbook."],
    ["Record count", records.length],
    ["Category count", categories.length],
    ["Inclusion rule", "Only fallacies that also existed on the legacy WordPress site are included."],
    ["Workbook use", "Use the Fallacies sheet for entry-by-entry editing and the Categories sheet for counts."],
  ];
  overview.getRange(`A1:B${overviewRows.length}`).values = overviewRows;

  const fallaciesSheet = workbook.worksheets.add("Fallacies");
  const headers = [
    "Name",
    "Slug",
    "Primary Category",
    "Additional Categories",
    "Original Number",
    "Family",
    "Sub-Category",
    "Sub-Sub-Category",
    "Aliases",
    "Definition",
    "Example",
    "Notes",
    "Case Study 1",
    "Case Study 2",
    "Case Study 3",
    "Case Study 4",
    "Case Study 5",
    "Editorial Status",
  ];
  const rows = records.map((record) => [
    record.name,
    record.slug,
    record.categories[0] || "",
    record.categories.slice(1).join(", "),
    record.originalNumber,
    record.family,
    record.subCategory,
    record.subSubCategory,
    record.aliases.join(", "),
    record.definition,
    record.example,
    record.notes,
    record.caseStudies[0] || "",
    record.caseStudies[1] || "",
    record.caseStudies[2] || "",
    record.caseStudies[3] || "",
    record.caseStudies[4] || "",
    record.editorialStatus,
  ]);
  const fallaciesMatrix = [headers, ...rows];
  fallaciesSheet.getRange(`A1:${columnLetter(headers.length)}${fallaciesMatrix.length}`).values =
    fallaciesMatrix;

  const categorySheet = workbook.worksheets.add("Categories");
  const categoryRows = [["Category", "Count", "Description"], ...categories.map((category) => [category.name, category.count, category.description])];
  categorySheet.getRange(`A1:C${categoryRows.length}`).values = categoryRows;

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
    categories: sortCategories(record.categories),
  }));
  const categories = payload.categories.map((category) => ({
    ...category,
    description: categoryDescriptions[category.name] || "A reasoning category in the LogFall taxonomy.",
  }));

  await pruneGeneratedDirectories(path.join(distRoot, "fallacies"), new Set(records.map((record) => record.slug)));
  await pruneGeneratedDirectories(path.join(distRoot, "categories"), new Set(categories.map((category) => category.slug)));

  await fs.copyFile(path.join(siteRoot, "styles.css"), path.join(distRoot, "styles.css"));
  await fs.copyFile(path.join(siteRoot, "app.js"), path.join(distRoot, "app.js"));

  await writeText("index.html", buildHomePage(records, categories));
  await writeText("about/index.html", buildAboutPage());
  await writeText("fallacies/index.html", buildAllFallaciesPage(records, categories));
  await writeText("categories/index.html", buildCategoriesIndexPage(categories));
  await writeText("404.html", build404Page());
  const sitemapEntries = [
    { path: "" },
    { path: "about/" },
    { path: "fallacies/" },
    { path: "categories/" },
    ...categories.map((category) => ({ path: `categories/${category.slug}/` })),
    ...records.map((record) => ({ path: `fallacies/${record.slug}/` })),
  ];
  await writeText("sitemap.xml", buildSitemap(sitemapEntries));
  await writeText("robots.txt", `User-agent: *\nAllow: /\nSitemap: ${absoluteUrl("sitemap.xml")}\n`);

  for (const category of categories) {
    await writeText(`categories/${category.slug}/index.html`, buildCategoryPage(category, records));
  }

  for (const record of records) {
    await writeText(`fallacies/${record.slug}/index.html`, buildDetailPage(record, records));
  }

  await buildWorkbook(records, categories);

  console.log(
    JSON.stringify(
      {
        distRoot,
        pageCount: 5 + categories.length + records.length,
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
