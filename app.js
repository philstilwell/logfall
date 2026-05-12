const searchInput = document.querySelector("[data-search-input]");
const categoryFilter = document.querySelector("[data-category-filter]");
const difficultyFilter = document.querySelector("[data-difficulty-filter]");
const classroomFilter = document.querySelector("[data-classroom-filter]");
const resetButton = document.querySelector("[data-search-reset]");
const cards = Array.from(document.querySelectorAll("[data-fallacy-card]"));
const countNode = document.querySelector("[data-search-count]");
const emptyState = document.querySelector("[data-search-empty]");
const totalCount = cards.length;
const countSingular = countNode?.dataset.searchUnitSingular || "fallacy";
const countPlural = countNode?.dataset.searchUnitPlural || "fallacies";

function escapeHtmlText(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function syncFilterUrl(query, category, difficulty, classroom) {
  if (!searchInput && !categoryFilter && !difficultyFilter && !classroomFilter) return;
  const url = new URL(window.location.href);
  if (query) {
    url.searchParams.set("q", query);
  } else {
    url.searchParams.delete("q");
  }
  if (category) {
    url.searchParams.set("category", category);
  } else {
    url.searchParams.delete("category");
  }
  if (difficulty) {
    url.searchParams.set("difficulty", difficulty);
  } else {
    url.searchParams.delete("difficulty");
  }
  if (classroom) {
    url.searchParams.set("classroom", classroom);
  } else {
    url.searchParams.delete("classroom");
  }
  window.history.replaceState({}, "", url);
}

function hydrateFiltersFromUrl() {
  const url = new URL(window.location.href);
  const query = url.searchParams.get("q") || "";
  const category = url.searchParams.get("category") || "";
  const difficulty = url.searchParams.get("difficulty") || "";
  const classroom = url.searchParams.get("classroom") || "";
  if (searchInput) searchInput.value = query;
  if (categoryFilter) categoryFilter.value = category;
  if (difficultyFilter) difficultyFilter.value = difficulty;
  if (classroomFilter) classroomFilter.value = classroom;
}

function applyFilters() {
  if (!cards.length) return;

  const query = (searchInput?.value || "").trim().toLowerCase();
  const category = categoryFilter?.value || "";
  const difficulty = difficultyFilter?.value || "";
  const classroom = classroomFilter?.value || "";
  let visible = 0;

  for (const card of cards) {
    const haystack = [
      card.dataset.name || "",
      card.dataset.aliases || "",
      card.dataset.categories || "",
      card.dataset.body || "",
    ]
      .join(" ")
      .toLowerCase();

    const matchesQuery = !query || haystack.includes(query);
    const matchesCategory = !category || (card.dataset.categories || "").includes(category);
    const matchesDifficulty = !difficulty || (card.dataset.difficulty || "") === difficulty;
    const matchesClassroom = !classroom || (card.dataset.classroom || "") === classroom;
    const show = matchesQuery && matchesCategory && matchesDifficulty && matchesClassroom;
    card.classList.toggle("hidden", !show);
    if (show) visible += 1;
  }

  if (countNode) {
    const unitLabel = totalCount === 1 ? countSingular : countPlural;
    countNode.textContent = `${visible} of ${totalCount} ${unitLabel} shown`;
  }

  if (emptyState) {
    emptyState.classList.toggle("hidden", visible !== 0);
  }

  syncFilterUrl(query, category, difficulty, classroom);
}

if (searchInput) {
  searchInput.addEventListener("input", applyFilters);
}

if (categoryFilter) {
  categoryFilter.addEventListener("change", applyFilters);
}

if (difficultyFilter) {
  difficultyFilter.addEventListener("change", applyFilters);
}

if (classroomFilter) {
  classroomFilter.addEventListener("change", applyFilters);
}

if (resetButton) {
  resetButton.addEventListener("click", () => {
    if (searchInput) searchInput.value = "";
    if (categoryFilter) categoryFilter.value = "";
    if (difficultyFilter) difficultyFilter.value = "";
    if (classroomFilter) classroomFilter.value = "";
    applyFilters();
    searchInput?.focus();
  });
}

hydrateFiltersFromUrl();
applyFilters();

for (const group of document.querySelectorAll("[data-tab-group]")) {
  const buttons = Array.from(group.querySelectorAll("[data-tab-button]"));
  const panels = Array.from(group.querySelectorAll("[data-tab-panel]"));

  function activateTab(index, options = {}) {
    buttons.forEach((button, buttonIndex) => {
      const active = buttonIndex === index;
      button.classList.toggle("active", active);
      button.setAttribute("aria-selected", active ? "true" : "false");
    });

    panels.forEach((panel, panelIndex) => {
      const active = panelIndex === index;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
    });

    if (options.scroll) {
      panels[index]?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  buttons.forEach((button, index) => {
    button.addEventListener("click", () => activateTab(index, { scroll: true }));
  });

  activateTab(0);
}

for (const widget of document.querySelectorAll("[data-audit-widget]")) {
  const buttons = Array.from(widget.querySelectorAll("[data-audit-button]"));
  const panels = Array.from(widget.querySelectorAll("[data-audit-panel]"));
  const progress = widget.querySelector("[data-audit-progress]");

  if (!buttons.length || !panels.length || !progress) {
    continue;
  }

  function activateAudit(index) {
    buttons.forEach((button, buttonIndex) => {
      const active = buttonIndex === index;
      button.classList.toggle("active", active);
      button.setAttribute("aria-selected", active ? "true" : "false");
    });

    panels.forEach((panel, panelIndex) => {
      const active = panelIndex === index;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
    });

    const denominator = Math.max(buttons.length - 1, 1);
    const percent = (index / denominator) * 100;
    progress.style.width = `${percent}%`;
  }

  buttons.forEach((button, index) => {
    button.addEventListener("click", () => activateAudit(index));
  });

  activateAudit(0);
}

for (const quiz of document.querySelectorAll("[data-quiz-widget]")) {
  const gradeButton = quiz.querySelector("[data-quiz-grade]");
  const responseField = quiz.querySelector("[data-quiz-response]");
  const feedback = quiz.querySelector("[data-quiz-feedback]");
  const answer = quiz.dataset.quizAnswer || "";
  const model = quiz.dataset.quizModel || "";
  let keywords = [];

  try {
    keywords = JSON.parse(quiz.dataset.quizKeywords || "[]");
  } catch {
    keywords = [];
  }

  if (!gradeButton || !responseField || !feedback) {
    continue;
  }

  gradeButton.addEventListener("click", () => {
    const selected = quiz.querySelector('input[type="radio"]:checked');
    const answerScore = selected?.value === answer ? 1 : 0;
    const response = (responseField.value || "").trim().toLowerCase();
    const matched = keywords.filter((keyword) => response.includes(String(keyword).toLowerCase()));
    let explanationScore = 0;
    if (response.length >= 35 && matched.length >= 4) {
      explanationScore = 2;
    } else if (response.length >= 25 && matched.length >= 2) {
      explanationScore = 1;
    }

    const total = answerScore + explanationScore;
    const correctAnswerLine = answerScore
      ? `You identified the fallacy correctly.`
      : `The best label here is ${answer || "the page's target fallacy"}.`;
    const explanationLine =
      explanationScore === 2
        ? `Your explanation names the key reasoning slip clearly.`
        : explanationScore === 1
          ? `Your explanation is partly there, but it could state the reasoning slip more directly.`
          : `Your explanation needs to say more explicitly what the reasoning mistake is.`;

    feedback.innerHTML = `
      <p><strong>Score:</strong> ${total}/3</p>
      <p>${correctAnswerLine}</p>
      <p>${explanationLine}</p>
      <p><strong>Model answer:</strong> ${model}</p>
    `;
    feedback.classList.remove("hidden");
  });
}

function hashSeed(input = "") {
  let hash = 2166136261;
  for (const char of String(input)) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function seededRandomFactory(seedInput = "") {
  let state = hashSeed(seedInput) || 123456789;
  return () => {
    state += 0x6d2b79f5;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffleWithSeed(items, random) {
  const copy = [...items];
  for (let index = copy.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(random() * (index + 1));
    [copy[index], copy[swapIndex]] = [copy[swapIndex], copy[index]];
  }
  return copy;
}

function makeAssessmentSeed() {
  if (window.crypto?.getRandomValues) {
    const values = new Uint32Array(2);
    window.crypto.getRandomValues(values);
    return `${values[0].toString(36)}${values[1].toString(36)}`;
  }
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}

for (const shell of document.querySelectorAll("[data-assessment-shell]")) {
  const bankNode = shell.querySelector("#assessment-bank");
  const itemsNode = shell.querySelector("[data-assessment-items]");
  const bannerNode = shell.querySelector("[data-assessment-banner]");
  const gradeButton = shell.querySelector("[data-assessment-grade]");
  const newButton = shell.querySelector("[data-assessment-new]");
  const resultsNode = shell.querySelector("[data-assessment-results]");
  const size = Number(shell.dataset.assessmentSize || 10);

  if (!bankNode || !itemsNode || !bannerNode || !gradeButton || !newButton || !resultsNode) {
    continue;
  }

  let bank = [];
  try {
    bank = JSON.parse(bankNode.textContent || "[]");
  } catch {
    bank = [];
  }

  if (!Array.isArray(bank) || !bank.length) {
    bannerNode.innerHTML = `
      <h4>Assessment unavailable</h4>
      <p class="muted">The quiz bank could not be loaded for this page.</p>
    `;
    continue;
  }

  const url = new URL(window.location.href);
  const requestedFocus = url.searchParams.get("focus") || "";
  const focusItem = bank.find((item) => item.slug === requestedFocus) || null;
  let currentSeed = url.searchParams.get("set") || makeAssessmentSeed();
  let currentSet = [];

  function syncAssessmentUrl() {
    const nextUrl = new URL(window.location.href);
    nextUrl.searchParams.set("set", currentSeed);
    if (focusItem) {
      nextUrl.searchParams.set("focus", focusItem.slug);
    } else {
      nextUrl.searchParams.delete("focus");
    }
    window.history.replaceState({}, "", nextUrl);
  }

  function buildAssessmentSet() {
    const random = seededRandomFactory(currentSeed);
    const selected = [];
    const used = new Set();

    if (focusItem) {
      selected.push(focusItem);
      used.add(focusItem.slug);
      const samePrimary = shuffleWithSeed(
        bank.filter((item) => !used.has(item.slug) && item.categories?.[0] === focusItem.categories?.[0]),
        random,
      ).slice(0, Math.min(2, size - 1));
      for (const item of samePrimary) {
        selected.push(item);
        used.add(item.slug);
      }
    }

    const shuffled = shuffleWithSeed(
      bank.filter((item) => !used.has(item.slug)),
      random,
    );

    for (const item of shuffled) {
      if (selected.length >= size) break;
      selected.push(item);
      used.add(item.slug);
    }

    return shuffleWithSeed(selected.slice(0, size), random).map((item, index) => ({
      ...item,
      questionNumber: index + 1,
    }));
  }

  function renderAssessmentBanner() {
    if (focusItem) {
      bannerNode.innerHTML = `
        <h4>Focused mixed set</h4>
        <p class="muted">This 10-question set quietly includes one item centered on a fallacy from the page you came from, plus nearby and contrasting fallacies, so the answer is not announced in advance.</p>
      `;
      return;
    }

    bannerNode.innerHTML = `
      <h4>Mixed set</h4>
      <p class="muted">This assessment draws semirandomly from the full LogFall library. Use “Load another set” for a fresh group of 10 example claims.</p>
    `;
  }

  function renderAssessmentQuestions() {
    itemsNode.innerHTML = currentSet
      .map(
        (item, index) => `
        <article class="detail-section quiz-card assessment-question-card" data-assessment-question="${index}">
          <p class="eyebrow">Question ${item.questionNumber}</p>
          <h3 class="assessment-question-title">Identify the fallacy</h3>
          <div class="quiz-example-shell">
            <p class="quiz-example-label">Claim to diagnose</p>
            <p class="quiz-example-text">${escapeHtmlText(item.example)}</p>
            <p class="quiz-example-note">Choose the label that best fits the reasoning move in this exact claim. Do not answer from the topic alone or from whether you agree with the conclusion.</p>
          </div>
          <div class="quiz-options">
            ${item.options
              .map(
                (option) => `
                <label class="quiz-option">
                  <input type="radio" name="assessment-question-${index}" value="${escapeHtmlText(option)}" />
                  <span>${escapeHtmlText(option)}</span>
                </label>`,
              )
              .join("")}
          </div>
          <div class="quiz-feedback hidden" data-assessment-feedback role="status" aria-live="polite"></div>
        </article>`,
      )
      .join("");
  }

  function renderResults(score, answeredCount) {
    const total = currentSet.length;
    const reviewLinks = currentSet
      .map(
        (item, index) => `
          <li>
            <a class="text-link" href="${item.fallacyUrl}">Question ${index + 1}: ${escapeHtmlText(item.name)}</a>
          </li>`,
      )
      .join("");

    resultsNode.innerHTML = `
      <p class="eyebrow">Assessment results</p>
      <h3 class="assessment-score-title">Score: ${score}/${total}</h3>
      <p class="muted">You answered ${answeredCount} of ${total} questions and identified ${score} correctly. Review the per-question feedback above, then open the accordion below to study the correct fallacies from this exact set.</p>
      <details class="assessment-review-accordion">
        <summary>Review the 10 correct fallacies from this set</summary>
        <ul class="assessment-review-links">
          ${reviewLinks}
        </ul>
      </details>
    `;
    resultsNode.classList.remove("hidden");
  }

  function gradeAssessment() {
    let score = 0;
    let answeredCount = 0;
    const cards = Array.from(itemsNode.querySelectorAll("[data-assessment-question]"));

    cards.forEach((card, index) => {
      const item = currentSet[index];
      const feedback = card.querySelector("[data-assessment-feedback]");
      const selected = card.querySelector('input[type="radio"]:checked');
      const selectedValue = selected?.value || "";
      const isCorrect = selectedValue === item.answer;

      if (selectedValue) answeredCount += 1;
      if (isCorrect) score += 1;
      if (!feedback) return;

      feedback.innerHTML = `
        <p><strong>${isCorrect ? "Correct." : "Not quite."}</strong> ${
          isCorrect
            ? "You chose the best label for this claim."
            : selectedValue
              ? `You chose ${escapeHtmlText(selectedValue)}, but the best label is ${escapeHtmlText(item.answer)}.`
              : `No answer was selected. The best label is ${escapeHtmlText(item.answer)}.`
        }</p>
        <p><strong>Why this fits:</strong> ${escapeHtmlText(item.model)}</p>
        <p><a class="text-link" href="${item.fallacyUrl}">Review ${escapeHtmlText(item.name)}</a></p>
      `;
      feedback.classList.remove("hidden");
    });

    renderResults(score, answeredCount);
    resultsNode.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function loadAssessmentSet() {
    syncAssessmentUrl();
    currentSet = buildAssessmentSet();
    renderAssessmentBanner();
    renderAssessmentQuestions();
    resultsNode.classList.add("hidden");
    resultsNode.innerHTML = "";
  }

  gradeButton.addEventListener("click", gradeAssessment);
  newButton.addEventListener("click", () => {
    currentSeed = makeAssessmentSeed();
    loadAssessmentSet();
    shell.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  loadAssessmentSet();
}

for (const button of document.querySelectorAll("[data-copy-button]")) {
  button.addEventListener("click", async () => {
    const targetId = button.getAttribute("data-copy-button");
    const source = targetId ? document.getElementById(targetId) : null;
    if (!source) return;

    const text = source.value || source.textContent || "";
    let copied = false;
    let fallbackSelected = false;

    try {
      await navigator.clipboard.writeText(text);
      copied = true;
    } catch {
      if (source.select) {
        source.focus();
        source.select();
        copied = document.execCommand("copy");
        fallbackSelected = true;
      }
    }

    const original = button.textContent;
    if (copied) {
      button.textContent = "Copied";
    } else if (fallbackSelected) {
      button.textContent = "Press Cmd/Ctrl+C";
    } else {
      button.textContent = "Copy unavailable";
    }
    button.classList.add("copied");
    window.setTimeout(() => {
      button.textContent = original;
      button.classList.remove("copied");
      if (fallbackSelected) {
        source.setSelectionRange?.(0, 0);
      }
    }, 1600);
  });
}
