const searchInput = document.querySelector("[data-search-input]");
const categoryFilter = document.querySelector("[data-category-filter]");
const difficultyFilter = document.querySelector("[data-difficulty-filter]");
const classroomFilter = document.querySelector("[data-classroom-filter]");
const resetButton = document.querySelector("[data-search-reset]");
const cards = Array.from(document.querySelectorAll("[data-fallacy-card]"));
const countNode = document.querySelector("[data-search-count]");
const emptyState = document.querySelector("[data-search-empty]");
const totalCount = cards.length;

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
    countNode.textContent = `${visible} of ${totalCount} fallac${totalCount === 1 ? "y" : "ies"} shown`;
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
