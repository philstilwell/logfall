const searchInput = document.querySelector("[data-search-input]");
const categoryFilter = document.querySelector("[data-category-filter]");
const resetButton = document.querySelector("[data-search-reset]");
const cards = Array.from(document.querySelectorAll("[data-fallacy-card]"));
const countNode = document.querySelector("[data-search-count]");
const emptyState = document.querySelector("[data-search-empty]");
const totalCount = cards.length;

function syncFilterUrl(query, category) {
  if (!searchInput && !categoryFilter) return;
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
  window.history.replaceState({}, "", url);
}

function hydrateFiltersFromUrl() {
  const url = new URL(window.location.href);
  const query = url.searchParams.get("q") || "";
  const category = url.searchParams.get("category") || "";
  if (searchInput) searchInput.value = query;
  if (categoryFilter) categoryFilter.value = category;
}

function applyFilters() {
  if (!cards.length) return;

  const query = (searchInput?.value || "").trim().toLowerCase();
  const category = categoryFilter?.value || "";
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
    const show = matchesQuery && matchesCategory;
    card.classList.toggle("hidden", !show);
    if (show) visible += 1;
  }

  if (countNode) {
    countNode.textContent = `${visible} of ${totalCount} fallac${totalCount === 1 ? "y" : "ies"} shown`;
  }

  if (emptyState) {
    emptyState.classList.toggle("hidden", visible !== 0);
  }

  syncFilterUrl(query, category);
}

if (searchInput) {
  searchInput.addEventListener("input", applyFilters);
}

if (categoryFilter) {
  categoryFilter.addEventListener("change", applyFilters);
}

if (resetButton) {
  resetButton.addEventListener("click", () => {
    if (searchInput) searchInput.value = "";
    if (categoryFilter) categoryFilter.value = "";
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
