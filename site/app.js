const searchInput = document.querySelector("[data-search-input]");
const categoryFilter = document.querySelector("[data-category-filter]");
const cards = Array.from(document.querySelectorAll("[data-fallacy-card]"));
const countNode = document.querySelector("[data-search-count]");

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
    countNode.textContent = `${visible} fallac${visible === 1 ? "y" : "ies"} shown`;
  }
}

if (searchInput) {
  searchInput.addEventListener("input", applyFilters);
}

if (categoryFilter) {
  categoryFilter.addEventListener("change", applyFilters);
}

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
