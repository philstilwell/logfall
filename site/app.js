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

  function activateTab(index) {
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
  }

  buttons.forEach((button, index) => {
    button.addEventListener("click", () => activateTab(index));
  });

  activateTab(0);
}

for (const lab of document.querySelectorAll("[data-confidence-lab]")) {
  const surface = lab.querySelector("[data-lab-surface]");
  const evidence = lab.querySelector("[data-lab-evidence]");
  const surfaceValue = lab.querySelector("[data-lab-surface-value]");
  const evidenceValue = lab.querySelector("[data-lab-evidence-value]");
  const gapFill = lab.querySelector("[data-lab-gap-fill]");
  const gapOutput = lab.querySelector("[data-lab-gap-output]");

  if (!surface || !evidence || !surfaceValue || !evidenceValue || !gapFill || !gapOutput) {
    continue;
  }

  function updateLab() {
    const surfaceScore = Number(surface.value || 0);
    const evidenceScore = Number(evidence.value || 0);
    const gap = surfaceScore - evidenceScore;

    surfaceValue.textContent = String(surfaceScore);
    evidenceValue.textContent = String(evidenceScore);
    gapFill.style.width = `${Math.abs(gap)}%`;

    if (gap > 0) {
      gapOutput.textContent = `Confidence outruns support by ${gap} points. Slow down and inspect what is doing the persuasive work.`;
      gapFill.style.background = "linear-gradient(90deg, #e63b34, #14a8d7)";
    } else if (gap < 0) {
      gapOutput.textContent = `Support exceeds surface pull by ${Math.abs(gap)} points. The argument may feel weaker than the evidence actually warrants.`;
      gapFill.style.background = "linear-gradient(90deg, #14a8d7, #e63b34)";
    } else {
      gapOutput.textContent = "Confidence and support are aligned. Keep testing whether the match survives closer scrutiny.";
      gapFill.style.background = "linear-gradient(90deg, #14a8d7, #e63b34)";
    }
  }

  surface.addEventListener("input", updateLab);
  evidence.addEventListener("input", updateLab);
  updateLab();
}
