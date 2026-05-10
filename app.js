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
