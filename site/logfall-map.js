(function () {
  var dataNode = document.getElementById("logfall-map-data");
  var shell = document.querySelector("[data-logfall-map-shell]");
  if (!dataNode || !shell) return;

  var data = JSON.parse(dataNode.textContent || "{}");
  var dimensions = data.dimensions || [];
  var dimensionBySlug = {};
  dimensions.forEach(function (dimension) {
    dimensionBySlug[dimension.slug] = dimension;
  });

  var state = {
    x: data.defaultAxes && data.defaultAxes.x ? data.defaultAxes.x : "common",
    y: data.defaultAxes && data.defaultAxes.y ? data.defaultAxes.y : "difficulty",
    category: "all",
    query: "",
    selectedSlug: data.points && data.points.length ? data.points[0].slug : "",
  };

  var searchInput = shell.querySelector("[data-map-search]");
  var xSelect = shell.querySelector("[data-map-x-axis]");
  var ySelect = shell.querySelector("[data-map-y-axis]");
  var categorySelect = shell.querySelector("[data-map-category]");
  var resetButton = shell.querySelector("[data-map-reset]");
  var countNode = shell.querySelector("[data-map-count]");
  var plotNode = shell.querySelector("[data-map-plot]");
  var detailNode = shell.querySelector("[data-map-detail]");
  var legendNode = shell.querySelector("[data-map-legend]");

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function normalize(value) {
    return String(value || "").trim().toLowerCase();
  }

  function visiblePoints() {
    var query = normalize(state.query);
    return (data.points || []).filter(function (point) {
      var pointCategorySlugs = Array.isArray(point.categorySlugs) ? point.categorySlugs : point.categorySlug ? [point.categorySlug] : [];
      var matchesCategory = state.category === "all" || pointCategorySlugs.indexOf(state.category) !== -1;
      if (!matchesCategory) return false;
      if (!query) return true;
      var haystack = [point.name, point.definition, point.category, (point.categories || []).join(" ")].join(" ").toLowerCase();
      return haystack.indexOf(query) !== -1;
    });
  }

  function pointBySlug(slug) {
    return (data.points || []).find(function (point) {
      return point.slug === slug;
    });
  }

  function hashOffset(slug, axis) {
    var source = String(slug || "") + axis;
    var hash = 0;
    for (var index = 0; index < source.length; index += 1) {
      hash = (hash * 31 + source.charCodeAt(index)) % 997;
    }
    return (hash % 13) - 6;
  }

  function ensureDifferentAxes(changed) {
    if (state.x !== state.y) return;
    var fallback = dimensions.find(function (dimension) {
      return dimension.slug !== (changed === "x" ? state.x : state.y);
    });
    if (!fallback) return;
    if (changed === "x") {
      state.y = fallback.slug;
    } else {
      state.x = fallback.slug;
    }
  }

  function populateControls() {
    xSelect.innerHTML = dimensions
      .map(function (dimension) {
        return '<option value="' + escapeHtml(dimension.slug) + '">' + escapeHtml(dimension.label) + "</option>";
      })
      .join("");
    ySelect.innerHTML = xSelect.innerHTML;
    categorySelect.innerHTML =
      '<option value="all">All categories</option>' +
      (data.categories || [])
        .map(function (category) {
          return '<option value="' + escapeHtml(category.slug) + '">' + escapeHtml(category.name) + "</option>";
        })
        .join("");
    xSelect.value = state.x;
    ySelect.value = state.y;
    categorySelect.value = state.category;
  }

  function syncControls() {
    searchInput.value = state.query;
    xSelect.value = state.x;
    ySelect.value = state.y;
    categorySelect.value = state.category;
  }

  function coordinate(value, min, max, invert) {
    var number = Number(value) || 0;
    var ratio = Math.max(0, Math.min(100, number)) / 100;
    if (invert) ratio = 1 - ratio;
    return min + ratio * (max - min);
  }

  function renderPlot(points) {
    var width = 900;
    var height = 620;
    var margin = { left: 72, right: 34, top: 46, bottom: 76 };
    var x0 = margin.left;
    var x1 = width - margin.right;
    var y0 = margin.top;
    var y1 = height - margin.bottom;
    var xDimension = dimensionBySlug[state.x] || dimensions[0];
    var yDimension = dimensionBySlug[state.y] || dimensions[1] || dimensions[0];
    var ticks = [0, 20, 40, 60, 80, 100];
    var pieces = [];

    plotNode.setAttribute("viewBox", "0 0 " + width + " " + height);
    plotNode.setAttribute("xmlns", "http://www.w3.org/2000/svg");
    pieces.push('<defs><filter id="fallacy-map-shadow" x="-40%" y="-40%" width="180%" height="180%"><feDropShadow dx="0" dy="7" stdDeviation="7" flood-color="#111111" flood-opacity="0.16"/></filter></defs>');
    pieces.push('<rect class="fallacy-map-frame" x="' + x0 + '" y="' + y0 + '" width="' + (x1 - x0) + '" height="' + (y1 - y0) + '" rx="14"></rect>');

    ticks.forEach(function (tick) {
      var x = coordinate(tick, x0, x1, false);
      var y = coordinate(tick, y0, y1, true);
      pieces.push('<line class="fallacy-map-grid-line" x1="' + x + '" x2="' + x + '" y1="' + y0 + '" y2="' + y1 + '"></line>');
      pieces.push('<line class="fallacy-map-grid-line" x1="' + x0 + '" x2="' + x1 + '" y1="' + y + '" y2="' + y + '"></line>');
      pieces.push('<text class="fallacy-map-tick" x="' + x + '" y="' + (y1 + 32) + '" text-anchor="middle">' + tick + "</text>");
      pieces.push('<text class="fallacy-map-tick" x="' + (x0 - 20) + '" y="' + (y + 5) + '" text-anchor="end">' + tick + "</text>");
    });

    pieces.push('<text class="fallacy-map-quadrant" x="' + (x0 + 18) + '" y="' + (y0 + 32) + '">' + escapeHtml(xDimension.lowLabel + " + " + yDimension.highLabel) + "</text>");
    pieces.push('<text class="fallacy-map-quadrant" x="' + (x1 - 18) + '" y="' + (y0 + 32) + '" text-anchor="end">' + escapeHtml(xDimension.highLabel + " + " + yDimension.highLabel) + "</text>");
    pieces.push('<text class="fallacy-map-quadrant" x="' + (x0 + 18) + '" y="' + (y1 - 20) + '">' + escapeHtml(xDimension.lowLabel + " + " + yDimension.lowLabel) + "</text>");
    pieces.push('<text class="fallacy-map-quadrant" x="' + (x1 - 18) + '" y="' + (y1 - 20) + '" text-anchor="end">' + escapeHtml(xDimension.highLabel + " + " + yDimension.lowLabel) + "</text>");

    pieces.push('<text class="fallacy-map-axis-title" x="' + ((x0 + x1) / 2) + '" y="' + (height - 18) + '" text-anchor="middle">' + escapeHtml(xDimension.label) + "</text>");
    pieces.push('<text class="fallacy-map-axis-title" transform="translate(24 ' + ((y0 + y1) / 2) + ') rotate(-90)" text-anchor="middle">' + escapeHtml(yDimension.label) + "</text>");
    pieces.push('<text class="fallacy-map-axis" x="' + x0 + '" y="' + (height - 2) + '" text-anchor="middle">' + escapeHtml(xDimension.lowLabel) + "</text>");
    pieces.push('<text class="fallacy-map-axis" x="' + x1 + '" y="' + (height - 2) + '" text-anchor="end">' + escapeHtml(xDimension.highLabel) + "</text>");
    pieces.push('<text class="fallacy-map-axis" x="' + (x0 - 20) + '" y="' + (y0 - 16) + '" text-anchor="middle">' + escapeHtml(yDimension.highLabel) + "</text>");
    pieces.push('<text class="fallacy-map-axis" x="' + (x0 - 20) + '" y="' + (y1 + 48) + '" text-anchor="middle">' + escapeHtml(yDimension.lowLabel) + "</text>");

    if (!points.length) {
      pieces.push('<text class="fallacy-map-empty" x="' + ((x0 + x1) / 2) + '" y="' + ((y0 + y1) / 2) + '" text-anchor="middle">No fallacies match those filters.</text>');
    }

    points.forEach(function (point) {
      var x = coordinate(point.values[state.x], x0, x1, false) + hashOffset(point.slug, "x") * 0.8;
      var y = coordinate(point.values[state.y], y0, y1, true) + hashOffset(point.slug, "y") * 0.8;
      var selectedClass = point.slug === state.selectedSlug ? " is-selected" : "";
      var radius = point.slug === state.selectedSlug ? 10 : 7;
      pieces.push('<circle tabindex="0" role="button" aria-label="' + escapeHtml(point.name) + '" data-map-slug="' + escapeHtml(point.slug) + '" class="fallacy-map-dot' + selectedClass + '" cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="' + radius + '" fill="' + escapeHtml(point.color) + '" filter="url(#fallacy-map-shadow)"><title>' + escapeHtml(point.name) + "</title></circle>");
    });

    plotNode.innerHTML = pieces.join("");

    plotNode.querySelectorAll("[data-map-slug]").forEach(function (node) {
      node.addEventListener("click", function () {
        state.selectedSlug = node.getAttribute("data-map-slug");
        update();
      });
      node.addEventListener("focus", function () {
        state.selectedSlug = node.getAttribute("data-map-slug");
        update();
      });
      node.addEventListener("mouseenter", function () {
        state.selectedSlug = node.getAttribute("data-map-slug");
        update();
      });
    });
  }

  function renderLegend() {
    legendNode.innerHTML =
      '<button type="button" data-map-legend-category="all" aria-pressed="' +
      (state.category === "all" ? "true" : "false") +
      '"><span class="fallacy-map-swatch" style="background: linear-gradient(90deg, #14a8d7, #e63b34)"></span>All</button>' +
      (data.categories || [])
        .map(function (category) {
          return (
            '<button type="button" data-map-legend-category="' +
            escapeHtml(category.slug) +
            '" aria-pressed="' +
            (state.category === category.slug ? "true" : "false") +
            '"><span class="fallacy-map-swatch" style="background:' +
            escapeHtml(category.color) +
            '"></span>' +
            escapeHtml(category.name) +
            "</button>"
          );
        })
        .join("");

    legendNode.querySelectorAll("[data-map-legend-category]").forEach(function (button) {
      button.addEventListener("click", function () {
        state.category = button.getAttribute("data-map-legend-category") || "all";
        update();
      });
    });
  }

  function renderDetail(points) {
    var selected = pointBySlug(state.selectedSlug);
    if (!selected || points.indexOf(selected) === -1) {
      selected = points[0] || (data.points || [])[0];
      state.selectedSlug = selected ? selected.slug : "";
    }

    if (!selected) {
      detailNode.innerHTML = '<span class="eyebrow-pill">Selected fallacy</span><p>No fallacies are available.</p>';
      return;
    }

    var xDimension = dimensionBySlug[state.x] || dimensions[0];
    var yDimension = dimensionBySlug[state.y] || dimensions[1] || dimensions[0];
    var categoryChips = (selected.categories || [])
      .map(function (category) {
        return '<span class="fallacy-map-chip">' + escapeHtml(category) + "</span>";
      })
      .join("");
    var classroomChips = (selected.classroomTags || [])
      .map(function (tag) {
        return '<span class="fallacy-map-chip">' + escapeHtml(tag) + "</span>";
      })
      .join("");

    detailNode.innerHTML =
      '<span class="eyebrow-pill">Selected fallacy</span>' +
      '<h3><a href="' +
      escapeHtml(selected.href) +
      '">' +
      escapeHtml(selected.name) +
      "</a></h3>" +
      '<div class="fallacy-map-category-line"><span class="fallacy-map-category-dot" style="background:' +
      escapeHtml(selected.color) +
      '"></span>' +
      escapeHtml(selected.category) +
      "</div>" +
      '<div class="fallacy-map-stat-grid">' +
      '<div class="fallacy-map-stat"><span>' +
      escapeHtml(xDimension.label) +
      "</span><strong>" +
      escapeHtml(selected.values[state.x]) +
      "</strong></div>" +
      '<div class="fallacy-map-stat"><span>' +
      escapeHtml(yDimension.label) +
      "</span><strong>" +
      escapeHtml(selected.values[state.y]) +
      "</strong></div>" +
      "</div>" +
      "<p>" +
      escapeHtml(selected.definition) +
      "</p>" +
      '<p class="fallacy-map-note"><strong>' +
      escapeHtml(xDimension.label) +
      " note:</strong> " +
      escapeHtml(selected.dimensionNotes[state.x].summary) +
      "</p>" +
      '<p class="fallacy-map-note"><strong>' +
      escapeHtml(yDimension.label) +
      " note:</strong> " +
      escapeHtml(selected.dimensionNotes[state.y].summary) +
      "</p>" +
      '<div class="fallacy-map-chip-row">' +
      categoryChips +
      classroomChips +
      "</div>" +
      '<a class="fallacy-map-open" href="' +
      escapeHtml(selected.href) +
      '">Open fallacy page</a>';
  }

  function update() {
    ensureDifferentAxes();
    var points = visiblePoints();
    if (countNode) {
      countNode.textContent = points.length + (points.length === 1 ? " plotted fallacy" : " plotted fallacies");
    }
    syncControls();
    renderPlot(points);
    renderLegend();
    renderDetail(points);
  }

  populateControls();
  update();

  searchInput.addEventListener("input", function () {
    state.query = searchInput.value;
    update();
  });

  xSelect.addEventListener("change", function () {
    state.x = xSelect.value;
    ensureDifferentAxes("x");
    update();
  });

  ySelect.addEventListener("change", function () {
    state.y = ySelect.value;
    ensureDifferentAxes("y");
    update();
  });

  categorySelect.addEventListener("change", function () {
    state.category = categorySelect.value;
    update();
  });

  resetButton.addEventListener("click", function () {
    state.x = data.defaultAxes && data.defaultAxes.x ? data.defaultAxes.x : "common";
    state.y = data.defaultAxes && data.defaultAxes.y ? data.defaultAxes.y : "difficulty";
    state.category = "all";
    state.query = "";
    state.selectedSlug = data.points && data.points.length ? data.points[0].slug : "";
    update();
  });
})();
