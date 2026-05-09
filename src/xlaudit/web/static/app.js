/* ── xlaudit dashboard — app.js ────────────────────────────────────── */
(function () {
  "use strict";

  /* ── State ────────────────────────────────────────────────────── */
  let scanData = null;   // { report, graph, errors }
  let sortCol = 7;       // default: complexity (desc)
  let sortAsc = false;

  /* ── DOM refs ────────────────────────────────────────────────── */
  const $ = (s, p) => (p || document).querySelector(s);
  const $$ = (s, p) => [...(p || document).querySelectorAll(s)];

  const dropZone   = $(".drop-zone");
  const fileInput   = $("#file-input");
  const loader      = $(".loader");
  const errorBar    = $(".error-bar");
  const statsWrap   = $("#stats-wrap");
  const tableWrap   = $("#table-wrap");
  const graphPanel  = $("#panel-graph");
  const graphBox    = $("#graph-container");
  const tabBtns     = $$(".tab");
  const panels      = $$(".panel");

  /* ── Tabs ─────────────────────────────────────────────────────── */
  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tabBtns.forEach(b => b.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      $(`#panel-${btn.dataset.panel}`).classList.add("active");
    });
  });

  /* ── Drag & drop + file input ─────────────────────────────────── */
  ["dragenter", "dragover"].forEach(e =>
    dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.add("drag-over"); })
  );
  ["dragleave", "drop"].forEach(e =>
    dropZone.addEventListener(e, () => dropZone.classList.remove("drag-over"))
  );
  dropZone.addEventListener("drop", ev => {
    ev.preventDefault();
    handleFiles(ev.dataTransfer.files);
  });
  fileInput.addEventListener("change", () => handleFiles(fileInput.files));

  /* ── Upload + scan ────────────────────────────────────────────── */
  async function handleFiles(fileList) {
    const files = [...fileList].filter(f => f.name.toLowerCase().endsWith(".xlsx"));
    if (!files.length) { showError("No .xlsx files selected."); return; }

    hideError();
    dropZone.style.display = "none";
    loader.classList.add("active");

    const form = new FormData();
    files.forEach(f => form.append("files", f));

    try {
      const res = await fetch("/api/scan", { method: "POST", body: form });
      if (!res.ok) throw new Error(`Server error ${res.status}`);
      scanData = await res.json();
      renderResults();
    } catch (err) {
      showError(err.message);
      dropZone.style.display = "";
    } finally {
      loader.classList.remove("active");
    }
  }

  /* ── Render results ───────────────────────────────────────────── */
  function renderResults() {
    const wbs = scanData.report.workbooks;
    if (!wbs.length) { showError("No workbooks could be scanned."); dropZone.style.display = ""; return; }

    if (scanData.errors.length) showError(scanData.errors.join(" | "));

    // Stats
    const tf = wbs.reduce((s, w) => s + w.total_formulas, 0);
    const te = wbs.reduce((s, w) => s + w.total_external_links, 0);
    const tv = wbs.reduce((s, w) => s + w.total_volatile, 0);
    statsWrap.innerHTML = `
      <div class="stat"><div class="val">${wbs.length}</div><div class="lbl">Workbooks</div></div>
      <div class="stat"><div class="val">${tf.toLocaleString()}</div><div class="lbl">Total Formulas</div></div>
      <div class="stat"><div class="val">${te}</div><div class="lbl">External Links</div></div>
      <div class="stat"><div class="val">${tv}</div><div class="lbl">Volatile Calls</div></div>
    `;
    statsWrap.style.display = "";

    renderTable(wbs);
    tableWrap.style.display = "";

    // Enable graph tab
    renderGraph();

    // Show rescan link
    dropZone.style.display = "";
    dropZone.querySelector(".label").textContent = "Drop more files or click to rescan";
  }

  /* ── Table ────────────────────────────────────────────────────── */
  const cols = [
    { key: "file_name",           label: "File",         num: false },
    { key: "file_size_kb",        label: "KB",           num: true },
    { key: "sheet_count",         label: "Sheets",       num: true },
    { key: "total_formulas",      label: "Formulas",     num: true },
    { key: "total_external_links",label: "Ext. Links",   num: true },
    { key: "total_volatile",      label: "Volatile",     num: true },
    { key: "named_range_count",   label: "Named Ranges", num: true },
    { key: "complexity_score",    label: "Complexity",    num: true },
  ];

  function renderTable(wbs) {
    const sorted = [...wbs].sort((a, b) => {
      const c = cols[sortCol];
      let va = a[c.key], vb = b[c.key];
      if (c.num) return sortAsc ? va - vb : vb - va;
      return sortAsc ? String(va).localeCompare(vb) : String(vb).localeCompare(va);
    });

    let html = "<table><thead><tr>";
    cols.forEach((c, i) => {
      const arrow = i === sortCol ? (sortAsc ? " &#9650;" : " &#9660;") : "";
      html += `<th data-col="${i}">${c.label}<span class="sort-icon">${arrow}</span></th>`;
    });
    html += "</tr></thead><tbody>";

    sorted.forEach((wb, idx) => {
      const band = wb.complexity_band.toLowerCase();
      html += `<tr class="expandable" data-idx="${idx}">`;
      html += `<td>${esc(wb.file_name)}</td>`;
      html += `<td>${wb.file_size_kb}</td>`;
      html += `<td>${wb.sheet_count}</td>`;
      html += `<td>${wb.total_formulas}</td>`;
      html += `<td>${wb.total_external_links}</td>`;
      html += `<td>${wb.total_volatile}</td>`;
      html += `<td>${wb.named_range_count}</td>`;
      html += `<td>${wb.complexity_score} <span class="band band-${band}">${wb.complexity_band}</span></td>`;
      html += `</tr>`;
      // Detail row
      html += `<tr class="detail-row" id="detail-${idx}"><td colspan="8"><div class="detail-inner">`;
      html += `<h4>Sheets in ${esc(wb.file_name)}</h4>`;
      html += `<table><thead><tr><th>Sheet</th><th>Formulas</th><th>Volatile</th><th>Cross-sheet</th><th>External refs</th></tr></thead><tbody>`;
      (wb.sheets || []).forEach(s => {
        const refs = (s.external_refs || []).join(", ") || "\u2014";
        html += `<tr><td>${esc(s.name)}</td><td>${s.formula_count}</td><td>${s.volatile_count}</td><td>${s.cross_sheet_ref_count}</td><td>${refs}</td></tr>`;
      });
      html += `</tbody></table></div></td></tr>`;
    });
    html += "</tbody></table>";
    tableWrap.innerHTML = `<div class="tbl-wrap">${html}</div>`;

    // Sort headers
    $$("th", tableWrap).forEach(th => {
      th.addEventListener("click", () => {
        const ci = +th.dataset.col;
        if (ci === sortCol) sortAsc = !sortAsc; else { sortCol = ci; sortAsc = false; }
        renderTable(wbs);
      });
    });

    // Expandable rows
    $$("tr.expandable", tableWrap).forEach(tr => {
      tr.addEventListener("click", () => {
        const dr = $(`#detail-${tr.dataset.idx}`);
        dr.classList.toggle("open");
      });
    });
  }

  /* ── Graph (D3 force) ─────────────────────────────────────────── */
  function renderGraph() {
    if (!scanData || !scanData.graph) return;
    graphBox.innerHTML = "";

    const { nodes, links } = scanData.graph;
    if (!nodes.length) {
      graphBox.innerHTML = '<div class="empty-msg">No graph data available</div>';
      return;
    }

    // Deduplicate node IDs — keep first occurrence
    const nodeMap = new Map();
    nodes.forEach(n => { if (!nodeMap.has(n.id)) nodeMap.set(n.id, { ...n }); });
    const uNodes = [...nodeMap.values()];
    const nodeIds = new Set(uNodes.map(n => n.id));

    // Filter links to only reference existing nodes
    const uLinks = links.filter(l => nodeIds.has(l.source) && nodeIds.has(l.target) && l.type !== "contains");

    const W = graphBox.clientWidth || 800;
    const H = graphBox.clientHeight || 550;

    const svg = d3.select(graphBox).append("svg")
      .attr("width", W).attr("height", H)
      .attr("viewBox", [0, 0, W, H]);

    const g = svg.append("g");

    // Zoom
    svg.call(d3.zoom().scaleExtent([0.3, 4]).on("zoom", e => g.attr("transform", e.transform)));

    const sim = d3.forceSimulation(uNodes)
      .force("link", d3.forceLink(uLinks).id(d => d.id).distance(90).strength(0.5))
      .force("charge", d3.forceManyBody().strength(-250))
      .force("center", d3.forceCenter(W / 2, H / 2))
      .force("collision", d3.forceCollide().radius(d => nodeRadius(d) + 6));

    // Links
    const link = g.append("g").selectAll("line").data(uLinks).join("line")
      .attr("stroke", d => d.type === "external" ? "#38bdf8" : "#4b5563")
      .attr("stroke-width", d => Math.min(d.value, 4))
      .attr("stroke-opacity", 0.5)
      .attr("stroke-dasharray", d => d.type === "external" ? "6,3" : "none");

    // Nodes
    const node = g.append("g").selectAll("g").data(uNodes).join("g")
      .call(d3.drag().on("start", dragStart).on("drag", dragging).on("end", dragEnd));

    node.append("circle")
      .attr("r", d => nodeRadius(d))
      .attr("fill", d => nodeColor(d))
      .attr("stroke", d => nodeStroke(d))
      .attr("stroke-width", 2)
      .attr("opacity", 0.9);

    node.append("text")
      .text(d => nodeLabel(d))
      .attr("font-size", d => d.type === "workbook" ? "10px" : "8px")
      .attr("font-family", "var(--font)")
      .attr("fill", "#e4e6f0")
      .attr("text-anchor", "middle")
      .attr("dy", d => nodeRadius(d) + 14);

    // Tooltip
    const tooltip = d3.select(".tooltip");
    node.on("mouseenter", (ev, d) => {
      let html = `<div class="tt-title">${esc(d.id)}</div>`;
      if (d.type === "workbook") html += `<div class="tt-row">Score: ${d.score} (${d.band})</div><div class="tt-row">Size: ${d.kb} KB</div>`;
      else if (d.type === "sheet") html += `<div class="tt-row">Formulas: ${d.formulas}</div><div class="tt-row">Volatile: ${d.volatile}</div>`;
      else html += `<div class="tt-row">External workbook</div>`;
      tooltip.html(html).style("display", "block");
    })
    .on("mousemove", ev => {
      tooltip.style("left", (ev.clientX + 12) + "px").style("top", (ev.clientY - 10) + "px");
    })
    .on("mouseleave", () => tooltip.style("display", "none"));

    sim.on("tick", () => {
      link.attr("x1", d => d.source.x).attr("y1", d => d.source.y)
          .attr("x2", d => d.target.x).attr("y2", d => d.target.y);
      node.attr("transform", d => `translate(${d.x},${d.y})`);
    });

    function dragStart(ev, d) { if (!ev.active) sim.alphaTarget(0.3).restart(); d.fx = d.x; d.fy = d.y; }
    function dragging(ev, d) { d.fx = ev.x; d.fy = ev.y; }
    function dragEnd(ev, d) { if (!ev.active) sim.alphaTarget(0); d.fx = null; d.fy = null; }
  }

  function nodeRadius(d) { return d.type === "workbook" ? 20 : d.type === "external" ? 14 : 10; }
  function nodeColor(d) {
    if (d.type === "external") return "#1e3a5f";
    if (d.type === "workbook") {
      return { LOW: "#064e3b", MED: "#78350f", HIGH: "#7f1d1d" }[d.band] || "#1f2937";
    }
    return "#1e1b4b";
  }
  function nodeStroke(d) {
    if (d.type === "external") return "#38bdf8";
    if (d.type === "workbook") {
      return { LOW: "#34d399", MED: "#fbbf24", HIGH: "#f87171" }[d.band] || "#6b7280";
    }
    return "#818cf8";
  }
  function nodeLabel(d) {
    if (d.type === "sheet") return d.id.split("/").pop();
    return d.id.length > 20 ? d.id.slice(0, 18) + "\u2026" : d.id;
  }

  /* ── Utilities ────────────────────────────────────────────────── */
  function esc(s) {
    const el = document.createElement("span");
    el.textContent = s;
    return el.innerHTML;
  }
  function showError(msg) { errorBar.textContent = msg; errorBar.classList.add("visible"); }
  function hideError() { errorBar.classList.remove("visible"); }
})();
