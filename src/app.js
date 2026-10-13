/* Everything this app needs is embedded by scripts/build.py. No network calls. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = {figures: [], filtered: [], current: null, tab: 'chart', dataPage: 0, selectedSeries: 0, imageZoom: 1, renderSequence: 0};
  const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const pretty = value => typeof value === 'object' && value !== null ? JSON.stringify(value) : String(value ?? '');
  const clean = text => String(text || '').replace(/\s+/g, ' ').trim();
  const figureLabel = f => clean(f.title).match(/^Figure\s+[\w.]+/i)?.[0] || f.sheet;
  const shortTitle = f => clean(f.title).replace(/^Figure\s+[\w.]+\s*[-–—]\s*/i, '');
  const kindLabel = f => ({interactive: 'Interactive', reference: 'Reference figure', unavailable: 'Source issue'}[f.kind] || 'Interactive');
  const clone = value => structuredClone(value);
  function notify(message) { $('toast').textContent = message; $('toast').hidden = false; clearTimeout(notify.timer); notify.timer = setTimeout(() => $('toast').hidden = true, 3500); }
  function closeMenu() { $('sidebar').classList.remove('open'); $('scrim').hidden = true; $('menu-button').setAttribute('aria-expanded', 'false'); }

  function filterFigures() {
    const terms = $('search').value.toLowerCase().split(/\s+/).filter(Boolean);
    state.filtered = state.figures.filter(f => (!$('section').value || f.section === $('section').value) && (!$('kind').value || f.kind === $('kind').value) && terms.every(t => `${f.title} ${f.subtitle || ''} ${f.sheet} ${f.section} ${f.id}`.toLowerCase().includes(t)));
    $('result-count').textContent = `${state.filtered.length} ${state.filtered.length === 1 ? 'figure' : 'figures'}`;
    let lastSection = '';
    $('figure-list').innerHTML = state.filtered.map(f => {
      const heading = f.section !== lastSection ? `<div class="nav-group">${escapeHTML(f.section)}</div>` : '';
      lastSection = f.section;
      return `${heading}<a class="figure-link" href="#${encodeURIComponent(f.id)}" data-figure="${escapeHTML(f.id)}" ${state.current?.id === f.id ? 'aria-current="page"' : ''}><span class="figure-link-top"><span>${escapeHTML(figureLabel(f))}</span><span>${f.kind === 'reference' ? 'REF' : f.kind === 'unavailable' ? '!' : '↗'}</span></span><span class="figure-link-title">${escapeHTML(shortTitle(f))}${f.subtitle ? ' · ' + escapeHTML(f.subtitle) : ''}</span></a>`;
    }).join('') || '<p class="empty-results">No figures match. Try a different term or clear the filters.</p>';
    updatePagination();
  }

  function updatePagination() {
    const index = state.filtered.findIndex(f => f.id === state.current?.id);
    $('previous-figure').disabled = index <= 0;
    $('next-figure').disabled = index < 0 || index === state.filtered.length - 1;
    $('filtered-position').textContent = index < 0 ? 'Current figure is outside this filter' : `${index + 1} of ${state.filtered.length} figures`;
  }

  function displayLayout(figure) {
    const layout = clone(figure.layout || {});
    layout.autosize = true;
    layout.height = Math.max(window.innerWidth < 800 ? 440 : 510, layout.height || 0);
    delete layout.width;
    layout.paper_bgcolor = '#fff';
    layout.plot_bgcolor = '#fff';
    layout.font = {...layout.font, family: 'Arial, sans-serif', color: '#365047', size: window.innerWidth < 600 ? 10 : 11};
    layout.margin = {...{l: 75, r: 40, t: 40, b: 95}, ...layout.margin};
    // Pie legends represent categories; keep Plotly's category-aware toggles for those.
    layout.showlegend = (figure.data || []).some(t => t.type === 'pie');
    if (window.innerWidth < 600) {
      layout.margin.l = Math.min(layout.margin.l, 60);
      layout.margin.r = Math.min(layout.margin.r, 55);
      for (const [key, axis] of Object.entries(layout)) {
        if (/^[xy]axis\d*$/.test(key) && axis && axis.title) axis.title.font = {size: 10};
      }
    }
    layout.uirevision = figure.id;
    return layout;
  }

  function legendColor(trace) {
    const hasLine = (trace.mode || 'lines').includes('lines') && trace.line?.width !== 0;
    const isBar = trace.type === 'bar' || trace.type === 'histogram';
    const candidate = isBar ? (trace.marker?.pattern?.fgcolor || trace.marker?.color)
      : hasLine ? (trace.line?.color || trace.fillcolor || trace.marker?.color)
      : (trace.marker?.color || trace.fillcolor || trace.line?.color);
    const color = Array.isArray(candidate) ? candidate.find(c => typeof c === 'string') : candidate;
    return typeof color === 'string' && color.toUpperCase() !== '#FFFFFF' ? color : trace.marker?.line?.color || '#427867';
  }
  function renderLegend(figure) {
    const visible = (figure.data || []).map((trace, index) => ({trace, index})).filter(({trace}) => trace.showlegend !== false && trace.type !== 'pie');
    $('legend').hidden = !visible.length;
    $('legend').innerHTML = visible.length ? '<div class="legend-actions"><span>SERIES · click to toggle, shift-click to isolate</span><button id="show-all-series">Show all</button></div>' + visible.map(({trace, index}) => `<button class="legend-button" data-trace="${index}" aria-pressed="${trace.visible !== 'legendonly' && trace.visible !== false}" title="${escapeHTML(trace.name || 'Series ' + (index + 1))}"><span class="legend-swatch" style="background:${escapeHTML(legendColor(trace))};${trace.line?.dash && trace.line.dash !== 'solid' ? 'border-top:2px dashed white' : ''}"></span><span class="legend-name">${escapeHTML(trace.name || 'Series ' + (index + 1))}</span></button>`).join('') : '';
    $('show-all-series')?.addEventListener('click', async () => {
      const groups = new Set(visible.map(x => x.trace.legendgroup).filter(Boolean));
      const ids = figure.data.map((trace, index) => trace.showlegend !== false || groups.has(trace.legendgroup) ? index : -1).filter(index => index >= 0);
      await Plotly.restyle($('plot'), {visible: true}, ids);
      $('legend').querySelectorAll('[data-trace]').forEach(b => b.setAttribute('aria-pressed', 'true'));
    });
  }

  async function selectFigure(id, options = {}) {
    const figure = state.figures.find(f => f.id === id) || state.figures[0];
    if (!figure) return;
    const sequence = ++state.renderSequence;
    state.current = figure; state.dataPage = 0; state.selectedSeries = 0; state.imageZoom = 1;
    if (options.updateHash !== false && decodeURIComponent(location.hash.slice(1)) !== figure.id) history.replaceState(null, '', '#' + encodeURIComponent(figure.id));
    document.title = `${figureLabel(figure)} · WEM 2026 outlook`;
    $('figure-title').textContent = clean(figure.title);
    $('section-label').textContent = figure.section;
    $('type-label').textContent = kindLabel(figure);
    $('subtitle').textContent = figure.subtitle || '';
    $('subtitle').hidden = !figure.subtitle;
    $('sheet-label').textContent = `Worksheet: ${figure.sheet}`;
    $('series-label').textContent = figure.kind === 'interactive' ? `${figure.data.length} plotted series` : figure.kind === 'reference' ? 'Original workbook image' : 'Broken workbook references';
    $('position-label').textContent = `Figure ${state.figures.indexOf(figure) + 1} / ${state.figures.length}`;
    $('figure-permalink').href = '#' + encodeURIComponent(figure.id);
    const issues = (figure.notes || []).filter(n => /broken|placeholder|#REF|mismatch|differs|inconsisten|do not match|exceed|unavailable|stored as text|beyond|discrepan|not supplied/i.test(n));
    $('source-warning').hidden = !issues.length && figure.kind !== 'unavailable';
    $('source-warning').textContent = figure.kind === 'unavailable' ? 'This chart object has broken source references. Its cached placeholders are not forecast data. See Source & notes.' : issues.length ? `${issues[0]}${issues.length > 1 ? ' More details in Source & notes.' : ''}` : '';
    $('export-csv').disabled = figure.kind !== 'interactive';
    $('export-image').disabled = figure.kind === 'unavailable';
    $('reset').disabled = figure.kind === 'unavailable';
    $('tab-data').disabled = figure.kind !== 'interactive';
    if (state.tab === 'data' && figure.kind !== 'interactive') state.tab = 'chart';
    $('plot').hidden = figure.kind !== 'interactive';
    $('reference').hidden = figure.kind !== 'reference';
    $('unavailable').hidden = figure.kind !== 'unavailable';
    $('legend').innerHTML = '';
    $('legend').hidden = figure.kind !== 'interactive';
    $('interaction-hint').hidden = figure.kind === 'unavailable';
    $('interaction-hint').textContent = figure.kind === 'reference' ? 'Original workbook figure · use the zoom controls and scroll to inspect' : 'Hover to inspect · drag to zoom · double-click to reset · click a series to toggle';
    if ($('plot').data) Plotly.purge($('plot'));
    filterFigures();
    renderSource(figure);
    populateSeries(figure);
    setTab(state.tab, false);
    if (figure.kind === 'interactive') {
      const config = {responsive: true, displaylogo: false, scrollZoom: false, modeBarButtonsToRemove: ['select2d', 'lasso2d'], toImageButtonOptions: {format: 'png', filename: figure.id, scale: 2}};
      await Plotly.newPlot($('plot'), clone(figure.data), displayLayout(figure), config);
      if (sequence !== state.renderSequence) return;
      renderLegend(figure);
      if (state.tab === 'chart') Plotly.Plots.resize($('plot'));
    } else if (figure.kind === 'reference') renderReference(figure);
    else {
      $('unavailable').innerHTML = '<div class="eyebrow accent">SOURCE CHART UNAVAILABLE</div><h3>The workbook contains a broken chart object.</h3><p>The source series refer to missing cells and contain cached placeholder values. There is no valid numerical chart to recreate from this object.</p>' + (figure.related_chart_ids || []).map(related => `<button data-related="${escapeHTML(related)}">Open the related valid chart →</button>`).join('');
    }
    closeMenu();
    window.WEM_APP.lastRendered = figure.id;
    return figure.id;
  }

  function setTab(tab, resize = true) {
    state.tab = tab;
    for (const name of ['chart', 'data', 'source']) {
      $('panel-' + name).hidden = name !== tab;
      $('tab-' + name).setAttribute('aria-selected', String(name === tab));
      $('tab-' + name).tabIndex = name === tab ? 0 : -1;
    }
    if (tab === 'data') renderTable();
    if (tab === 'chart' && resize && state.current?.kind === 'interactive' && $('plot').data) requestAnimationFrame(() => Plotly.Plots.resize($('plot')));
  }

  function pointCount(trace) { return Math.max(trace.y?.length || 0, trace.values?.length || 0, trace.x?.length || 0); }
  function xValue(trace, index) {
    if (trace.labels) return trace.labels[index];
    if (trace.x) {
      // Multicategory Plotly arrays have an outer level for each category tier.
      if (Array.isArray(trace.x[0])) return trace.x.map(level => level[index]).join(' / ');
      return trace.x[index];
    }
    if (trace.x0 !== undefined && trace.dx !== undefined) {
      if (typeof trace.x0 === 'number') return trace.x0 + index * trace.dx;
      const initial = Date.parse(trace.x0);
      if (Number.isFinite(initial)) return new Date(initial + index * trace.dx).toISOString();
    }
    return index + 1;
  }
  const yValue = (trace, index) => trace.y?.[index] ?? trace.values?.[index] ?? null;
  function populateSeries(figure) {
    $('data-series').innerHTML = (figure.data || []).map((trace, i) => `<option value="${i}">${escapeHTML(trace.name || 'Series ' + (i + 1))}</option>`).join('');
  }
  function renderTable() {
    const trace = state.current?.data?.[state.selectedSeries];
    if (!trace) { $('data-table').innerHTML = '<p>No numerical data supplied for this figure.</p>'; return; }
    const count = pointCount(trace), start = state.dataPage * 100, end = Math.min(start + 100, count);
    const raw = Array.isArray(trace.customdata);
    $('table-info').textContent = `${count.toLocaleString()} observations · ${count ? start + 1 : 0}–${end}`;
    const rows = [];
    for (let i = start; i < end; i++) rows.push(`<tr><td>${i + 1}</td><td>${escapeHTML(pretty(xValue(trace, i)))}</td><td>${yValue(trace, i) === null ? '<em>Missing</em>' : escapeHTML(pretty(yValue(trace, i)))}</td>${raw ? `<td>${escapeHTML(pretty(trace.customdata[i]))}</td>` : ''}</tr>`);
    $('data-table').innerHTML = `<table><caption class="sr-only">${escapeHTML(trace.name)} data, in displayed chart units. Blank observations remain missing.</caption><thead><tr><th>Point</th><th>Category / x</th><th>Plotted value / y</th>${raw ? '<th>Source value / context</th>' : ''}</tr></thead><tbody>${rows.join('')}</tbody></table>`;
    $('data-prev').disabled = state.dataPage === 0;
    $('data-next').disabled = end >= count;
  }

  function renderSource(figure) {
    const notes = figure.notes || [];
    const description = figure.kind === 'reference' ? 'This worksheet contains an illustration, map or overview image. The original image is embedded here for zoomable inspection; numerical interactivity is provided wherever underlying values are available.' : figure.kind === 'unavailable' ? 'The Excel chart object is retained in the inventory. Its broken references prevent a valid reconstruction.' : 'Values come from the saved Excel chart series or the worksheet data behind its published figure. Display-unit conversions and chart-specific repairs are recorded below.';
    const sourceRows = (figure.sources || []).map(source => {
      if (typeof source === 'string') return `<tr><td>Source range</td><td><code>${escapeHTML(source)}</code></td></tr>`;
      const name = source.series_name || source.role || source.name || source.sheet || 'Source';
      const details = Object.entries(source).filter(([key]) => !['series_name', 'validation'].includes(key)).map(([key, value]) => `${key}: ${pretty(value)}`).join('\n');
      const validation = source.validation ? `\nSource check: ${source.validation.source_resolved ? 'resolved' : 'unresolved'}; ${source.validation.mismatch_count || 0} mismatches` : '';
      return `<tr><td>${escapeHTML(name)}</td><td><code>${escapeHTML(details + validation)}</code></td></tr>`;
    }).join('');
    $('panel-source').innerHTML = `<h3>Provenance & interpretation</h3><p>${escapeHTML(description)}</p><p><strong>Workbook sheet:</strong> ${escapeHTML(figure.sheet)}${figure.source_xml ? `<br><strong>Chart object:</strong> <code>${escapeHTML(figure.source_xml)}</code>` : ''}</p>${notes.length ? `<h3>Figure notes</h3><ul>${notes.map(note => `<li>${escapeHTML(note)}</li>`).join('')}</ul>` : '<p>No chart-specific corrections were required.</p>'}${sourceRows ? `<h3>Source ranges</h3><div class="table-wrap" style="margin:0;max-height:none"><table class="source-table"><thead><tr><th>Series / role</th><th>Worksheet references & transformations</th></tr></thead><tbody>${sourceRows}</tbody></table></div>` : ''}<h3>Original publications</h3><div class="source-downloads">${sourceLinks()}</div>${figure.image_data && figure.kind !== 'reference' ? `<details class="original-reference"><summary>Compare with the original workbook image</summary><img src="${figure.image_data}" alt="Original workbook figure: ${escapeHTML(figure.title)}" loading="lazy"></details>` : ''}`;
  }

  function renderReference(figure) {
    $('reference').innerHTML = `<div class="reference-tools"><button id="image-minus" aria-label="Zoom out">−</button><span id="image-scale">Fit to width</span><button id="image-plus" aria-label="Zoom in">+</button><button id="image-fit">Fit</button></div><div class="reference-frame"><img id="reference-image" src="${figure.image_data}" alt="${escapeHTML(figure.title + (figure.subtitle ? ' — ' + figure.subtitle : ''))}"></div><div class="reference-caption">Original figure supplied in the AEMO workbook. Open Source & notes for context and related numerical figures.</div>`;
    const applyZoom = () => { $('reference-image').style.width = `${Math.round(state.imageZoom * 100)}%`; $('image-scale').textContent = state.imageZoom === 1 ? 'Fit to width' : `${Math.round(state.imageZoom * 100)}%`; };
    $('image-plus').onclick = () => { state.imageZoom = Math.min(4, state.imageZoom + .25); applyZoom(); };
    $('image-minus').onclick = () => { state.imageZoom = Math.max(.5, state.imageZoom - .25); applyZoom(); };
    $('image-fit').onclick = () => { state.imageZoom = 1; applyZoom(); };
    applyZoom();
  }

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob), anchor = document.createElement('a');
    anchor.href = url; anchor.download = name; document.body.append(anchor); anchor.click(); anchor.remove(); setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
  const csvCell = value => {
    let text = pretty(value);
    if (typeof value === 'string' && /^[=+@-]/.test(text)) text = "'" + text;
    return '"' + text.replace(/"/g, '""') + '"';
  };
  function exportCSV() {
    const figure = state.current;
    if (figure.kind !== 'interactive') return;
    const chunks = ['\ufefffigure_id,worksheet,series,point,x,y,source_value_or_context\r\n'];
    for (const trace of figure.data) {
      const count = pointCount(trace);
      let chunk = '';
      for (let i = 0; i < count; i++) {
        chunk += [figure.id, figure.sheet, trace.name, i + 1, xValue(trace, i), yValue(trace, i), trace.customdata?.[i] ?? null].map(csvCell).join(',') + '\r\n';
        if (i % 10000 === 9999) { chunks.push(chunk); chunk = ''; }
      }
      if (chunk) chunks.push(chunk);
    }
    downloadBlob(new Blob(chunks, {type: 'text/csv;charset=utf-8'}), `${figure.id}-${figure.sheet.replace(/[^\w.-]/g, '_')}.csv`);
    notify('CSV exported with all series and observations.');
  }

  function sourceLinks() {
    return (state.sources || []).map(s => `<a href="${escapeHTML(s.url)}" target="_blank" rel="noopener">${escapeHTML(s.filename.includes('data-register') ? 'Data register ↗' : s.filename.includes('demand-traces') ? 'Demand traces ↗' : s.filename.includes('appendices') ? 'Appendices ↗' : 'Report ↗')}</a>`).join('');
  }
  function moveFigure(offset) { const index = state.filtered.findIndex(f => f.id === state.current?.id), next = state.filtered[index + offset]; if (next) selectFigure(next.id); }

  async function init() {
    const binary = atob($('embedded-data').textContent.trim());
    const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));
    if (!window.DecompressionStream) throw new Error('This offline file needs a current version of Chrome, Edge, Firefox or Safari with gzip decompression support.');
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
    const payload = JSON.parse(await new Response(stream).text());
    state.figures = payload.figures; state.sources = payload.sources;
    window.WEM_APP = {state, selectFigure, filterFigures, setTab, displayLayout, xValue, yValue, pointCount, lastRendered: null};
    const interactive = state.figures.filter(f => f.kind === 'interactive').length, references = state.figures.filter(f => f.kind === 'reference').length, unavailable = state.figures.filter(f => f.kind === 'unavailable').length;
    $('figure-total').textContent = `${state.figures.length} figures & chart objects`;
    $('summary-strip').innerHTML = `<span><strong>${interactive}</strong> interactive charts</span><span class="summary-dot"></span><span><strong>${references}</strong> reference figures</span><span class="summary-dot"></span><span><strong>${unavailable}</strong> source issues documented</span><span class="summary-dot"></span><span>2026 edition · original AEMO data</span>`;
    for (const section of new Set(state.figures.map(f => f.section))) $('section').add(new Option(section, section));
    $('source-links').innerHTML = sourceLinks();
    filterFigures();
    await selectFigure(decodeURIComponent(location.hash.slice(1)) || state.figures[0].id);
    return true;
  }

  $('search').addEventListener('input', filterFigures);
  $('section').addEventListener('change', filterFigures);
  $('kind').addEventListener('change', filterFigures);
  $('clear-filters').onclick = () => { $('search').value = ''; $('section').value = ''; $('kind').value = ''; filterFigures(); };
  $('figure-list').addEventListener('click', event => { const target = event.target.closest('[data-figure]'); if (target) { event.preventDefault(); selectFigure(target.dataset.figure); } });
  $('unavailable').addEventListener('click', event => { if (event.target.dataset.related) selectFigure(event.target.dataset.related); });
  document.querySelectorAll('[data-tab]').forEach(button => { button.onclick = () => setTab(button.dataset.tab); button.addEventListener('keydown', event => { if (['ArrowRight','ArrowLeft'].includes(event.key)) { const tabs = [...document.querySelectorAll('[data-tab]')].filter(tab => !tab.disabled), index = tabs.indexOf(button), next = tabs[(index + (event.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length]; next.click(); next.focus(); event.preventDefault(); } }); });
  $('legend').addEventListener('click', async event => {
    const button = event.target.closest('[data-trace]'); if (!button) return;
    const index = Number(button.dataset.trace), plot = $('plot'), trace = plot.data[index];
    const grouped = trace.legendgroup ? plot.data.map((t, i) => t.legendgroup === trace.legendgroup ? i : -1).filter(i => i >= 0) : [index];
    if (event.shiftKey) {
      const groups = new Set(plot.data.filter(t => t.showlegend !== false).map(t => t.legendgroup).filter(Boolean));
      const ids = plot.data.map((t, i) => t.showlegend !== false || groups.has(t.legendgroup) ? i : -1).filter(i => i >= 0);
      await Plotly.restyle(plot, {visible: ids.map(i => grouped.includes(i) ? true : 'legendonly')}, ids);
    } else await Plotly.restyle(plot, {visible: trace.visible === 'legendonly' || trace.visible === false ? true : 'legendonly'}, grouped);
    $('legend').querySelectorAll('[data-trace]').forEach(b => b.setAttribute('aria-pressed', String(plot.data[Number(b.dataset.trace)].visible !== 'legendonly' && plot.data[Number(b.dataset.trace)].visible !== false)));
  });
  $('previous-figure').onclick = () => moveFigure(-1);
  $('next-figure').onclick = () => moveFigure(1);
  $('reset').onclick = async () => { if (state.current?.kind === 'reference') $('image-fit').click(); else if (state.current?.kind === 'interactive') await selectFigure(state.current.id); };
  $('export-csv').onclick = exportCSV;
  $('export-image').onclick = async () => {
    if (state.current.kind === 'reference') { const response = await fetch(state.current.image_data); downloadBlob(await response.blob(), state.current.id + '.png'); }
    else if (state.current.kind === 'interactive') {
      try {
        // Include native legend and title in the export so it is useful outside the explorer.
        const layout = clone($('plot').layout);
        await Plotly.relayout($('plot'), {showlegend: true, title: {text: figureLabel(state.current) + ' · ' + state.current.sheet, font: {size: 15}}, 'margin.t': 55});
        await Plotly.downloadImage($('plot'), {format: 'png', filename: state.current.id, width: 1500, height: Math.max(850, layout.height || 0), scale: 1.5});
        await Plotly.relayout($('plot'), {showlegend: layout.showlegend, title: layout.title || {}, 'margin.t': layout.margin.t});
      } catch (error) { notify('Image export failed. Try the chart toolbar camera.'); console.error(error); }
    }
  };
  $('data-series').onchange = () => { state.selectedSeries = Number($('data-series').value); state.dataPage = 0; renderTable(); };
  $('data-prev').onclick = () => { state.dataPage--; renderTable(); };
  $('data-next').onclick = () => { state.dataPage++; renderTable(); };
  $('menu-button').onclick = () => { const open = !$('sidebar').classList.contains('open'); $('sidebar').classList.toggle('open', open); $('scrim').hidden = !open; $('menu-button').setAttribute('aria-expanded', String(open)); };
  $('scrim').onclick = closeMenu;
  window.addEventListener('hashchange', () => { if (state.figures.length) selectFigure(decodeURIComponent(location.hash.slice(1))); });
  document.addEventListener('keydown', event => { if (event.key === '/' && !/INPUT|TEXTAREA|SELECT/.test(event.target.tagName)) { event.preventDefault(); if (window.innerWidth < 801) $('menu-button').click(); $('search').focus(); } if (event.key === 'Escape') closeMenu(); });
  window.WEM_READY = init().catch(error => { $('figure-title').textContent = 'The explorer could not load.'; $('subtitle').hidden = false; $('subtitle').textContent = error.message; console.error(error); throw error; });
})();
