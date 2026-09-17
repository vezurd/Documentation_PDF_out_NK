(function () {
  "use strict";

  const TAB_IDS = ["kits", "an"];

  const ENDPOINTS = {
    kits: "/api/kits",
    an: "/api/an",
    kitsXlsx: "/api/kits/xlsx",
  };

  const PATH_HEADER_RE = /путь|папка/i;
  const KITS_CARD_WIDTH_KEY = "rdweb-kits-card-width";
  const KITS_CARD_MIN = 140;
  const KITS_CARD_MAX = 480;
  const KITS_CARD_DEFAULT = 163;

  const state = {
    meta: null,
    tables: {},
    tableSpecs: {},
    pending: { tab: "kits", title: "", mark: "" },
    kitsCardWidth: KITS_CARD_DEFAULT,
  };

  const els = {
    meta: document.getElementById("meta-line"),
    warning: document.getElementById("warning"),
    refresh: document.getElementById("refresh-btn"),
    tabBar: document.getElementById("tab-bar"),
    filterKits: document.getElementById("filter-kits"),
    codeA: document.getElementById("flt-code-a"),
    tdo: document.getElementById("flt-tdo"),
    noAb: document.getElementById("flt-no-ab"),
    onlyAb: document.getElementById("flt-only-ab"),
    mtoProblems: document.getElementById("flt-mto-problems"),
    anCloses: document.getElementById("flt-an-closes"),
    legendBtn: document.getElementById("btn-kits-legend"),
    xlsxBtn: document.getElementById("btn-kits-xlsx"),
    legendDialog: document.getElementById("kits-legend-dialog"),
    legendTitle: document.getElementById("kits-legend-title"),
    legendIntro: document.getElementById("kits-legend-intro"),
    legendBody: document.getElementById("kits-legend-body"),
    kitPaneButtons: document.querySelectorAll("#kit-card [data-kit-pane]"),
    kitPaneTips: document.getElementById("kit-pane-tips"),
    kitPaneCard: document.getElementById("kit-pane-card"),
    kitTips: document.getElementById("kit-tips"),
    cardEmpty: document.getElementById("card-empty"),
    cardBody: document.getElementById("card-body"),
    cardIdent: document.getElementById("card-ident"),
    cardBadges: document.getElementById("card-badges"),
    cardMto: document.getElementById("card-mto"),
    cardPackages: document.getElementById("card-packages"),
    kitsSplit: document.getElementById("kits-split"),
    kitsGutter: document.getElementById("kits-gutter"),
    kitsLoadOverlay: document.getElementById("kits-load-overlay"),
  };

  function cellText(value) {
    if (value == null) return "";
    if (typeof value === "object") return value.text == null ? "" : String(value.text);
    return String(value);
  }

  function cellSortKey(value) {
    if (value && typeof value === "object" && value.sort_key != null) {
      return value.sort_key;
    }
    return cellText(value);
  }

  function compareValues(a, b) {
    const ka = cellSortKey(a);
    const kb = cellSortKey(b);
    if (typeof ka === "number" && typeof kb === "number") return ka - kb;
    const sa = Array.isArray(ka) ? ka.join("\0") : String(ka);
    const sb = Array.isArray(kb) ? kb.join("\0") : String(kb);
    return sa.localeCompare(sb, "ru", { numeric: true, sensitivity: "base" });
  }

  function paintCell(el, value) {
    el.style.backgroundColor = "";
    el.style.color = "";
    el.style.fontWeight = "";
    el.style.textDecoration = "";
    el.removeAttribute("title");
    if (!value || typeof value !== "object") return;
    if (value.fill) el.style.backgroundColor = value.fill;
    if (value.foreground) el.style.color = value.foreground;
    if (value.bold) el.style.fontWeight = "700";
    if (value.underline) el.style.textDecoration = "underline";
    if (value.tooltip) el.title = value.tooltip;
  }

  function cellFormatter(cell) {
    const value = cell.getValue();
    const el = cell.getElement();
    if (el) paintCell(el, value);
    return cellText(value);
  }

  function copyText(text) {
    const value = text == null ? "" : String(text);
    if (!value) return;
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(value).catch(function () {
        fallbackCopy(value);
      });
      return;
    }
    fallbackCopy(value);
  }

  function fallbackCopy(text) {
    const area = document.createElement("textarea");
    area.value = text;
    area.setAttribute("readonly", "");
    area.style.position = "fixed";
    area.style.left = "-9999px";
    document.body.appendChild(area);
    area.select();
    try {
      document.execCommand("copy");
    } catch (_err) {
      /* ignore */
    }
    area.remove();
  }

  function looksLikePath(field, text) {
    if (!text) return false;
    if (PATH_HEADER_RE.test(field || "")) return true;
    return text.startsWith("\\\\") || /^[A-Za-z]:[\\/]/.test(text) || text.startsWith("//");
  }

  function kitFromRow(data) {
    const title = (data && (data._title || data.title)) || "";
    const mark = (data && (data._mark || data.mark)) || "";
    return { title: String(title), mark: String(mark) };
  }

  function fold(text) {
    return String(text || "").toLocaleLowerCase("ru");
  }

  function sameKit(aTitle, aMark, bTitle, bMark) {
    return fold(aTitle) === fold(bTitle) && fold(aMark) === fold(bMark);
  }

  function contextMenuItems(row, cell) {
    const data = row.getData();
    const kit = kitFromRow(data);
    const hasKit = Boolean(kit.title && kit.mark);
    const items = [];
    function nav(label, tab) {
      return {
        label: label,
        disabled: !hasKit,
        action: function () {
          navigateTo(tab, kit.title, kit.mark);
        },
      };
    }
    items.push(nav("Показать в Комплекты", "kits"));
    items.push(nav("Показать в АН", "an"));
    if (cell) {
      const field = cell.getField() || "";
      const text = cellText(cell.getValue());
      if (looksLikePath(field, text)) {
        items.push({ separator: true });
        items.push({
          label: "Копировать путь",
          action: function () {
            copyText(text);
          },
        });
      }
    }
    return items;
  }

  function columnsFromHeaders(headers, options) {
    const opts = options || {};
    const frozenCount = opts.frozenCount || 0;
    const revStart = opts.revStart;
    const widths = opts.widths || {};
    return headers.map(function (header, index) {
      const isRev = revStart != null && index >= revStart;
      const column = {
        title: header,
        field: header,
        formatter: cellFormatter,
        sorter: function (a, b) {
          return compareValues(a, b);
        },
        headerSort: true,
        resizable: true,
        movable: isRev ? false : opts.movable !== false,
        frozen: index < frozenCount,
        hozAlign: isRev ? "center" : "left",
        headerHozAlign: isRev ? "center" : "left",
        minWidth: isRev ? 36 : 72,
      };
      const width = widths[header];
      if (typeof width === "number" && width > 0) {
        column.width = width;
      }
      return column;
    });
  }

  function flattenGeneric(rows, headers) {
    return (rows || []).map(function (row) {
      const kit = Array.isArray(row.kit_key) ? row.kit_key : [];
      const titleCell = row.cells && row.cells["Титул"];
      const markCell = row.cells && row.cells["Марка"];
      const out = {
        _title: row.title || kit[0] || cellText(titleCell),
        _mark: row.mark || kit[1] || cellText(markCell),
        _haystack: String(row.haystack || ""),
        _tooltips_text: String(row.tooltips_text || ""),
        paints_kits_issuance: Boolean(row.paints_kits_issuance),
        code_a: Boolean(row.code_a),
        kit_ok: Boolean(row.kit_ok),
        kit_tdo_passed: Boolean(row.kit_tdo_passed),
        as_build: Boolean(row.as_build),
        has_mto_problem: Boolean(row.has_mto_problem),
        an_closes_auto_mto: Boolean(row.an_closes_auto_mto),
      };
      const cells = row.cells || {};
      headers.forEach(function (header) {
        out[header] = cells[header] || { text: "" };
      });
      return out;
    });
  }

  function haystackOf(data) {
    if (data._haystack) return String(data._haystack);
    const parts = [];
    Object.keys(data).forEach(function (key) {
      if (key.charAt(0) === "_") return;
      const text = cellText(data[key]);
      if (text) parts.push(text);
    });
    return parts.join(" ");
  }

  function textFilter(needle) {
    const folded = fold(needle.trim());
    if (!folded) return null;
    return function (data) {
      return fold(haystackOf(data)).indexOf(folded) !== -1;
    };
  }

  function applyTabFilter(tabId) {
    const table = state.tables[tabId];
    if (!table) return;
    const input = document.getElementById("filter-" + tabId);
    const needle = input ? input.value : "";
    if (tabId === "kits") {
      table.setFilter(kitsFilterFn(needle));
      return;
    }
    const fn = textFilter(needle);
    if (fn) table.setFilter(fn);
    else table.clearFilter();
  }

  function kitsFilterFn(needle) {
    const textFn = textFilter(needle);
    const codeA = els.codeA.checked;
    const tdo = els.tdo.checked;
    const noAb = els.noAb.checked;
    const onlyAb = els.onlyAb.checked;
    const mtoProblems = els.mtoProblems.checked;
    const anCloses = els.anCloses.checked;
    return function (data) {
      if (textFn && !textFn(data)) return false;
      if (codeA && !data.code_a) return false;
      if (tdo && !data.kit_tdo_passed) return false;
      if (noAb && data.as_build) return false;
      if (onlyAb && !data.as_build) return false;
      if (mtoProblems && !data.has_mto_problem) return false;
      if (anCloses && !data.an_closes_auto_mto) return false;
      return true;
    };
  }

  function updateCount(tabId) {
    const table = state.tables[tabId];
    const node = document.getElementById("count-" + tabId);
    if (!table || !node) return;
    const visible = table.getDataCount("active");
    const total = table.getDataCount();
    if (tabId === "kits") {
      const meta = state.meta || {};
      let text = meta.kits_progress_text || "";
      const sep = meta.kits_progress_visible_sep || " · видно: ";
      if (visible !== total) text += sep + visible;
      node.textContent = text;
      return;
    }
    node.textContent = "Строки: " + visible + " / " + total;
  }

  function persistOptions(tableId, extra) {
    const options = {
      persistenceMode: "local",
      persistenceID: "rdweb-" + tableId,
      persistence: extra && extra.persistence ? extra.persistence : { sort: true, columns: true },
    };
    return options;
  }

  function destroyTable(tabId) {
    const table = state.tables[tabId];
    if (!table) return;
    try {
      table.destroy();
    } catch (_err) {
      /* ignore */
    }
    state.tables[tabId] = null;
  }

  function makeTable(tabId, elementId, headers, rows, extra) {
    extra = extra || {};
    destroyTable(tabId);
    const host = document.getElementById(elementId);
    if (!host) return null;
    const columns = columnsFromHeaders(headers, extra);
    const options = Object.assign(
      {
        data: rows,
        columns: columns,
        layout: extra.layout || "fitDataFill",
        nestedFieldSeparator: false,
        movableColumns: extra.movableColumns !== false,
        resizableColumns: true,
        headerSort: true,
        reactiveData: false,
        placeholder: "Нет строк",
        selectableRows: 1,
        rowHeight: 24,
        columnDefaults: { vertAlign: "middle" },
        cellContextMenu: function (_e, cell) {
          return contextMenuItems(cell.getRow(), cell);
        },
      },
      persistOptions(tabId, extra)
    );
    if (extra.renderHorizontal) {
      options.renderHorizontal = extra.renderHorizontal;
    }
    let table;
    try {
      table = new Tabulator(host, options);
    } catch (err) {
      console.error("Tabulator init failed:", tabId, err);
      els.warning.hidden = false;
      els.warning.textContent =
        "Ошибка таблицы: " + (err && err.message ? err.message : String(err));
      return null;
    }
    table.on("dataFiltered", function () {
      updateCount(tabId);
    });
    table.on("dataLoaded", function () {
      updateCount(tabId);
    });
    if (tabId === "kits") {
      table.on("rowClick", function (_e, row) {
        const data = row.getData();
        const kit = kitFromRow(data);
        showKitTips(data);
        if (kit.title && kit.mark) loadCard(kit.title, kit.mark);
      });
    }
    state.tables[tabId] = table;
    updateCount(tabId);
    applyTabFilter(tabId);
    return table;
  }

  function ensureTable(tabId) {
    if (state.tables[tabId]) return state.tables[tabId];
    const spec = state.tableSpecs[tabId];
    if (!spec) return null;
    return makeTable(tabId, spec.elementId, spec.headers, spec.rows, spec.extra);
  }

  function setTab(tabId, pushUrl) {
    if (TAB_IDS.indexOf(tabId) < 0) tabId = "kits";
    document.querySelectorAll(".tab-bar [role='tab']").forEach(function (btn) {
      btn.setAttribute("aria-selected", btn.getAttribute("data-tab") === tabId ? "true" : "false");
    });
    document.querySelectorAll(".tab-panel").forEach(function (panel) {
      panel.hidden = panel.getAttribute("data-panel") !== tabId;
    });
    ensureTable(tabId);
    const table = state.tables[tabId];
    if (table) {
      try {
        table.redraw(true);
      } catch (_err) {
        /* Tabulator may measure before the first layout pass. */
      }
    }
    if (pushUrl !== false) writeUrl(tabId, state.pending.title, state.pending.mark);
  }

  function readUrl() {
    const params = new URLSearchParams(window.location.search);
    return {
      tab: params.get("tab") || "kits",
      title: params.get("title") || "",
      mark: params.get("mark") || "",
    };
  }

  function writeUrl(tab, title, mark) {
    const params = new URLSearchParams();
    params.set("tab", tab);
    if (title) params.set("title", title);
    if (mark) params.set("mark", mark);
    const next = "?" + params.toString();
    if (next !== window.location.search) {
      history.replaceState(null, "", next);
    }
    state.pending = { tab: tab, title: title, mark: mark };
  }

  function selectKitOnTable(tabId, title, mark) {
    const table = state.tables[tabId];
    if (!table || !title || !mark) return false;
    const rows = table.getRows("active");
    let match = null;
    function consider(row) {
      const kit = kitFromRow(row.getData());
      if (!sameKit(kit.title, kit.mark, title, mark)) return;
      if (!match) match = row;
    }
    for (let i = 0; i < rows.length; i += 1) {
      consider(rows[i]);
    }
    if (!match) {
      const all = table.getRows();
      for (let j = 0; j < all.length; j += 1) {
        consider(all[j]);
      }
    }
    const chosen = match;
    if (!chosen) return false;
    chosen.select();
    table.scrollToRow(chosen, "center", false).catch(function () {
      /* ignore */
    });
    if (tabId === "kits") {
      const data = chosen.getData();
      const kit = kitFromRow(data);
      showKitTips(data);
      loadCard(kit.title, kit.mark);
    }
    return true;
  }

  function navigateTo(tab, title, mark) {
    state.pending = { tab: tab, title: title, mark: mark };
    setTab(tab, true);
    selectKitOnTable(tab, title, mark);
  }

  function showKitTips(data) {
    if (!els.kitTips) return;
    const text = (data && data._tooltips_text) || "";
    els.kitTips.textContent = text || "Нет подсказок для этой строки.";
  }

  function setKitPane(pane) {
    const tips = pane !== "card";
    if (els.kitPaneTips) els.kitPaneTips.hidden = !tips;
    if (els.kitPaneCard) els.kitPaneCard.hidden = tips;
    if (els.kitPaneButtons) {
      els.kitPaneButtons.forEach(function (btn) {
        btn.setAttribute(
          "aria-selected",
          btn.getAttribute("data-kit-pane") === (tips ? "tips" : "card")
            ? "true"
            : "false"
        );
      });
    }
  }

  function paintBadge(cell) {
    const span = document.createElement("span");
    span.className = "badge";
    span.textContent = cellText(cell) || "—";
    paintCell(span, cell);
    return span;
  }

  function renderCard(card) {
    els.cardEmpty.hidden = true;
    els.cardBody.hidden = false;
    els.cardIdent.textContent = (card.title || "") + "–" + (card.mark || "");
    els.cardBadges.replaceChildren();
    els.cardBadges.appendChild(paintBadge(card.review));
    els.cardBadges.appendChild(paintBadge(card.approval));
    els.cardMto.textContent = card.mto_rollup || "";
    const headers =
      (state.meta && state.meta.card_package_headers) || [
        "NN",
        "Пакет",
        "Рев.",
        "Дата",
        "Выдача",
        "Рассмотрение",
        "MTO",
        "AB",
        "Ликвидность",
      ];
    const table = els.cardPackages;
    table.replaceChildren();
    const thead = document.createElement("thead");
    const hr = document.createElement("tr");
    headers.forEach(function (header) {
      const th = document.createElement("th");
      th.textContent = header;
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    table.appendChild(thead);
    const tbody = document.createElement("tbody");
    (card.packages || []).forEach(function (pkg) {
      const tr = document.createElement("tr");
      const cells = pkg.cells || {};
      headers.forEach(function (header) {
        const td = document.createElement("td");
        const cell = cells[header] || { text: "" };
        td.textContent = cellText(cell);
        paintCell(td, cell);
        if (pkg.package_path && header === "Пакет") {
          td.title = (cell.tooltip ? cell.tooltip + "\n" : "") + pkg.package_path;
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }

  function loadCard(title, mark) {
    els.cardEmpty.hidden = false;
    els.cardEmpty.textContent = "Загрузка карточки…";
    els.cardBody.hidden = true;
    const url =
      "/api/kits/" + encodeURIComponent(title) + "/" + encodeURIComponent(mark) + "/card";
    fetchJson(url)
      .then(function (card) {
        renderCard(card);
        writeUrl("kits", title, mark);
      })
      .catch(function () {
        els.cardEmpty.textContent = "Комплект не найден";
        els.cardBody.hidden = true;
      });
  }

  function formatScan(lastScan) {
    if (!lastScan) return "скан: нет";
    const stamp = lastScan.completed_at || lastScan.started_at || "";
    const status = lastScan.status || "";
    return "скан: " + (stamp || "—") + (status ? " · " + status : "");
  }

  function showMeta(meta) {
    state.meta = meta;
    const warning = (meta && meta.warning) || "";
    if (warning) {
      els.warning.hidden = false;
      els.warning.textContent = warning;
    } else {
      els.warning.hidden = true;
      els.warning.textContent = "";
    }
    const generated = (meta && meta.generated_at) || "";
    els.meta.textContent =
      formatScan(meta && meta.last_scan) + (generated ? " · сборка: " + generated : "");
    if (els.legendBtn && meta && meta.kits_paint_legend_button) {
      els.legendBtn.textContent = meta.kits_paint_legend_button;
    }
    if (els.xlsxBtn && meta && meta.kits_xlsx_button) {
      els.xlsxBtn.textContent = meta.kits_xlsx_button;
    }
  }

  function paintLegendSample(sample) {
    const el = document.createElement("span");
    el.className = "legend-sample";
    el.textContent = sample && sample.text != null ? String(sample.text) : "";
    if (sample && sample.fill) el.style.backgroundColor = sample.fill;
    if (sample && sample.foreground) el.style.color = sample.foreground;
    if (sample && sample.bold) el.classList.add("is-bold");
    if (sample && sample.underline) el.classList.add("is-underline");
    return el;
  }

  function renderKitsLegend(meta) {
    if (!els.legendTitle || !els.legendIntro || !els.legendBody) return;
    els.legendTitle.textContent =
      (meta && meta.kits_paint_legend_title) || "Легенда таблицы Комплекты";
    els.legendIntro.textContent = (meta && meta.kits_paint_legend_intro) || "";
    els.legendBody.replaceChildren();
    const sections = (meta && meta.kits_paint_legend) || [];
    sections.forEach(function (section) {
      const block = document.createElement("section");
      block.className = "legend-section";
      const heading = document.createElement("h3");
      heading.textContent = section.title || "";
      block.appendChild(heading);
      if (section.intro) {
        const intro = document.createElement("p");
        intro.textContent = section.intro;
        block.appendChild(intro);
      }
      const table = document.createElement("table");
      table.className = "legend-table";
      const head = document.createElement("thead");
      head.innerHTML = "<tr><th>Столбец</th><th>Пример</th><th>Что значит</th></tr>";
      table.appendChild(head);
      const body = document.createElement("tbody");
      (section.samples || []).forEach(function (sample) {
        const row = document.createElement("tr");
        const col = document.createElement("td");
        col.textContent = sample.column || "";
        const example = document.createElement("td");
        example.appendChild(paintLegendSample(sample));
        const meaning = document.createElement("td");
        meaning.textContent = sample.meaning || "";
        row.appendChild(col);
        row.appendChild(example);
        row.appendChild(meaning);
        body.appendChild(row);
      });
      table.appendChild(body);
      block.appendChild(table);
      els.legendBody.appendChild(block);
    });
  }

  function openKitsLegend() {
    renderKitsLegend(state.meta || {});
    const dialog = els.legendDialog;
    if (!dialog) return;
    if (typeof dialog.showModal === "function") {
      dialog.showModal();
      return;
    }
    dialog.setAttribute("open", "");
  }

  function visibleKitKeys(table) {
    const keys = [];
    let rows = [];
    if (table && typeof table.getRows === "function") {
      try {
        rows = table.getRows("active") || [];
      } catch (_err) {
        rows = [];
      }
    }
    if (rows.length && rows[0] && typeof rows[0].getData === "function") {
      rows.forEach(function (row) {
        const kit = kitFromRow(row.getData());
        if (kit.title && kit.mark) {
          keys.push({ title: kit.title, mark: kit.mark });
        }
      });
      return keys;
    }
    const data =
      (table && table.getData && (table.getData("active") || table.getData())) ||
      [];
    data.forEach(function (item) {
      const kit = kitFromRow(item);
      if (kit.title && kit.mark) {
        keys.push({ title: kit.title, mark: kit.mark });
      }
    });
    return keys;
  }

  function kitsExportColumns(table) {
    if (!table || typeof table.getColumns !== "function") return [];
    let cols = [];
    try {
      cols = table.getColumns() || [];
    } catch (_err) {
      return [];
    }
    const out = [];
    cols.forEach(function (col) {
      if (typeof col.isVisible === "function" && !col.isVisible()) return;
      const def = col.getDefinition ? col.getDefinition() : {};
      const header = def.title || def.field || "";
      if (!header) return;
      let width = 80;
      if (typeof col.getWidth === "function") {
        const live = col.getWidth();
        if (typeof live === "number" && live > 0) width = Math.round(live);
      }
      out.push({ header: header, width_px: width });
    });
    return out;
  }

  function filenameFromDisposition(header) {
    if (!header) return "";
    const star = /filename\*\s*=\s*UTF-8''([^;]+)/i.exec(header);
    if (star) {
      try {
        return decodeURIComponent(star[1]);
      } catch (_err) {
        return star[1];
      }
    }
    const plain = /filename\s*=\s*"([^"]+)"/i.exec(header);
    return plain ? plain[1] : "";
  }

  function downloadKitsXlsx() {
    const table = state.tables.kits || ensureTable("kits");
    if (!table || !els.xlsxBtn) return;
    els.xlsxBtn.disabled = true;
    fetch(ENDPOINTS.kitsXlsx, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        keys: visibleKitKeys(table),
        columns: kitsExportColumns(table),
      }),
    })
      .then(function (response) {
        if (!response.ok) {
          const error = new Error("xlsx → " + response.status);
          error.status = response.status;
          throw error;
        }
        const filename =
          filenameFromDisposition(response.headers.get("Content-Disposition")) ||
          "Комплекты.xlsx";
        return response.blob().then(function (blob) {
          return { blob: blob, filename: filename };
        });
      })
      .then(function (payload) {
        const url = URL.createObjectURL(payload.blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = payload.filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
      })
      .catch(function (err) {
        els.warning.hidden = false;
        els.warning.textContent =
          "Не удалось сохранить Excel: " +
          (err && err.message ? err.message : err);
      })
      .finally(function () {
        els.xlsxBtn.disabled = false;
      });
  }

  function fetchJson(url) {
    return fetch(url).then(function (response) {
      if (!response.ok) {
        const error = new Error(url + " → " + response.status);
        error.status = response.status;
        throw error;
      }
      return response.json();
    });
  }

  function clampKitsCardWidth(px) {
    const n = Number(px);
    if (!Number.isFinite(n)) return KITS_CARD_DEFAULT;
    return Math.min(KITS_CARD_MAX, Math.max(KITS_CARD_MIN, Math.round(n)));
  }

  function readStoredKitsCardWidth() {
    try {
      const raw = localStorage.getItem(KITS_CARD_WIDTH_KEY);
      if (raw == null || raw === "") return null;
      const n = parseInt(raw, 10);
      if (!Number.isFinite(n) || n <= 0) return null;
      return clampKitsCardWidth(n);
    } catch (_err) {
      return null;
    }
  }

  function applyKitsCardWidth(px) {
    const width = clampKitsCardWidth(px);
    state.kitsCardWidth = width;
    if (els.kitsSplit) {
      els.kitsSplit.style.setProperty("--kits-card-width", width + "px");
    }
    return width;
  }

  function persistKitsCardWidth(px) {
    try {
      localStorage.setItem(KITS_CARD_WIDTH_KEY, String(clampKitsCardWidth(px)));
    } catch (_err) {
      /* ignore quota / private mode */
    }
  }

  function applyMetaKitsCardWidthIfNeeded(meta) {
    if (readStoredKitsCardWidth() != null) return;
    const widths = (meta && meta.kits_header_widths) || {};
    const fromMeta = widths["Статус согласования"];
    if (typeof fromMeta === "number" && fromMeta > 0 && Number.isFinite(fromMeta)) {
      applyKitsCardWidth(fromMeta);
    }
  }

  function showKitsLoading(on) {
    if (els.kitsLoadOverlay) els.kitsLoadOverlay.hidden = !on;
  }

  function bindKitsGutter() {
    const gutter = els.kitsGutter;
    const split = els.kitsSplit;
    if (!gutter || !split) return;
    let dragging = false;

    function cardWidthFromPointer(clientX) {
      const rect = split.getBoundingClientRect();
      const gutterW = gutter.getBoundingClientRect().width || 8;
      return split.clientWidth - gutterW - (clientX - rect.left);
    }

    function endDrag(event) {
      if (!dragging) return;
      dragging = false;
      split.classList.remove("is-resizing");
      try {
        gutter.releasePointerCapture(event.pointerId);
      } catch (_err) {
        /* already released */
      }
      persistKitsCardWidth(state.kitsCardWidth);
    }

    gutter.addEventListener("pointerdown", function (event) {
      if (event.button != null && event.button !== 0) return;
      dragging = true;
      split.classList.add("is-resizing");
      gutter.setPointerCapture(event.pointerId);
      applyKitsCardWidth(cardWidthFromPointer(event.clientX));
      event.preventDefault();
    });
    gutter.addEventListener("pointermove", function (event) {
      if (!dragging) return;
      applyKitsCardWidth(cardWidthFromPointer(event.clientX));
    });
    gutter.addEventListener("pointerup", endDrag);
    gutter.addEventListener("pointercancel", endDrag);
  }

  function assignKitsSpec(kits) {
    const meta = state.meta || {};
    state.tableSpecs.kits = {
      elementId: "table-kits",
      headers: meta.kits_headers || [],
      rows: flattenGeneric(kits, meta.kits_headers || []),
      extra: {
        layout: "fitDataFill",
        widths: meta.kits_header_widths || {},
        persistence: { sort: true, columns: false },
      },
    };
  }

  function assignAnSpec(rows) {
    const meta = state.meta || {};
    state.tableSpecs.an = {
      elementId: "table-an",
      headers: meta.an_headers || [],
      rows: flattenGeneric(rows, meta.an_headers || []),
      extra: {},
    };
  }

  function revealPendingSelection(tabId) {
    const pending = state.pending;
    if (!pending.title || !pending.mark) return;
    selectKitOnTable(tabId, pending.title, pending.mark);
    if (tabId === "kits") loadCard(pending.title, pending.mark);
  }

  function restoreMetaLine() {
    if (state.meta) showMeta(state.meta);
  }

  function refreshAll(force) {
    showKitsLoading(true);
    els.meta.textContent = "Загрузка…";
    const metaUrl = force ? "/api/meta?refresh=1" : "/api/meta";
    return fetchJson(metaUrl)
      .then(function (meta) {
        showMeta(meta);
        applyMetaKitsCardWidthIfNeeded(meta);
        return fetchJson(ENDPOINTS.kits);
      })
      .then(function (kits) {
        assignKitsSpec(kits);
        destroyTable("kits");
        const pending = state.pending;
        setTab(pending.tab, true);
        ensureTable("kits");
        showKitsLoading(false);
        if (pending.tab === "kits") {
          revealPendingSelection("kits");
        }
        els.meta.textContent = "Загружаем АН…";
        return fetchJson(ENDPOINTS.an);
      })
      .then(function (anRows) {
        assignAnSpec(anRows);
        destroyTable("an");
        const pending = state.pending;
        if (pending.tab === "an") {
          ensureTable("an");
          revealPendingSelection("an");
        }
        restoreMetaLine();
      })
      .catch(function (err) {
        showKitsLoading(false);
        restoreMetaLine();
        els.warning.hidden = false;
        els.warning.textContent = "Ошибка загрузки: " + (err && err.message ? err.message : err);
      });
  }

  function bindFilters() {
    TAB_IDS.forEach(function (tabId) {
      const input = document.getElementById("filter-" + tabId);
      if (!input) return;
      input.addEventListener("input", function () {
        applyTabFilter(tabId);
      });
    });
    [els.codeA, els.tdo, els.noAb, els.onlyAb, els.mtoProblems, els.anCloses].forEach(
      function (box) {
        box.addEventListener("change", function () {
          if (box === els.noAb && els.noAb.checked) els.onlyAb.checked = false;
          if (box === els.onlyAb && els.onlyAb.checked) els.noAb.checked = false;
          applyTabFilter("kits");
        });
      }
    );
  }

  function init() {
    if (typeof Tabulator === "undefined") {
      els.warning.hidden = false;
      els.warning.textContent =
        "Не загрузился Tabulator (/static/vendor/tabulator.min.js).";
      return;
    }
    state.pending = readUrl();
    const storedWidth = readStoredKitsCardWidth();
    applyKitsCardWidth(storedWidth != null ? storedWidth : KITS_CARD_DEFAULT);
    bindKitsGutter();
    els.tabBar.addEventListener("click", function (event) {
      const btn = event.target.closest("[data-tab]");
      if (!btn) return;
      navigateTo(btn.getAttribute("data-tab"), state.pending.title, state.pending.mark);
    });
    els.refresh.addEventListener("click", function () {
      refreshAll(true);
    });
    bindFilters();
    if (els.kitPaneButtons) {
      els.kitPaneButtons.forEach(function (btn) {
        btn.addEventListener("click", function () {
          setKitPane(btn.getAttribute("data-kit-pane") || "tips");
        });
      });
    }
    setKitPane("tips");
    if (els.legendBtn) {
      els.legendBtn.addEventListener("click", openKitsLegend);
    }
    if (els.xlsxBtn) {
      els.xlsxBtn.addEventListener("click", downloadKitsXlsx);
    }
    const kitsCount = document.getElementById("count-kits");
    if (kitsCount) {
      kitsCount.addEventListener("click", function () {
        copyText(kitsCount.textContent || "");
      });
    }
    refreshAll(false);
  }

  document.addEventListener("DOMContentLoaded", init);
})();
