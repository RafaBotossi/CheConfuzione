/* global XLSX */

const state = {
  items: [],
  selectedDay: "all",
  search: "",
  source: "",
  editingId: null,
  addresses: {},
};

const elements = {
  fileInput: document.querySelector("#fileInput"),
  importButton: document.querySelector("#importButton"),
  searchInput: document.querySelector("#searchInput"),
  daySelect: document.querySelector("#daySelect"),
  dayStrip: document.querySelector("#dayStrip"),
  table: document.querySelector("#itineraryTable"),
  mobileList: document.querySelector("#mobileList"),
  resultsCount: document.querySelector("#resultsCount"),
  clearFilters: document.querySelector("#clearFilters"),
  emptyClear: document.querySelector("#emptyClear"),
  emptyState: document.querySelector("#emptyState"),
  tableShell: document.querySelector("#tableShell"),
  sourceLabel: document.querySelector("#sourceLabel"),
  toast: document.querySelector("#toast"),
  addressDialog: document.querySelector("#addressDialog"),
  addressForm: document.querySelector("#addressForm"),
  addressInput: document.querySelector("#addressInput"),
  addressItemName: document.querySelector("#addressItemName"),
  removeAddress: document.querySelector("#removeAddress"),
};

const columnAliases = {
  date: ["data", "date", "dia"],
  time: ["horario", "hora", "time"],
  city: ["cidade", "city", "localidade", "destino"],
  type: ["tipo", "categoria", "type"],
  item: ["item", "atividade", "o que fazer", "programacao", "passeio"],
  notes: ["observacao", "observacoes", "obs", "detalhes"],
  address: ["endereco", "local", "address"],
  value: ["valor", "preco", "custo", "total"],
};

function normalize(value = "") {
  return String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function escapeHtml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function linkify(value = "") {
  return escapeHtml(value).replace(
    /(https?:\/\/[^\s<]+)/g,
    '<a href="$1" target="_blank" rel="noopener noreferrer">Abrir link ↗</a>',
  );
}

function excelDateToIso(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return [
      value.getUTCFullYear(),
      String(value.getUTCMonth() + 1).padStart(2, "0"),
      String(value.getUTCDate()).padStart(2, "0"),
    ].join("-");
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const text = String(value ?? "").trim();
  let match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (match) {
    const year = match[3].length === 2 ? `20${match[3]}` : match[3];
    return `${year}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  }

  match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : "";
}

function formatTime(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${String(value.getUTCHours()).padStart(2, "0")}:${String(value.getUTCMinutes()).padStart(2, "0")}`;
  }

  if (typeof value === "number" && value >= 0 && value < 1) {
    const totalMinutes = Math.round(value * 24 * 60) % (24 * 60);
    return `${String(Math.floor(totalMinutes / 60)).padStart(2, "0")}:${String(totalMinutes % 60).padStart(2, "0")}`;
  }

  const text = String(value ?? "").trim();
  const match = text.match(/(\d{1,2})\s*[:h]\s*(\d{2})/i);
  if (match) return `${match[1].padStart(2, "0")}:${match[2]}`;
  return text || "—";
}

function findHeaderRow(rows) {
  let best = { index: -1, score: 0 };
  rows.slice(0, 15).forEach((row, index) => {
    const normalized = row.map(normalize);
    const score = Object.entries(columnAliases)
      .filter(([key]) => key !== "value")
      .reduce((sum, [, aliases]) => sum + (normalized.some((cell) => aliases.includes(cell)) ? 1 : 0), 0);
    if (score > best.score) best = { index, score };
  });
  return best.score >= 3 ? best.index : -1;
}

function mapColumns(header) {
  const mapping = {};
  header.forEach((cell, index) => {
    const name = normalize(cell);
    Object.entries(columnAliases).forEach(([key, aliases]) => {
      if (mapping[key] === undefined && aliases.includes(name)) mapping[key] = index;
    });
  });
  return mapping;
}

function parseWorkbook(workbook) {
  let chosen = null;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: "",
      blankrows: false,
    });
    const headerIndex = findHeaderRow(rows);
    if (headerIndex >= 0) {
      chosen = { sheetName, rows, headerIndex };
      break;
    }
  }

  if (!chosen) {
    throw new Error("Não encontrei uma aba com as colunas Data, Cidade e Item.");
  }

  const mapping = mapColumns(chosen.rows[chosen.headerIndex]);
  const get = (row, key) => (mapping[key] === undefined ? "" : row[mapping[key]]);

  const items = chosen.rows
    .slice(chosen.headerIndex + 1)
    .map((row, index) => {
      const date = excelDateToIso(get(row, "date"));
      const time = formatTime(get(row, "time"));
      const city = String(get(row, "city") ?? "").trim();
      const type = String(get(row, "type") ?? "").trim();
      const item = String(get(row, "item") ?? "").trim();
      const notes = String(get(row, "notes") ?? "").trim();
      const address = String(get(row, "address") ?? "").trim();
      const id = [date, time, normalize(city), normalize(item), index].join("|");
      return { id, date, time, city, type, item, notes, address };
    })
    .filter((entry) => entry.date && (entry.item || entry.city || entry.type))
    .filter((entry) => !/^(total|subtotal)/i.test(entry.item || entry.notes))
    .sort((a, b) => {
      const dateCompare = a.date.localeCompare(b.date);
      if (dateCompare !== 0) return dateCompare;
      return timeSortValue(a.time) - timeSortValue(b.time);
    });

  if (!items.length) throw new Error("A planilha foi lida, mas não encontrei itens do roteiro.");
  return { items, sheetName: chosen.sheetName };
}

function timeSortValue(time) {
  const match = String(time).match(/^(\d{1,2}):(\d{2})$/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : 9999;
}

function readWorkbook(buffer) {
  // Keep Excel dates as serial numbers. This avoids historical timezone offsets
  // affecting time-only cells (1899) when the trip is opened in another country.
  return XLSX.read(buffer, { type: "array", cellDates: false });
}

function initializeApp() {
  buildDayFilters();
  elements.sourceLabel.textContent = "Importe a planilha para começar";
  elements.resultsCount.textContent = "Nenhuma planilha carregada";
  elements.tableShell.hidden = true;
  elements.mobileList.hidden = true;
  elements.emptyState.hidden = true;
}

function bootRender() {
  buildDayFilters();
  render();
  elements.sourceLabel.textContent = state.source;
}

function getDays() {
  return [...new Set(state.items.map((item) => item.date).filter(Boolean))].sort();
}

function formatDate(date, options = {}) {
  if (!date) return "Sem data";
  const value = new Date(`${date}T12:00:00`);
  return new Intl.DateTimeFormat("pt-BR", options).format(value);
}

function buildDayFilters() {
  const days = getDays();
  elements.dayStrip.style.setProperty("--day-count", days.length + 1);
  elements.daySelect.innerHTML = '<option value="all">Todos os dias</option>';
  elements.dayStrip.innerHTML = `
    <button class="day-chip active" data-day="all" type="button">
      <span>Todos</span><strong>dias</strong>
    </button>`;

  days.forEach((day) => {
    const selectOption = document.createElement("option");
    selectOption.value = day;
    selectOption.textContent = formatDate(day, { weekday: "short", day: "2-digit", month: "2-digit" });
    elements.daySelect.append(selectOption);

    const button = document.createElement("button");
    button.className = "day-chip";
    button.dataset.day = day;
    button.type = "button";
    button.innerHTML = `
      <span>${escapeHtml(formatDate(day, { weekday: "short" }).replace(".", ""))}</span>
      <strong>${escapeHtml(formatDate(day, { day: "2-digit", month: "2-digit" }))}</strong>`;
    elements.dayStrip.append(button);
  });
}

function getVisibleItems() {
  const term = normalize(state.search);
  return state.items.filter((item) => {
    if (state.selectedDay !== "all" && item.date !== state.selectedDay) return false;
    if (!term) return true;
    return normalize(Object.values(item).join(" ")).includes(term);
  });
}

function getAddress(item) {
  return state.addresses[item.id] ?? item.address ?? "";
}

function getMapQuery(item) {
  const address = getAddress(item);
  return address || [item.item, item.city].filter(Boolean).join(", ");
}

function mapMarkup(item) {
  const address = getAddress(item);
  const query = getMapQuery(item);
  const mapUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}`;
  return `
    <div class="map-cell">
      <span class="address-text">${address ? escapeHtml(address) : "Endereço ainda não preenchido"}</span>
      <div class="map-actions">
        <a class="map-link" href="${mapUrl}" target="_blank" rel="noopener noreferrer">
          <svg aria-hidden="true" viewBox="0 0 24 24">
            <path d="M20 10c0 5-8 11-8 11S4 15 4 10a8 8 0 1 1 16 0Z"></path>
            <circle cx="12" cy="10" r="2.5"></circle>
          </svg>
          ${address ? "Abrir Maps" : "Buscar local"}
        </a>
        <button class="edit-address" data-address-id="${escapeHtml(item.id)}" type="button">
          ${address ? "Editar" : "+ Endereço"}
        </button>
      </div>
    </div>`;
}

function typeMarkup(type) {
  if (!type) return "—";
  return `<span class="type-pill" data-type="${escapeHtml(normalize(type))}">${escapeHtml(type)}</span>`;
}

function render() {
  const visible = getVisibleItems();
  const hasFilters = state.selectedDay !== "all" || Boolean(state.search.trim());

  elements.resultsCount.textContent =
    visible.length === 1 ? "1 momento no roteiro" : `${visible.length} momentos no roteiro`;
  elements.clearFilters.hidden = !hasFilters;
  elements.emptyState.hidden = visible.length !== 0;
  elements.tableShell.hidden = visible.length === 0;
  elements.mobileList.hidden = visible.length === 0;

  elements.table.innerHTML = visible
    .map(
      (item, index) => `
        <tr style="animation-delay:${Math.min(index, 12) * 28}ms">
          <td class="when-cell">
            <strong>${escapeHtml(formatDate(item.date, { day: "2-digit", month: "short" }))}</strong>
            <span>${escapeHtml(item.time)}</span>
          </td>
          <td class="city-cell">${escapeHtml(item.city || "—")}</td>
          <td>${typeMarkup(item.type)}</td>
          <td class="item-cell">${escapeHtml(item.item || "—")}</td>
          <td class="notes-cell">${item.notes ? linkify(item.notes) : "—"}</td>
          <td>${mapMarkup(item)}</td>
        </tr>`,
    )
    .join("");

  elements.mobileList.innerHTML = visible
    .map(
      (item, index) => `
        <article class="mobile-card" style="animation-delay:${Math.min(index, 12) * 38}ms">
          <div class="mobile-card-top">
            <div>
              <div class="mobile-when">
                ${escapeHtml(formatDate(item.date, { weekday: "short", day: "2-digit", month: "short" }))}
                · ${escapeHtml(item.time)}
              </div>
              <div class="mobile-city">${escapeHtml(item.city || "Destino a definir")}</div>
            </div>
            ${typeMarkup(item.type)}
          </div>
          <h3>${escapeHtml(item.item || "Atividade")}</h3>
          ${item.notes ? `<p class="mobile-notes">${linkify(item.notes)}</p>` : ""}
          ${mapMarkup(item)}
        </article>`,
    )
    .join("");
}

function chooseDay(day) {
  state.selectedDay = day;
  elements.daySelect.value = day;
  elements.dayStrip.querySelectorAll(".day-chip").forEach((chip) => {
    chip.classList.toggle("active", chip.dataset.day === day);
  });
  render();
}

function clearFilters() {
  state.search = "";
  elements.searchInput.value = "";
  chooseDay("all");
}

async function importFile(file) {
  if (!file) return;
  try {
    const parsed = parseWorkbook(readWorkbook(await file.arrayBuffer()));
    state.items = parsed.items;
    state.source = `${file.name} · aba ${parsed.sheetName}`;
    state.selectedDay = "all";
    state.search = "";
    state.addresses = {};
    bootRender();
    showToast(`${state.items.length} itens importados com sucesso.`);
  } catch (error) {
    console.error(error);
    showToast(error.message || "Não foi possível ler essa planilha.");
  } finally {
    elements.fileInput.value = "";
  }
}

function openAddressDialog(id) {
  const item = state.items.find((entry) => entry.id === id);
  if (!item) return;
  state.editingId = id;
  elements.addressItemName.textContent = [item.item, item.city].filter(Boolean).join(" · ");
  elements.addressInput.value = getAddress(item);
  elements.removeAddress.hidden = !getAddress(item);
  elements.addressDialog.showModal();
  requestAnimationFrame(() => elements.addressInput.focus());
}

function saveAddress() {
  if (!state.editingId) return;
  const value = elements.addressInput.value.trim();
  if (value) state.addresses[state.editingId] = value;
  else delete state.addresses[state.editingId];
  render();
  showToast(value ? "Endereço salvo durante esta sessão." : "Endereço removido.");
}

function showToast(message) {
  elements.toast.textContent = message;
  elements.toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => elements.toast.classList.remove("show"), 3600);
}

elements.importButton.addEventListener("click", () => elements.fileInput.click());
elements.fileInput.addEventListener("change", (event) => importFile(event.target.files[0]));
elements.searchInput.addEventListener("input", (event) => {
  state.search = event.target.value;
  render();
});
elements.daySelect.addEventListener("change", (event) => chooseDay(event.target.value));
elements.dayStrip.addEventListener("click", (event) => {
  const chip = event.target.closest("[data-day]");
  if (chip) chooseDay(chip.dataset.day);
});
elements.clearFilters.addEventListener("click", clearFilters);
elements.emptyClear.addEventListener("click", clearFilters);
document.addEventListener("click", (event) => {
  const button = event.target.closest("[data-address-id]");
  if (button) openAddressDialog(button.dataset.addressId);
});
elements.addressForm.addEventListener("submit", (event) => {
  if (event.submitter?.value === "save") saveAddress();
});
elements.removeAddress.addEventListener("click", () => {
  elements.addressInput.value = "";
  saveAddress();
  elements.addressDialog.close();
});
if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
  window.addEventListener("load", () => navigator.serviceWorker.register("./sw.js").catch(() => {}));
}

initializeApp();
