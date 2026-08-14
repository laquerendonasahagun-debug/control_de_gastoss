const STORAGE_KEY = 'la-querendona-control-gastos-v1';

const budgetItems = [
  { id: 'abarrote', name: 'Abarrote', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'verdura', name: 'Verdura / chiles secos / hierbas de olor', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'pan', name: 'Pan', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'basura', name: 'Basura', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'agua', name: 'Agua', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'limpieza', name: 'Producto limpieza y mantelería', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'gas', name: 'Gas', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'carne', name: 'Carne', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'coca-cola', name: 'Coca-Cola', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'jarritos', name: 'Jarritos', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'cortes', name: 'Cortes, snacks', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'cerveza', name: 'Cerveza', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'leche', name: 'Leche', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'pollo', name: 'Pollo', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'bistek', name: 'Bistek', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'cremeria', name: 'Cremería', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'tortillas', name: 'Tortillas y masa', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'mandaditos', name: 'Mandaditos', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'desechable', name: 'Desechable', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'vinos', name: 'Vinos y licores', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'papeleria', name: 'Papelería', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'comision', name: 'Comisión billipocket', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'caja', name: 'Caja', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'otros', name: 'Otros', weekly: 0, monthly: 0, group: 'operating' },
  { id: 'nomina', name: 'Nómina', weekly: 0, monthly: 0, group: 'fixed' },
  { id: 'renta', name: 'Renta', weekly: 0, monthly: 0, group: 'fixed' },
  { id: 'luz', name: 'Luz', weekly: 0, monthly: 0, group: 'fixed' },
  { id: 'gerencia', name: 'Gerencia', weekly: 0, monthly: 0, group: 'fixed' },
  { id: 'reserva', name: 'Fondo de reserva', weekly: 0, monthly: 0, group: 'fixed' },
];

const spendingPieColors = ['#24584a', '#8cbf8d', '#c17db9', '#83cfc5', '#e4a84c', '#7b8fc5', '#d4776a'];

const periods = (window.EXCEL_PERIODS || []).slice().sort((a, b) => Number(Boolean(b.current)) - Number(Boolean(a.current)) || String(b.weeks.at(-1)?.end || '').localeCompare(String(a.weeks.at(-1)?.end || '')));
const currentPeriodId = periods.find(period => period.current)?.id || periods[0]?.id || '';

const defaultState = () => ({
  budgets: Object.fromEntries(budgetItems.map(item => [item.id, { weekly: item.weekly, monthly: item.monthly }])),
  manualEntries: [],
});

const excelEntries = window.EXCEL_ENTRIES || [];
let state = loadState();
let selectedPeriodId = currentPeriodId;
let selectedWeekIndex = 0;
let activeView = 'dashboard';
let toastTimer;
let activeCaptureMode = 'single';
let bulkDraftEntries = [];
let rangeStartDate = '';
let rangeEndDate = '';

const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const money = value => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 2 }).format(Number(value) || 0);
const compactMoney = value => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', notation: 'compact', maximumFractionDigits: 1 }).format(Number(value) || 0);
const shortDate = value => value ? new Intl.DateTimeFormat('es-MX', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(`${value}T12:00:00`)).replace('.', '') : '—';
const localToday = () => {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};
rangeEndDate = localToday();
rangeStartDate = `${rangeEndDate.slice(0, 7)}-01`;
const escapeHtml = value => String(value ?? '').replace(/[&<>'"]/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
const getPeriod = () => periods.find(period => period.id === selectedPeriodId) || periods[0];
const getWeek = () => getPeriod().weeks[selectedWeekIndex] || getPeriod().weeks[0];
const periodRangeLabel = period => {
  const dates = period.weeks.flatMap(week => [week.start, week.end]).filter(Boolean).sort();
  if (!dates.length) return '';
  const start = new Date(`${dates[0]}T12:00:00`);
  const end = new Date(`${dates.at(-1)}T12:00:00`);
  const months = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const startMonth = months[start.getUTCMonth()];
  const endMonth = months[end.getUTCMonth()];
  const startYear = start.getUTCFullYear();
  const endYear = end.getUTCFullYear();
  if (startYear === endYear && start.getUTCMonth() === end.getUTCMonth()) return `${startMonth} ${startYear}`;
  if (startYear === endYear) return `${startMonth}–${endMonth} ${startYear}`;
  return `${startMonth} ${startYear}–${endMonth} ${endYear}`;
};
const orderedWeeks = period => period.weeks.map((week, index) => ({ week, index })).sort((a, b) => {
  return String(b.week.start || '').localeCompare(String(a.week.start || ''));
});
const getBudget = id => state.budgets[id] || { weekly: 0, monthly: 0 };
const getItem = id => budgetItems.find(item => item.id === id);
const budgetTotal = group => budgetItems.filter(item => !group || item.group === group).reduce((sum, item) => sum + Number(getBudget(item.id).weekly || 0), 0);
const monthlyTotal = group => budgetItems.filter(item => !group || item.group === group).reduce((sum, item) => sum + Number(getBudget(item.id).monthly || 0), 0);
const excelForSelection = () => excelEntries.filter(entry => entry.periodId === selectedPeriodId && Number(entry.weekIndex) === Number(selectedWeekIndex));
const manualForSelection = () => state.manualEntries.filter(entry => entry.periodId === selectedPeriodId && Number(entry.weekIndex) === Number(selectedWeekIndex));
const baseTotal = () => {
  const rows = excelForSelection();
  return rows.length ? rows.reduce((sum, entry) => sum + Number(entry.amount || 0), 0) : (getWeek().total || 0);
};
const selectedTotal = () => baseTotal() + manualForSelection().reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
const selectedMonthKey = () => String(getWeek().start || localToday()).slice(0, 7);
const monthYearLabel = monthKey => {
  const label = new Intl.DateTimeFormat('es-MX', { month: 'long', year: 'numeric' }).format(new Date(`${monthKey}-01T12:00:00`));
  return label.charAt(0).toUpperCase() + label.slice(1);
};

function monthlySpentForSelection() {
  const monthKey = selectedMonthKey();
  return periods.reduce((periodTotal, period) => periodTotal + period.weeks.reduce((weekTotal, week, weekIndex) => {
    const excelRows = excelEntries.filter(entry => entry.periodId === period.id && Number(entry.weekIndex) === weekIndex);
    const excelTotal = excelRows.length
      ? excelRows.filter(entry => String(entry.date || '').startsWith(monthKey)).reduce((sum, entry) => sum + Number(entry.amount || 0), 0)
      : (String(week.start || '').startsWith(monthKey) ? Number(week.total || 0) : 0);
    const manualTotal = state.manualEntries
      .filter(entry => entry.periodId === period.id && Number(entry.weekIndex) === weekIndex && String(entry.date || '').startsWith(monthKey))
      .reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
    return weekTotal + excelTotal + manualTotal;
  }, 0), 0);
}

const entryIsInRange = entry => entry.periodId === selectedPeriodId && entry.date >= rangeStartDate && entry.date <= rangeEndDate;
const dateRangeEntries = () => [...excelEntries, ...state.manualEntries].filter(entryIsInRange);
const dateRangeLabel = () => rangeStartDate === rangeEndDate ? shortDate(rangeStartDate) : `${shortDate(rangeStartDate)} – ${shortDate(rangeEndDate)}`;

function dateRangeTotal() {
  const period = getPeriod();
  const excelTotal = period.weeks.reduce((total, week, weekIndex) => {
    const rows = excelEntries.filter(entry => entry.periodId === period.id && Number(entry.weekIndex) === weekIndex);
    if (rows.length) return total + rows.filter(entryIsInRange).reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
    const overlapsRange = week.start <= rangeEndDate && week.end >= rangeStartDate;
    return total + (overlapsRange ? Number(week.total || 0) : 0);
  }, 0);
  const manualTotal = state.manualEntries.filter(entryIsInRange).reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
  return excelTotal + manualTotal;
}

function dateRangeBreakdown() {
  const rows = dateRangeEntries();
  if (!rows.length) return null;
  return rows.reduce((breakdown, entry) => {
    breakdown[entry.category] = (breakdown[entry.category] || 0) + Number(entry.amount || 0);
    return breakdown;
  }, {});
}

function loadState() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (saved?.budgets && Array.isArray(saved.manualEntries)) {
      let migrated = false;
      saved.manualEntries = saved.manualEntries.map(entry => {
        if (periods.some(period => period.id === entry.periodId)) return entry;
        const matchingPeriod = periods.find(period => period.weeks.some(week => entry.date && entry.date >= week.start && entry.date <= week.end));
        const period = matchingPeriod || periods.find(item => item.id === currentPeriodId);
        if (!period) return entry;
        const matchingWeekIndex = period.weeks.findIndex(week => entry.date && entry.date >= week.start && entry.date <= week.end);
        migrated = true;
        return { ...entry, periodId: period.id, weekIndex: matchingWeekIndex >= 0 ? matchingWeekIndex : 0 };
      });
      if (migrated) localStorage.setItem(STORAGE_KEY, JSON.stringify(saved));
      return saved;
    }
  } catch (error) { console.warn('No se pudo leer la sesión guardada', error); }
  return defaultState();
}

function saveState() { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }

function snapshotBreakdown(date = '') {
  const rows = [...excelForSelection(), ...manualForSelection()].filter(entry => !date || entry.date === date);
  if (!rows.length) return null;
  const breakdown = {};
  for (const entry of rows) breakdown[entry.category] = (breakdown[entry.category] || 0) + Number(entry.amount || 0);
  return breakdown;
}

function movementRows() {
  return dateRangeEntries()
    .map((row, index) => ({ row, index }))
    .sort((a, b) => String(b.row.date).localeCompare(String(a.row.date)) || b.index - a.index)
    .map(({ row }) => row);
}

function showToast(message) {
  const toast = $('#toast');
  toast.textContent = message;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 3200);
}

function setView(view) {
  activeView = view;
  $$('.view').forEach(section => section.classList.toggle('active-view', section.id === `${view}View`));
  $$('.nav-item').forEach(button => button.classList.toggle('active', button.dataset.view === view));
  if (view === 'capture') renderCapture();
  if (view === 'budget') renderBudget();
  if (view === 'history') renderHistory();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function setCaptureMode(mode) {
  activeCaptureMode = mode === 'bulk' ? 'bulk' : 'single';
  $$('[data-capture-mode]').forEach(button => {
    const isActive = button.dataset.captureMode === activeCaptureMode;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', String(isActive));
  });
  const singlePanel = $('#singleCapturePanel');
  const bulkPanel = $('#bulkCapturePanel');
  singlePanel.classList.toggle('is-active', activeCaptureMode === 'single');
  singlePanel.hidden = activeCaptureMode !== 'single';
  bulkPanel.classList.toggle('is-active', activeCaptureMode === 'bulk');
  bulkPanel.hidden = activeCaptureMode !== 'bulk';
}

function renderSelectors() {
  const currentPeriod = getPeriod();
  setPeriodMenuOpen(false);
  $('#periodSelectValue').textContent = currentPeriod.name;
  $('#periodSelectMenu').innerHTML = periods.map(period => `<button type="button" class="custom-option period-option ${period.id === selectedPeriodId ? 'selected' : ''}" role="option" aria-selected="${period.id === selectedPeriodId}" data-period-id="${period.id}"><span><strong>${escapeHtml(period.name)}</strong><small>${escapeHtml(periodRangeLabel(period))}</small></span><span class="period-check">${period.id === selectedPeriodId ? '✓' : ''}</span></button>`).join('');
  const ordered = orderedWeeks(getPeriod());
  const weekOptions = ordered.map(({ week, index }) => `<option value="${index}" ${index === selectedWeekIndex ? 'selected' : ''}>${escapeHtml(week.label)} · ${money(week.total)}</option>`).join('');
  $('#captureWeek').innerHTML = weekOptions;
  $('#bulkCaptureWeek').innerHTML = weekOptions;
}

function renderDashboard() {
  const period = getPeriod();
  const week = getWeek();
  $('#dashboardSubtitle').textContent = `${period.name} · ${period.sheet}`;
  $('#rangeStartDate').value = rangeStartDate;
  $('#rangeEndDate').value = rangeEndDate;
  $('#kpiSpent').textContent = money(selectedTotal());
  $('#kpiSpentNote').textContent = `${period.sheet} · ${week.label}`;
  $('#kpiMonthlySpent').textContent = money(monthlySpentForSelection());
  $('#kpiMonthlySpentNote').textContent = monthYearLabel(selectedMonthKey());
  $('#conceptCaption').textContent = dateRangeLabel();
  $('#spendingPieCaption').textContent = dateRangeLabel();
  renderConceptBars();
  renderSpendingPie();
  renderMovementTable();
}

function renderConceptBars() {
  const breakdown = dateRangeBreakdown();
  const container = $('#conceptBars');
  if (!breakdown) {
    container.innerHTML = `<div class="empty-row">No hay movimientos con desglose en el rango seleccionado.</div>`;
    return;
  }
  const rows = Object.entries(breakdown).filter(([, amount]) => amount > 0).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const max = rows[0]?.[1] || 1;
  container.innerHTML = rows.map(([id, amount]) => `<div class="concept-row"><span class="concept-name" title="${escapeHtml(getItem(id)?.name || id)}">${escapeHtml(getItem(id)?.name || id)}</span><div class="concept-track"><div class="concept-fill" style="width:${Math.max(3, (amount / max) * 100)}%"></div></div><span class="concept-amount">${compactMoney(amount)}</span></div>`).join('');
}

function renderSpendingPie() {
  const breakdown = dateRangeBreakdown();
  const container = $('#spendingPieContent');
  const rows = Object.entries(breakdown || {}).filter(([, amount]) => amount > 0).sort((a, b) => b[1] - a[1]);
  if (!rows.length) {
    container.innerHTML = `<div class="empty-row spending-pie-empty">La gráfica aparecerá cuando registres gastos en este rango.</div>`;
    return;
  }

  const visibleRows = rows.slice(0, 6);
  if (rows.length > 6) visibleRows.push(['other', rows.slice(6).reduce((sum, [, amount]) => sum + amount, 0)]);
  const total = visibleRows.reduce((sum, [, amount]) => sum + amount, 0);
  let accumulated = 0;
  const segments = visibleRows.map(([id, amount], index) => {
    const start = accumulated / total * 100;
    accumulated += amount;
    const end = accumulated / total * 100;
    return { id, amount, color: spendingPieColors[index % spendingPieColors.length], start, end };
  });
  const gradient = segments.map(segment => `${segment.color} ${segment.start}% ${segment.end}%`).join(', ');
  const accessibleSummary = segments.map(segment => `${segment.id === 'other' ? 'Otros' : getItem(segment.id)?.name || segment.id}: ${money(segment.amount)}`).join(', ');

  container.innerHTML = `<div class="spending-pie-layout"><div class="spending-pie" role="img" aria-label="Distribución del gasto. ${escapeHtml(accessibleSummary)}" style="background:conic-gradient(${gradient})"><div class="spending-pie-center"><span>Total</span><strong>${money(total)}</strong></div></div><div class="spending-pie-legend">${segments.map((segment, index) => { const name = segment.id === 'other' ? 'Otros' : getItem(segment.id)?.name || segment.id; const percentage = Math.round(segment.amount / total * 100); return `<div class="spending-pie-legend-row ${index === 0 ? 'is-leading' : ''}"><span class="spending-pie-dot" style="background:${segment.color}"></span><span class="spending-pie-name" title="${escapeHtml(name)}">${escapeHtml(name)}</span><strong>${percentage}%</strong>${index === 0 ? '<small>Mayor gasto</small>' : ''}</div>`; }).join('')}</div></div>`;
}

function renderMovementTable() {
  const rows = movementRows().slice(0, 8);
  $('#movementRows').innerHTML = rows.length ? rows.map(row => movementRowHtml(row, false, true)).join('') : `<tr><td colspan="6" class="empty-row">No hay movimientos en el rango seleccionado.</td></tr>`;
}

function movementRowHtml(row, includePayment, includeSpender = includePayment) {
  const item = getItem(row.category);
  return `<tr><td class="date-cell">${shortDate(row.date)}</td><td><span class="tag">${escapeHtml(item?.name || row.category)}</span></td>${includeSpender ? `<td>${escapeHtml(row.spender || '—')}</td>` : ''}${includePayment ? `<td>${escapeHtml(row.payment || '—')}</td>` : ''}<td class="note-cell">${escapeHtml(row.note || 'Sin nota')}</td><td class="align-right amount-cell">${money(row.amount)}</td><td>${row.source === 'manual' ? `<button class="delete-button" data-delete-id="${row.id}" aria-label="Eliminar gasto">×</button>` : ''}</td></tr>`;
}

function syncExpenseTypeFromCategory() {
  const selectedItem = getItem($('#expenseCategory').value);
  $('#expenseType').value = selectedItem?.group === 'fixed' ? 'Fijo' : 'Operativo';
}

function syncBulkExpenseTypeFromCategory() {
  const selectedItem = getItem($('#bulkExpenseCategory').value);
  $('#bulkExpenseType').value = selectedItem?.group === 'fixed' ? 'Fijo' : 'Operativo';
}

function renderExpenseCategories() {
  const selectedCategory = $('#expenseCategory').value;
  const selectedBulkCategory = $('#bulkExpenseCategory').value;
  const options = budgetItems.map(item => `<option value="${item.id}">${escapeHtml(item.name)}</option>`).join('');
  $('#expenseCategory').innerHTML = options;
  $('#bulkExpenseCategory').innerHTML = options;
  if (budgetItems.some(item => item.id === selectedCategory)) $('#expenseCategory').value = selectedCategory;
  if (budgetItems.some(item => item.id === selectedBulkCategory)) $('#bulkExpenseCategory').value = selectedBulkCategory;
  syncExpenseTypeFromCategory();
  syncBulkExpenseTypeFromCategory();
}

function bulkDraftRowHtml(row) {
  const item = getItem(row.category);
  const period = periods.find(candidate => candidate.id === row.periodId);
  const week = period?.weeks[row.weekIndex];
  return `<tr><td class="date-cell">${shortDate(row.date)}</td><td>${escapeHtml(week?.label || '—')}</td><td><span class="tag">${escapeHtml(item?.name || row.category)}</span></td><td>${escapeHtml(row.expenseType)}</td><td>${escapeHtml(row.spender)}</td><td>${escapeHtml(row.payment)}</td><td class="note-cell">${escapeHtml(row.note)}</td><td class="align-right amount-cell">${money(row.amount)}</td><td><button class="delete-button" type="button" data-delete-bulk-id="${row.id}" aria-label="Quitar gasto de la lista">×</button></td></tr>`;
}

function renderBulkDrafts() {
  const count = bulkDraftEntries.length;
  const total = bulkDraftEntries.reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
  $('#bulkTabCount').textContent = count;
  $('#bulkEntryCount').textContent = `${count} ${count === 1 ? 'gasto' : 'gastos'}`;
  $('#bulkEntryTotal').textContent = money(total);
  $('#clearBulkExpenses').disabled = count === 0;
  $('#saveBulkExpenses').disabled = count === 0;
  $('#bulkEntryList').innerHTML = count
    ? `<div class="table-scroll"><table><thead><tr><th>Fecha</th><th>Semana</th><th>Concepto</th><th>Tipo</th><th>Realizó el gasto</th><th>Pago</th><th>Nota</th><th class="align-right">Monto</th><th></th></tr></thead><tbody>${bulkDraftEntries.map(bulkDraftRowHtml).join('')}</tbody></table></div>`
    : `<div class="bulk-empty-state">Aún no hay gastos en la lista. Completa el formulario y selecciona “Agregar a la lista”.</div>`;
}

function renderCapture() {
  renderExpenseCategories();
  $('#captureWeeklySpent').textContent = money(selectedTotal());
  const entries = [...manualForSelection()].sort((a, b) => String(b.date).localeCompare(String(a.date)));
  const count = entries.length;
  $('#newMovementCount').textContent = `${count} ${count === 1 ? 'registro' : 'registros'}`;
  $('#newMovementRows').innerHTML = entries.length ? entries.map(row => movementRowHtml(row, true)).join('') : `<tr><td colspan="7" class="empty-row">Los gastos que guardes para esta semana aparecerán aquí.</td></tr>`;
  if (!$('#expenseDate').value) $('#expenseDate').value = localToday();
  if (!$('#bulkExpenseDate').value) $('#bulkExpenseDate').value = localToday();
  renderBulkDrafts();
  setCaptureMode(activeCaptureMode);
}

function renderBudget() {
  const breakdown = snapshotBreakdown() || {};
  const todayBreakdown = snapshotBreakdown(localToday()) || {};
  const operating = budgetItems.filter(item => item.group === 'operating');
  const fixed = budgetItems.filter(item => item.group === 'fixed');
  $('#monthlyBudgetTotal').textContent = money(monthlyTotal());
  $('#operatingWeekly').textContent = money(budgetTotal('operating'));
  $('#fixedWeekly').textContent = money(budgetTotal('fixed'));
  $('#budgetGrandMonthly').textContent = money(monthlyTotal());
  $('#operatingBudgetRows').innerHTML = operating.map(item => budgetRowHtml(item, breakdown[item.id] || 0, todayBreakdown[item.id] || 0)).join('');
  $('#fixedBudgetRows').innerHTML = fixed.map(item => budgetRowHtml(item, breakdown[item.id] || 0, todayBreakdown[item.id] || 0, true)).join('');
  $('#operatingBudgetTotal').innerHTML = totalRowHtml(budgetTotal('operating'), monthlyTotal('operating'), operating.reduce((sum, item) => sum + (breakdown[item.id] || 0), 0), operating.reduce((sum, item) => sum + (todayBreakdown[item.id] || 0), 0), true);
  $('#fixedBudgetTotal').innerHTML = totalRowHtml(budgetTotal('fixed'), monthlyTotal('fixed'), fixed.reduce((sum, item) => sum + (breakdown[item.id] || 0), 0), fixed.reduce((sum, item) => sum + (todayBreakdown[item.id] || 0), 0), false);
}

function budgetRowHtml(item, spent, spentToday, fixed = false) {
  const budget = getBudget(item.id);
  return `<tr><td><strong>${escapeHtml(item.name)}</strong></td><td class="align-right daily-amount">${money(spentToday)}</td><td class="align-right">${money(budget.weekly)}</td><td class="align-right">${money(budget.monthly)}</td>${fixed ? '' : `<td class="align-right amount-cell">${money(spent)}</td>`}</tr>`;
}

function totalRowHtml(weekly, monthly, spent, spentToday, showSpent) {
  return `<tr class="total-row"><td>Total</td><td class="align-right">${money(spentToday)}</td><td class="align-right">${money(weekly)}</td><td class="align-right">${money(monthly)}</td>${showSpent ? `<td class="align-right">${money(spent)}</td>` : ''}</tr>`;
}

function renderHistory() {
  const allWeeks = periods.flatMap(period => period.weeks.map((week, index) => ({ period, week, index, total: period.id === selectedPeriodId && index === selectedWeekIndex ? selectedTotal() : week.total })));
  const sortedWeeks = allWeeks.slice().sort((a, b) => String(b.week.start || '').localeCompare(String(a.week.start || '')));
  const chartWeeks = sortedWeeks.filter(item => item.total > 0).slice(0, 18).reverse();
  const max = Math.max(...chartWeeks.map(item => item.total), 1);
  $('#historyChart').innerHTML = chartWeeks.map(item => `<div class="chart-column"><div class="chart-bar-wrap"><div class="chart-bar" style="height:${Math.max(3, item.total / max * 100)}%" data-value="${money(item.total)}"></div></div><span class="chart-label" title="${escapeHtml(item.period.name)} · ${escapeHtml(item.week.label)}">${escapeHtml(item.week.label)}</span></div>`).join('');
  $('#historyRows').innerHTML = sortedWeeks.map(item => {
    const budget = budgetTotal(); const difference = budget - item.total; const status = item.total === 0 ? ['Sem dados', 'neutral'] : difference < 0 ? ['Excedido', 'over'] : ['En rango', ''];
    return `<tr><td><span class="tag">${escapeHtml(item.period.sheet)}</span></td><td>${escapeHtml(item.week.label)}</td><td class="align-right amount-cell">${money(item.total)}</td><td class="align-right">${money(budget)}</td><td class="align-right" style="color:${difference < 0 ? '#a34641' : 'inherit'}">${difference >= 0 ? '+' : ''}${money(difference)}</td><td><span class="status-pill ${status[1]}">${status[0]}</span></td></tr>`;
  }).join('');
}

function setPeriodMenuOpen(isOpen) {
  $('#periodSelectControl').classList.toggle('open', isOpen);
  $('#periodSelectButton').setAttribute('aria-expanded', String(isOpen));
}

function downloadCsv(filename, rows) {
  const csv = rows.map(row => row.map(value => `"${String(value ?? '').replace(/"/g, '""')}"`).join(',')).join('\n');
  const link = document.createElement('a');
  link.href = URL.createObjectURL(new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' }));
  link.download = filename; link.click(); URL.revokeObjectURL(link.href);
}

function bindEvents() {
  document.addEventListener('click', event => {
    const viewButton = event.target.closest('[data-view]');
    if (viewButton) setView(viewButton.dataset.view);
    const deleteButton = event.target.closest('[data-delete-id]');
    if (deleteButton) {
      state.manualEntries = state.manualEntries.filter(entry => entry.id !== deleteButton.dataset.deleteId);
      saveState(); renderDashboard(); renderCapture(); renderBudget(); showToast('Gasto eliminado.');
    }
    const deleteBulkButton = event.target.closest('[data-delete-bulk-id]');
    if (deleteBulkButton) {
      bulkDraftEntries = bulkDraftEntries.filter(entry => entry.id !== deleteBulkButton.dataset.deleteBulkId);
      renderBulkDrafts();
      showToast('Gasto retirado de la lista.');
    }
  });
  $$('[data-capture-mode]').forEach(button => button.addEventListener('click', () => setCaptureMode(button.dataset.captureMode)));
  $('#periodSelectButton').addEventListener('click', event => { event.stopPropagation(); setPeriodMenuOpen(!$('#periodSelectControl').classList.contains('open')); });
  $('#periodSelectMenu').addEventListener('click', event => {
    const option = event.target.closest('[data-period-id]');
    if (!option) return;
    selectedPeriodId = option.dataset.periodId;
    selectedWeekIndex = 0;
    setPeriodMenuOpen(false);
    renderAll();
  });
  document.addEventListener('click', event => {
    if (!event.target.closest('#periodSelectControl')) setPeriodMenuOpen(false);
  });
  $('#rangeStartDate').addEventListener('change', event => {
    rangeStartDate = event.target.value || `${localToday().slice(0, 7)}-01`;
    if (rangeStartDate > rangeEndDate) rangeEndDate = rangeStartDate;
    renderDashboard();
  });
  $('#rangeEndDate').addEventListener('change', event => {
    rangeEndDate = event.target.value || localToday();
    if (rangeEndDate < rangeStartDate) rangeStartDate = rangeEndDate;
    renderDashboard();
  });
  $('#captureWeek').addEventListener('change', event => { selectedWeekIndex = Number(event.target.value); renderSelectors(); renderCapture(); });
  $('#expenseCategory').addEventListener('change', syncExpenseTypeFromCategory);
  $('#bulkCaptureWeek').addEventListener('change', event => { selectedWeekIndex = Number(event.target.value); renderSelectors(); renderCapture(); });
  $('#bulkExpenseCategory').addEventListener('change', syncBulkExpenseTypeFromCategory);

  $('#expenseForm').addEventListener('submit', event => {
    event.preventDefault();
    const amount = Number($('#expenseAmount').value);
    if (!amount || amount <= 0) return showToast('Escribe un monto mayor a cero.');
    const category = $('#expenseCategory').value;
    
    state.manualEntries.push({ 
      id: `manual-${Date.now()}`, 
      date: $('#expenseDate').value, 
      category,
      amount, 
      note: $('#expenseNote').value.trim() || 'Sin nota', 
      payment: $('#expensePayment').value,
      spender: $('#expenseSpender').value.trim(), // Nuevo campo guardado
      expenseType: getItem(category)?.group === 'fixed' ? 'Fijo' : 'Operativo',
      periodId: selectedPeriodId, 
      weekIndex: selectedWeekIndex, 
      source: 'manual' 
    });
    
    saveState(); 
    event.target.reset(); 
    $('#expenseDate').value = localToday();
    renderAll(); 
    showToast('Gasto guardado y totales actualizados.'); 
  });

  $('#bulkExpenseForm').addEventListener('submit', event => {
    event.preventDefault();
    const amount = Number($('#bulkExpenseAmount').value);
    if (!amount || amount <= 0) return showToast('Escribe un monto mayor a cero.');
    const date = $('#bulkExpenseDate').value;
    const weekIndex = Number($('#bulkCaptureWeek').value);
    const category = $('#bulkExpenseCategory').value;
    const spender = $('#bulkExpenseSpender').value.trim();
    const payment = $('#bulkExpensePayment').value;

    bulkDraftEntries.push({
      id: `bulk-${Date.now()}-${bulkDraftEntries.length}`,
      date,
      category,
      amount,
      note: $('#bulkExpenseNote').value.trim() || 'Sin nota',
      payment,
      spender,
      expenseType: getItem(category)?.group === 'fixed' ? 'Fijo' : 'Operativo',
      periodId: selectedPeriodId,
      weekIndex,
      source: 'manual',
    });

    event.target.reset();
    $('#bulkExpenseDate').value = date;
    $('#bulkCaptureWeek').value = String(weekIndex);
    $('#bulkExpenseSpender').value = spender;
    $('#bulkExpensePayment').value = payment;
    renderExpenseCategories();
    renderBulkDrafts();
    showToast('Gasto agregado a la lista.');
  });

  $('#clearBulkExpenses').addEventListener('click', () => {
    bulkDraftEntries = [];
    renderBulkDrafts();
    showToast('Lista de gastos vaciada.');
  });

  $('#saveBulkExpenses').addEventListener('click', () => {
    if (!bulkDraftEntries.length) return;
    const entriesToSave = [...bulkDraftEntries];
    const idBase = Date.now();
    state.manualEntries.push(...entriesToSave.map((entry, index) => ({ ...entry, id: `manual-${idBase + index}`, source: 'manual' })));
    bulkDraftEntries = [];
    saveState();
    renderAll();
    showToast(`${entriesToSave.length} ${entriesToSave.length === 1 ? 'gasto guardado' : 'gastos guardados'} correctamente.`);
  });
  $('#exportCsv').addEventListener('click', () => { 
    const rows = [['Fecha', 'Concepto', 'Tipo de gasto', 'Quién realizó el gasto', 'Nota', 'Forma de pago', 'Monto']].concat(movementRows().map(row => [
      row.date, 
      getItem(row.category)?.name || row.category, 
      row.expenseType || (getItem(row.category)?.group === 'fixed' ? 'Fijo' : 'Operativo'),
      row.spender || '—',
      row.note, 
      row.payment || 'Excel', 
      row.amount
    ])); 
    downloadCsv(`control-gastos-${selectedPeriodId}.csv`, rows); 
    showToast('CSV descargado.'); 
  });

  $('#historyExport').addEventListener('click', () => { const rows = [['Hoja', 'Semana', 'Gasto', 'Presupuesto semanal', 'Variación']]; periods.forEach(period => period.weeks.forEach((week, index) => rows.push([period.sheet, week.label, index === selectedWeekIndex && period.id === selectedPeriodId ? selectedTotal() : week.total, budgetTotal(), budgetTotal() - week.total]))); downloadCsv('historico-control-gastos.csv', rows); showToast('Histórico descargado.'); });
  $('#resetData').addEventListener('click', () => { if (!window.confirm('¿Restaurar los datos demo y borrar los gastos capturados?')) return; state = defaultState(); bulkDraftEntries = []; activeCaptureMode = 'single'; saveState(); selectedPeriodId = currentPeriodId; selectedWeekIndex = 0; renderAll(); showToast('Datos demo restaurados.'); });
}

function renderAll() { renderSelectors(); renderDashboard(); if (activeView === 'capture') renderCapture(); if (activeView === 'budget') renderBudget(); if (activeView === 'history') renderHistory(); }

bindEvents();
$('#expenseDate').value = localToday();
$('#bulkExpenseDate').value = localToday();
renderAll();
