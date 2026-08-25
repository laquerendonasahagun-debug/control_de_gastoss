const STORAGE_KEY = 'la-querendona-control-gastos-v1';
const EXPENSES_API = '/api/expenses';
const AUTH_API = '/api/auth';

const expenseItems = [
  { id: 'abarrote', name: 'Abarrote', group: 'operating' },
  { id: 'verdura', name: 'Verdura / chiles secos / hierbas de olor', group: 'operating' },
  { id: 'pan', name: 'Pan', group: 'operating' },
  { id: 'basura', name: 'Basura', group: 'operating' },
  { id: 'agua', name: 'Agua', group: 'operating' },
  { id: 'limpieza', name: 'Producto limpieza y mantelería', group: 'operating' },
  { id: 'gas', name: 'Gas', group: 'operating' },
  { id: 'carne', name: 'Carne', group: 'operating' },
  { id: 'coca-cola', name: 'Coca-Cola', group: 'operating' },
  { id: 'jarritos', name: 'Jarritos', group: 'operating' },
  { id: 'cortes', name: 'Cortes, snacks', group: 'operating' },
  { id: 'cerveza', name: 'Cerveza', group: 'operating' },
  { id: 'leche', name: 'Leche', group: 'operating' },
  { id: 'pollo', name: 'Pollo', group: 'operating' },
  { id: 'bistek', name: 'Bistek', group: 'operating' },
  { id: 'cremeria', name: 'Cremería', group: 'operating' },
  { id: 'tortillas', name: 'Tortillas y masa', group: 'operating' },
  { id: 'mandaditos', name: 'Mandaditos', group: 'operating' },
  { id: 'desechable', name: 'Desechable', group: 'operating' },
  { id: 'vinos', name: 'Vinos y licores', group: 'operating' },
  { id: 'papeleria', name: 'Papelería', group: 'operating' },
  { id: 'comision', name: 'Comisión billipocket', group: 'operating' },
  { id: 'caja', name: 'Caja', group: 'operating' },
  { id: 'otros', name: 'Otros', group: 'operating' },
  { id: 'nomina', name: 'Nómina', group: 'fixed' },
  { id: 'renta', name: 'Renta', group: 'fixed' },
  { id: 'luz', name: 'Luz', group: 'fixed' },
  { id: 'gerencia', name: 'Gerencia', group: 'fixed' },
  { id: 'reserva', name: 'Fondo de reserva', group: 'fixed' },
];
const expenseSpenders = ['Alejandra', 'Horacio', 'Diego', 'Haytham', 'Mary'];

const spendingPieColors = ['#24584a', '#8cbf8d', '#c17db9', '#83cfc5', '#e4a84c', '#7b8fc5', '#d4776a'];

const periods = (window.EXCEL_PERIODS || []).slice().sort((a, b) => Number(Boolean(b.current)) - Number(Boolean(a.current)) || String(b.weeks.at(-1)?.end || '').localeCompare(String(a.weeks.at(-1)?.end || '')));
const currentPeriodId = periods.find(period => period.current)?.id || periods[0]?.id || '';

const excelEntries = window.EXCEL_ENTRIES || [];
let pendingLegacyEntries = loadLegacyEntries();
let state = { manualEntries: [...pendingLegacyEntries] };
let selectedPeriodId = currentPeriodId;
let selectedWeekIndex = 0;
let activeView = 'dashboard';
let toastTimer;
let databaseReady = false;
let databaseSyncing = false;
let activeCaptureMode = 'single';
let bulkDraftEntries = [];
let editingExpenseId = '';
let rangeStartDate = '';
let rangeEndDate = '';
let activeDateFilter = 'month';
let selectedFilterDay = '';
let selectedFilterWeek = '';
let selectedFilterMonth = '';
let currentUser = null;

const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const money = value => new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 2 }).format(Number(value) || 0);
const newExpenseId = () => `manual-${globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`}`;
const shortDate = value => value ? new Intl.DateTimeFormat('es-MX', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(`${value}T12:00:00`)).replace('.', '') : '—';
const localToday = () => {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};
const dateToIsoWeek = value => {
  const source = new Date(`${value}T12:00:00`);
  const date = new Date(Date.UTC(source.getFullYear(), source.getMonth(), source.getDate()));
  const dayNumber = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNumber);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNumber = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${date.getUTCFullYear()}-W${String(weekNumber).padStart(2, '0')}`;
};
const isoWeekRange = value => {
  const match = /^(\d{4})-W(\d{2})$/.exec(value || '');
  if (!match) return { start: localToday(), end: localToday() };
  const year = Number(match[1]);
  const week = Number(match[2]);
  const januaryFourth = new Date(Date.UTC(year, 0, 4));
  const januaryFourthDay = januaryFourth.getUTCDay() || 7;
  const monday = new Date(januaryFourth);
  monday.setUTCDate(januaryFourth.getUTCDate() - januaryFourthDay + 1 + ((week - 1) * 7));
  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);
  const format = date => date.toISOString().slice(0, 10);
  return { start: format(monday), end: format(sunday) };
};
const monthRange = value => {
  const monthKey = /^\d{4}-\d{2}$/.test(value || '') ? value : localToday().slice(0, 7);
  const [year, month] = monthKey.split('-').map(Number);
  const lastDay = new Date(year, month, 0).getDate();
  const endOfMonth = `${monthKey}-${String(lastDay).padStart(2, '0')}`;
  return { start: `${monthKey}-01`, end: monthKey === localToday().slice(0, 7) ? localToday() : endOfMonth };
};
selectedFilterDay = localToday();
selectedFilterWeek = dateToIsoWeek(localToday());
selectedFilterMonth = localToday().slice(0, 7);
({ start: rangeStartDate, end: rangeEndDate } = monthRange(selectedFilterMonth));
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
const getItem = id => expenseItems.find(item => item.id === id);
const excelForSelection = () => excelEntries.filter(entry => entry.periodId === selectedPeriodId && Number(entry.weekIndex) === Number(selectedWeekIndex));
const manualForSelection = () => state.manualEntries.filter(entry => entry.periodId === selectedPeriodId && Number(entry.weekIndex) === Number(selectedWeekIndex));
const weekActualTotal = (periodId, weekIndex, week) => {
  const excelRows = excelEntries.filter(entry => entry.periodId === periodId && Number(entry.weekIndex) === Number(weekIndex));
  const importedTotal = excelRows.length ? excelRows.reduce((sum, entry) => sum + Number(entry.amount || 0), 0) : Number(week.total || 0);
  const capturedTotal = state.manualEntries
    .filter(entry => entry.periodId === periodId && Number(entry.weekIndex) === Number(weekIndex))
    .reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
  return importedTotal + capturedTotal;
};
const selectedTotal = () => weekActualTotal(selectedPeriodId, selectedWeekIndex, getWeek());
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

function monthlyBreakdownForSelection() {
  const monthKey = selectedMonthKey();
  return [...excelEntries, ...state.manualEntries]
    .filter(entry => String(entry.date || '').startsWith(monthKey))
    .reduce((breakdown, entry) => {
      breakdown[entry.category] = (breakdown[entry.category] || 0) + Number(entry.amount || 0);
      return breakdown;
    }, {});
}

const breakdownTotal = (items, breakdown) => items.reduce((sum, item) => sum + Number(breakdown[item.id] || 0), 0);

const entryIsInRange = entry => entry.periodId === selectedPeriodId && entry.date >= rangeStartDate && entry.date <= rangeEndDate;
const dateRangeEntries = () => [...excelEntries, ...state.manualEntries].filter(entryIsInRange);
const dateSpanLabel = (start, end) => start === end ? shortDate(start) : `${shortDate(start)} – ${shortDate(end)}`;
const dateRangeLabel = () => dateSpanLabel(rangeStartDate, rangeEndDate);
const dashboardTotalForRange = (start, end) => [...excelEntries, ...state.manualEntries]
  .filter(entry => entry.periodId === selectedPeriodId && String(entry.date || '') >= start && String(entry.date || '') <= end)
  .reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
const dashboardDailyTotal = () => dashboardTotalForRange(selectedFilterDay, selectedFilterDay);
const dashboardWeeklyRange = () => isoWeekRange(selectedFilterWeek);
const dashboardWeeklyTotal = () => {
  const { start, end } = dashboardWeeklyRange();
  return dashboardTotalForRange(start, end);
};
const dashboardMonthlyTotal = () => {
  const { start, end } = monthRange(selectedFilterMonth);
  return dashboardTotalForRange(start, end);
};

function applyDateFilter() {
  if (activeDateFilter === 'day') {
    rangeStartDate = selectedFilterDay;
    rangeEndDate = selectedFilterDay;
  } else if (activeDateFilter === 'week') {
    ({ start: rangeStartDate, end: rangeEndDate } = isoWeekRange(selectedFilterWeek));
  } else {
    ({ start: rangeStartDate, end: rangeEndDate } = monthRange(selectedFilterMonth));
  }
}

function renderDateFilter() {
  $$('[data-date-filter]').forEach(button => {
    const isActive = button.dataset.dateFilter === activeDateFilter;
    button.classList.toggle('is-active', isActive);
    button.setAttribute('aria-selected', String(isActive));
  });
  $$('[data-date-control]').forEach(control => { control.hidden = control.dataset.dateControl !== activeDateFilter; });
  $('#dayFilterDate').value = selectedFilterDay;
  $('#weekFilterWeek').value = selectedFilterWeek;
  $('#monthFilterMonth').value = selectedFilterMonth;
  $('#dateFilterRange').textContent = dateRangeLabel();
}

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

function loadLegacyEntries() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
    if (Array.isArray(saved?.manualEntries)) {
      return saved.manualEntries.map(entry => {
        if (periods.some(period => period.id === entry.periodId)) return entry;
        const matchingPeriod = periods.find(period => period.weeks.some(week => entry.date && entry.date >= week.start && entry.date <= week.end));
        const period = matchingPeriod || periods.find(item => item.id === currentPeriodId);
        if (!period) return entry;
        const matchingWeekIndex = period.weeks.findIndex(week => entry.date && entry.date >= week.start && entry.date <= week.end);
        return { ...entry, periodId: period.id, weekIndex: matchingWeekIndex >= 0 ? matchingWeekIndex : 0 };
      });
    }
  } catch (error) { console.warn('No se pudo leer la sesión guardada', error); }
  return [];
}

async function apiRequest(options = {}) {
  const url = options.id ? `${EXPENSES_API}?id=${encodeURIComponent(options.id)}` : EXPENSES_API;
  const response = await fetch(url, {
    method: options.method || 'GET',
    headers: options.body ? { 'Content-Type': 'application/json' } : undefined,
    body: options.body ? JSON.stringify(options.body) : undefined,
    cache: 'no-store',
  });
  const payload = await response.json().catch(() => ({}));
  if (response.status === 401) showLogin(payload.error || 'Inicia sesión nuevamente.');
  if (!response.ok) throw new Error(payload.error || 'No fue posible conectar con la base de datos.');
  return payload;
}

async function authRequest(options = {}) {
  const response = await fetch(AUTH_API, {
    method: options.method || 'GET',
    headers: options.body ? { 'Content-Type': 'application/json' } : undefined,
    body: options.body ? JSON.stringify(options.body) : undefined,
    cache: 'no-store',
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || 'No fue posible iniciar sesión.');
  return payload;
}

function showLogin(message = '') {
  currentUser = null;
  databaseReady = false;
  state.manualEntries = [];
  closeExpenseEdit();
  $('#appShell').hidden = true;
  $('#loginScreen').hidden = false;
  $('#loginError').textContent = message;
  $('#loginPassword').value = '';
  window.setTimeout(() => $('#loginUsername').focus(), 0);
}

function applyRoleVisibility() {
  const isAdmin = currentUser?.role === 'admin';
  $$('[data-admin-only]').forEach(element => { element.hidden = !isAdmin; });
  $('#appShell').dataset.role = currentUser?.role || '';
  $('#sessionDisplayName').textContent = currentUser?.displayName || 'Usuario';
  $('#sessionRole').textContent = isAdmin ? 'Acceso completo' : 'Solo captura de gastos';
}

async function startSession(user) {
  currentUser = user;
  applyRoleVisibility();
  $('#loginScreen').hidden = true;
  $('#appShell').hidden = false;
  activeView = user.role === 'admin' ? 'dashboard' : 'capture';
  databaseReady = user.role === 'employee';
  renderAll();
  setView(activeView);
  if (user.role === 'admin') await syncExpenses({ migrateLegacy: true, silent: true });
}

async function initializeSession() {
  try {
    const payload = await authRequest();
    await startSession(payload.user);
  } catch (error) {
    showLogin('');
  }
}

function setDatabaseStatus(status, label) {
  const statusElement = $('#databaseStatus');
  const statusDot = $('#databaseStatusDot');
  if (statusElement) statusElement.textContent = label;
  if (statusDot) statusDot.dataset.status = status;
}

async function syncExpenses({ migrateLegacy = true, silent = false } = {}) {
  if (currentUser?.role !== 'admin') return;
  if (databaseSyncing) return;
  databaseSyncing = true;
  setDatabaseStatus('syncing', 'Sincronizando…');

  try {
    let payload = await apiRequest();
    if (migrateLegacy && pendingLegacyEntries.length) {
      payload = await apiRequest({ method: 'POST', body: { entries: pendingLegacyEntries } });
      pendingLegacyEntries = [];
      localStorage.removeItem(STORAGE_KEY);
    }
    state.manualEntries = Array.isArray(payload.entries) ? payload.entries : [];
    databaseReady = true;
    setDatabaseStatus('online', 'Base de datos conectada');
    renderAll();
    if (!silent) showToast('Gastos sincronizados con Neon.');
  } catch (error) {
    databaseReady = false;
    setDatabaseStatus('offline', 'Sin conexión a la base de datos');
    console.error('No se pudieron sincronizar los gastos:', error);
    if (!silent) showToast(error.message);
  } finally {
    databaseSyncing = false;
  }
}

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
  if (currentUser?.role !== 'admin' && view !== 'capture') return;
  activeView = view;
  $$('.view').forEach(section => section.classList.toggle('active-view', section.id === `${view}View`));
  $$('.nav-item').forEach(button => button.classList.toggle('active', button.dataset.view === view));
  if (view === 'capture') renderCapture();
  if (view === 'expenses') renderExpenses();
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
  const weekOptions = ordered.map(({ week, index }) => `<option value="${index}" ${index === selectedWeekIndex ? 'selected' : ''}>${escapeHtml(week.label)} · ${money(weekActualTotal(currentPeriod.id, index, week))}</option>`).join('');
  $('#captureWeek').innerHTML = weekOptions;
  $('#bulkCaptureWeek').innerHTML = weekOptions;
}

function renderDashboard() {
  const period = getPeriod();
  const weeklyRange = dashboardWeeklyRange();
  $('#dashboardSubtitle').textContent = `${period.name} · ${period.sheet}`;
  renderDateFilter();
  $('#kpiDailySpent').textContent = money(dashboardDailyTotal());
  $('#kpiDailySpentNote').textContent = `${selectedFilterDay === localToday() ? 'Hoy' : 'Día seleccionado'} · ${shortDate(selectedFilterDay)}`;
  $('#kpiSpent').textContent = money(dashboardWeeklyTotal());
  $('#kpiSpentNote').textContent = `${period.sheet} · ${dateSpanLabel(weeklyRange.start, weeklyRange.end)}`;
  $('#kpiMonthlySpent').textContent = money(dashboardMonthlyTotal());
  $('#kpiMonthlySpentNote').textContent = monthYearLabel(selectedFilterMonth);
  $('#conceptCaption').textContent = dateRangeLabel();
  $('#spendingPieCaption').textContent = dateRangeLabel();
  $('#operatingPieCaption').textContent = dateRangeLabel();
  $('#fixedPieCaption').textContent = dateRangeLabel();
  renderConceptBars();
  renderSpendingPie('spendingPieContent');
  renderSpendingPie('operatingPieContent', 'operating');
  renderSpendingPie('fixedPieContent', 'fixed');
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
  container.innerHTML = rows.map(([id, amount]) => `<div class="concept-row"><span class="concept-name" title="${escapeHtml(getItem(id)?.name || id)}">${escapeHtml(getItem(id)?.name || id)}</span><div class="concept-track"><div class="concept-fill" style="width:${Math.max(3, (amount / max) * 100)}%"></div></div><span class="concept-amount">${money(amount)}</span></div>`).join('');
}

function renderSpendingPie(containerId, group = '') {
  const breakdown = dateRangeBreakdown();
  const container = $(`#${containerId}`);
  const rows = Object.entries(breakdown || {})
    .filter(([id, amount]) => amount > 0 && (!group || getItem(id)?.group === group))
    .sort((a, b) => b[1] - a[1]);
  if (!rows.length) {
    const typeLabel = group === 'operating' ? 'operativos' : group === 'fixed' ? 'fijos' : '';
    const emptyMessage = typeLabel
      ? `No hay gastos ${typeLabel} registrados en este rango.`
      : 'La gráfica aparecerá cuando registres gastos en este rango.';
    container.innerHTML = `<div class="empty-row spending-pie-empty">${emptyMessage}</div>`;
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

function movementRowHtml(row, includePayment, includeSpender = includePayment, allowEdit = false) {
  const item = getItem(row.category);
  const actions = row.source === 'manual'
    ? `<div class="row-actions">${allowEdit ? `<button class="edit-button" type="button" data-edit-id="${row.id}">Editar</button>` : ''}<button class="delete-button" type="button" data-delete-id="${row.id}" aria-label="Eliminar gasto">×</button></div>`
    : '';
  return `<tr><td class="date-cell">${shortDate(row.date)}</td><td><span class="tag">${escapeHtml(item?.name || row.category)}</span></td>${includeSpender ? `<td>${escapeHtml(row.spender || '—')}</td>` : ''}${includePayment ? `<td>${escapeHtml(row.payment || '—')}</td>` : ''}<td class="note-cell">${escapeHtml(row.note || 'Sin nota')}</td><td class="align-right amount-cell">${money(row.amount)}</td><td>${actions}</td></tr>`;
}

function resetExpenseForm() {
  $('#expenseForm').reset();
  $('#expenseDate').value = localToday();
  renderExpenseCategories();
}

function closeExpenseEdit() {
  editingExpenseId = '';
  if ($('#editExpenseDialog').open) $('#editExpenseDialog').close();
}

function syncEditExpenseTypeFromCategory() {
  const selectedItem = getItem($('#editExpenseCategory').value);
  $('#editExpenseType').value = selectedItem?.group === 'fixed' ? 'Fijo' : 'Operativo';
}

function setEditExpenseSpender(spender) {
  const selectedSpender = String(spender || '').trim();
  const choices = selectedSpender && !expenseSpenders.includes(selectedSpender) ? [...expenseSpenders, selectedSpender] : expenseSpenders;
  $('#editExpenseSpender').innerHTML = `<option value="" disabled>Selecciona una persona</option>${choices.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}${!expenseSpenders.includes(name) ? ' (registro anterior)' : ''}</option>`).join('')}`;
  $('#editExpenseSpender').value = selectedSpender;
}

function startExpenseEdit(id) {
  const entry = state.manualEntries.find(candidate => candidate.id === id);
  if (!entry) return showToast('Este gasto ya no está disponible para editar.');
  editingExpenseId = id;
  const period = periods.find(candidate => candidate.id === entry.periodId) || getPeriod();
  $('#editExpenseSubtitle').textContent = `${shortDate(entry.date)} · ${getItem(entry.category)?.name || entry.category}`;
  $('#editExpenseWeek').innerHTML = period.weeks.map((week, index) => `<option value="${index}">${escapeHtml(week.label)}</option>`).join('');
  $('#editExpenseCategory').innerHTML = expenseItems.map(item => `<option value="${item.id}">${escapeHtml(item.name)}</option>`).join('');
  $('#editExpenseDate').value = entry.date;
  $('#editExpenseWeek').value = Number(entry.weekIndex) || 0;
  $('#editExpenseCategory').value = entry.category;
  syncEditExpenseTypeFromCategory();
  setEditExpenseSpender(entry.spender);
  $('#editExpensePayment').value = entry.payment || 'Efectivo';
  $('#editExpenseAmount').value = Number(entry.amount);
  $('#editExpenseNote').value = entry.note === 'Sin nota' ? '' : entry.note || '';
  $('#editExpenseDialog').showModal();
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
  const options = expenseItems.map(item => `<option value="${item.id}">${escapeHtml(item.name)}</option>`).join('');
  $('#expenseCategory').innerHTML = options;
  $('#bulkExpenseCategory').innerHTML = options;
  if (expenseItems.some(item => item.id === selectedCategory)) $('#expenseCategory').value = selectedCategory;
  if (expenseItems.some(item => item.id === selectedBulkCategory)) $('#bulkExpenseCategory').value = selectedBulkCategory;
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
  $('#newMovementRows').innerHTML = entries.length ? entries.map(row => movementRowHtml(row, true, true, true)).join('') : `<tr><td colspan="7" class="empty-row">Los gastos que guardes para esta semana aparecerán aquí.</td></tr>`;
  if (!$('#expenseDate').value) $('#expenseDate').value = localToday();
  if (!$('#bulkExpenseDate').value) $('#bulkExpenseDate').value = localToday();
  renderBulkDrafts();
  setCaptureMode(activeCaptureMode);
}

function renderExpenses() {
  const weeklyBreakdown = snapshotBreakdown() || {};
  const todayBreakdown = snapshotBreakdown(localToday()) || {};
  const monthlyBreakdown = monthlyBreakdownForSelection();
  const operating = expenseItems.filter(item => item.group === 'operating');
  const fixed = expenseItems.filter(item => item.group === 'fixed');
  $('#operatingWeekly').textContent = money(breakdownTotal(operating, weeklyBreakdown));
  $('#fixedWeekly').textContent = money(breakdownTotal(fixed, weeklyBreakdown));
  $('#expenseGrandMonthly').textContent = money(monthlySpentForSelection());
  $('#operatingExpenseRows').innerHTML = operating.map(item => expenseRowHtml(item, todayBreakdown, weeklyBreakdown, monthlyBreakdown)).join('');
  $('#fixedExpenseRows').innerHTML = fixed.map(item => expenseRowHtml(item, todayBreakdown, weeklyBreakdown, monthlyBreakdown)).join('');
  $('#operatingExpenseTotal').innerHTML = totalRowHtml(breakdownTotal(operating, todayBreakdown), breakdownTotal(operating, weeklyBreakdown), breakdownTotal(operating, monthlyBreakdown));
  $('#fixedExpenseTotal').innerHTML = totalRowHtml(breakdownTotal(fixed, todayBreakdown), breakdownTotal(fixed, weeklyBreakdown), breakdownTotal(fixed, monthlyBreakdown));
}

function expenseRowHtml(item, todayBreakdown, weeklyBreakdown, monthlyBreakdown) {
  return `<tr><td><strong>${escapeHtml(item.name)}</strong></td><td class="align-right daily-amount">${money(todayBreakdown[item.id])}</td><td class="align-right amount-cell">${money(weeklyBreakdown[item.id])}</td><td class="align-right">${money(monthlyBreakdown[item.id])}</td></tr>`;
}

function totalRowHtml(today, weekly, monthly) {
  return `<tr class="total-row"><td>Total</td><td class="align-right">${money(today)}</td><td class="align-right">${money(weekly)}</td><td class="align-right">${money(monthly)}</td></tr>`;
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
  $('#loginForm').addEventListener('submit', async event => {
    event.preventDefault();
    const submitButton = event.submitter;
    $('#loginError').textContent = '';
    if (submitButton) submitButton.disabled = true;
    try {
      const payload = await authRequest({
        method: 'POST',
        body: { username: $('#loginUsername').value.trim(), password: $('#loginPassword').value },
      });
      await startSession(payload.user);
    } catch (error) {
      $('#loginError').textContent = error.message;
      $('#loginPassword').select();
    } finally {
      if (submitButton) submitButton.disabled = false;
    }
  });
  $('#logoutButton').addEventListener('click', async () => {
    try { await authRequest({ method: 'DELETE' }); } catch (error) { console.warn('No se pudo cerrar la sesión en el servidor:', error); }
    showLogin('Sesión cerrada correctamente.');
  });
  document.addEventListener('click', async event => {
    const viewButton = event.target.closest('[data-view]');
    if (viewButton) setView(viewButton.dataset.view);
    const editButton = event.target.closest('[data-edit-id]');
    if (editButton) startExpenseEdit(editButton.dataset.editId);
    const deleteButton = event.target.closest('[data-delete-id]');
    if (deleteButton) {
      if (!databaseReady) return showToast('Espera a que la base de datos esté conectada.');
      deleteButton.disabled = true;
      try {
        await apiRequest({ method: 'DELETE', body: undefined, id: deleteButton.dataset.deleteId });
        state.manualEntries = state.manualEntries.filter(entry => entry.id !== deleteButton.dataset.deleteId);
        if (editingExpenseId === deleteButton.dataset.deleteId) closeExpenseEdit();
        renderAll();
        showToast('Gasto eliminado de la base de datos.');
      } catch (error) {
        deleteButton.disabled = false;
        showToast(error.message);
      }
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
  $$('[data-date-filter]').forEach(button => button.addEventListener('click', () => {
    activeDateFilter = button.dataset.dateFilter;
    applyDateFilter();
    renderDashboard();
  }));
  $('#dayFilterDate').addEventListener('change', event => {
    selectedFilterDay = event.target.value || localToday();
    applyDateFilter();
    renderDashboard();
  });
  $('#weekFilterWeek').addEventListener('change', event => {
    selectedFilterWeek = event.target.value || dateToIsoWeek(localToday());
    applyDateFilter();
    renderDashboard();
  });
  $('#monthFilterMonth').addEventListener('change', event => {
    selectedFilterMonth = event.target.value || localToday().slice(0, 7);
    applyDateFilter();
    renderDashboard();
  });
  $('#captureWeek').addEventListener('change', event => { selectedWeekIndex = Number(event.target.value); renderSelectors(); renderCapture(); });
  $('#expenseCategory').addEventListener('change', syncExpenseTypeFromCategory);
  $('#expenseCancelButton').addEventListener('click', () => setView('dashboard'));
  $('#editExpenseCategory').addEventListener('change', syncEditExpenseTypeFromCategory);
  $('#closeEditExpense').addEventListener('click', closeExpenseEdit);
  $('#cancelEditExpense').addEventListener('click', closeExpenseEdit);
  $('#editExpenseDialog').addEventListener('click', event => { if (event.target === event.currentTarget) closeExpenseEdit(); });
  $('#editExpenseDialog').addEventListener('close', () => { editingExpenseId = ''; });
  $('#bulkCaptureWeek').addEventListener('change', event => { selectedWeekIndex = Number(event.target.value); renderSelectors(); renderCapture(); });
  $('#bulkExpenseCategory').addEventListener('change', syncBulkExpenseTypeFromCategory);

  $('#expenseForm').addEventListener('submit', async event => {
    event.preventDefault();
    if (!databaseReady) return showToast('Espera a que la base de datos esté conectada.');
    const amount = Number($('#expenseAmount').value);
    if (!amount || amount <= 0) return showToast('Escribe un monto mayor a cero.');
    const category = $('#expenseCategory').value;
    const submitButton = event.submitter;
    const newEntry = {
      id: newExpenseId(),
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
    };

    if (submitButton) submitButton.disabled = true;
    try {
      const payload = await apiRequest({ method: 'POST', body: { entry: newEntry } });
      if (Array.isArray(payload.entries)) state.manualEntries = payload.entries;
      resetExpenseForm();
      renderAll();
      showToast(currentUser?.role === 'admin' ? 'Gasto guardado en Neon y totales actualizados.' : 'Gasto guardado correctamente.');
    } catch (error) {
      showToast(error.message);
    } finally {
      if (submitButton) submitButton.disabled = false;
    }
  });

  $('#editExpenseForm').addEventListener('submit', async event => {
    event.preventDefault();
    if (!databaseReady) return showToast('Espera a que la base de datos esté conectada.');
    const currentEntry = state.manualEntries.find(entry => entry.id === editingExpenseId);
    if (!currentEntry) return closeExpenseEdit();
    const amount = Number($('#editExpenseAmount').value);
    if (!amount || amount <= 0) return showToast('Escribe un monto mayor a cero.');
    const category = $('#editExpenseCategory').value;
    const submitButton = event.submitter;
    const updatedEntry = {
      ...currentEntry,
      date: $('#editExpenseDate').value,
      weekIndex: Number($('#editExpenseWeek').value),
      category,
      amount,
      note: $('#editExpenseNote').value.trim() || 'Sin nota',
      payment: $('#editExpensePayment').value,
      spender: $('#editExpenseSpender').value.trim(),
      expenseType: getItem(category)?.group === 'fixed' ? 'Fijo' : 'Operativo',
      source: 'manual'
    };
    if (submitButton) submitButton.disabled = true;
    try {
      const payload = await apiRequest({ method: 'POST', body: { entry: updatedEntry } });
      if (Array.isArray(payload.entries)) state.manualEntries = payload.entries;
      closeExpenseEdit();
      renderAll();
      showToast('Gasto actualizado en Neon.');
    } catch (error) {
      showToast(error.message);
    } finally {
      if (submitButton) submitButton.disabled = false;
    }
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

  $('#saveBulkExpenses').addEventListener('click', async event => {
    if (!bulkDraftEntries.length) return;
    if (!databaseReady) return showToast('Espera a que la base de datos esté conectada.');
    const entriesToSave = bulkDraftEntries.map(entry => ({ ...entry, id: newExpenseId(), source: 'manual' }));
    event.currentTarget.disabled = true;
    try {
      const payload = await apiRequest({ method: 'POST', body: { entries: entriesToSave } });
      if (Array.isArray(payload.entries)) state.manualEntries = payload.entries;
      bulkDraftEntries = [];
      renderAll();
      showToast(`${entriesToSave.length} ${entriesToSave.length === 1 ? 'gasto guardado' : 'gastos guardados'} en Neon.`);
    } catch (error) {
      showToast(error.message);
    } finally {
      event.currentTarget.disabled = bulkDraftEntries.length === 0;
    }
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

  $('#refreshData').addEventListener('click', () => syncExpenses({ migrateLegacy: true }));
}

function renderAll() {
  renderSelectors();
  if (currentUser?.role === 'admin') renderDashboard();
  if (activeView === 'capture') renderCapture();
  if (activeView === 'expenses' && currentUser?.role === 'admin') renderExpenses();
}

bindEvents();
$('#expenseDate').value = localToday();
$('#bulkExpenseDate').value = localToday();
initializeSession();
window.addEventListener('focus', () => { if (currentUser?.role === 'admin') syncExpenses({ migrateLegacy: false, silent: true }); });
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && currentUser?.role === 'admin') syncExpenses({ migrateLegacy: false, silent: true });
});
window.setInterval(() => {
  if (document.visibilityState === 'visible' && currentUser?.role === 'admin') syncExpenses({ migrateLegacy: false, silent: true });
}, 30000);
