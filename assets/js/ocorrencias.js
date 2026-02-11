/* ============================================================
   Dashboard de Chamados - Script completo (HTML único)
   - Tema Claro/Escuro (salva em localStorage)
   - Aging funcionando (conta + filtro toggle por bloco)
   - Comparar Excel: modal ticket-a-ticket com tudo que mudou
   ============================================================ */

let allTickets = [];
let filteredTickets = [];
let activeKPI = 'all';

let statusSelected = new Set();
let lineChartInstance = null;

const SLA_PROXY_DAYS = 7; // proxy (ajustável)
let activeAgingBucket = null; // '0-3' | '4-7' | '8-15' | '15+' | null

// Categoria (Painel)
let activeCategory = ''; // valor do select (categoria original do Excel)

// Diretoria
let diretoriaView = 'all'; // 'all' | 'problema' | 'adeq_duvida'

// Comparação Excel
let compareOldTickets = [];
let compareNewTickets = [];
let compareDiffRows = []; // {numero, tipo, campo, antes, depois}
let compareFilterMode = 'all';

// Mapas para abrir modal do comparativo (snapshot por ticket)
let compareOldMap = new Map(); // numero -> ticket velho
let compareNewMap = new Map(); // numero -> ticket novo

function $(id) { return document.getElementById(id); }

function triggerFile() { $('fileInput').click(); }

function removeExcel(openAfter) {
  try { localStorage.removeItem('tickets_cache_v2'); } catch (_) { }
  allTickets = [];
  filteredTickets = [];
  $('fileLabel').textContent = 'Nenhum arquivo carregado';
  $('mobileFile').textContent = 'nenhum arquivo carregado';
  $('cacheFile').textContent = 'nenhum arquivo carregado';
  $('cacheCount').textContent = '0';
  try { clearAllFilters(false); } catch (_) { }
  const tbody = $('ticketTableBody');
  if (tbody) tbody.innerHTML = '<tr><td colspan="12" style="text-align:center; padding:26px; color:var(--muted);">Carregue um arquivo Excel para iniciar.</td></tr>';
  try { updateEverything(); } catch (_) { }
  setupTopTableScroll();
  if (openAfter) setTimeout(() => triggerFile(), 0);
}

function setupTopTableScroll() {
  const top = $('tableXScrollTop');
  const inner = $('tableXScrollTopInner');
  const wrap = $('tableWrap');
  const table = $('ticketTable');
  if (!top || !inner || !wrap || !table) return;

  if (!top._bound) {
    top.addEventListener('scroll', () => { wrap.scrollLeft = top.scrollLeft; });
    wrap.addEventListener('scroll', () => { top.scrollLeft = wrap.scrollLeft; });
    top._bound = true;
  }
  inner.style.width = table.scrollWidth + 'px';
  top.scrollLeft = wrap.scrollLeft;
}

function openSidebar() {
  $('sidebar').classList.add('open');
  $('sidebarOverlay').classList.add('open');
}
function closeSidebar() {
  $('sidebar').classList.remove('open');
  $('sidebarOverlay').classList.remove('open');
}

function showPage(page) {
  $('tabDashboard').classList.toggle('active', page === 'dashboard');
  $('tabDiretoria').classList.toggle('active', page === 'diretoria');
  $('tabCompare').classList.toggle('active', page === 'compare');
  $('tabTicket').classList.toggle('active', page === 'ticket');

  $('pageDashboard').style.display = page === 'dashboard' ? 'block' : 'none';
  $('pageDiretoria').style.display = page === 'diretoria' ? 'block' : 'none';
  $('pageCompare').style.display = page === 'compare' ? 'block' : 'none';
  $('pageTicket').style.display = page === 'ticket' ? 'block' : 'none';

  // Ao entrar em páginas novas, atualiza a UI específica
  if (page === 'diretoria') updateDiretoriaPage();
  if (page === 'compare') { setupCompareTopTableScroll(); renderCompareTable(); }

  closeSidebar();
}

/* ---------------------------
   Tema Claro/Escuro
--------------------------- */
function setTheme(theme) {
  document.body.setAttribute('data-theme', theme);
  try { localStorage.setItem('dash_theme', theme); } catch (_) { }
  const icon = (theme === 'dark') ? '☀️ Tema' : '🌙 Tema';
  if ($('themeBtn')) $('themeBtn').textContent = icon;
  if ($('themeBtnMobile')) $('themeBtnMobile').textContent = icon;
}
function toggleTheme() {
  const cur = document.body.getAttribute('data-theme') || 'light';
  setTheme(cur === 'dark' ? 'light' : 'dark');
}
function loadTheme() {
  try {
    const saved = localStorage.getItem('dash_theme');
    if (saved === 'dark' || saved === 'light') return saved;
  } catch (_) { }
  return 'light';
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    alert('Biblioteca XLSX não carregou. Verifique conexão com internet ou salve a biblioteca localmente.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'dd/mm/yyyy' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: 'dd/mm/yyyy' });

      allTickets = rows.map(normalizeTicket).filter(t => String(t.numero || '').trim() !== '');
      saveCache(file.name, allTickets);
      setFileLabel(file.name, allTickets.length);
      buildFilterOptions();
      buildStatusMultiSelect();
      clearAllFilters(false);
      updateEverything();
    } catch (err) {
      console.error(err);
      alert('Falha ao ler o arquivo. Verifique se é um Excel válido e se a primeira planilha contém os dados.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function setFileLabel(name, count) {
  $('fileLabel').textContent = name;
  $('mobileFile').textContent = name;
  $('cacheFile').textContent = name;
  $('cacheCount').textContent = String(count || 0);
}

function saveCache(fileName, data) {
  try {
    localStorage.setItem('tickets_cache_v2', JSON.stringify({ fileName, data, ts: Date.now() }));
  } catch (_) { }
}

function loadCache() {
  try {
    const raw = localStorage.getItem('tickets_cache_v2');
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj?.data || !Array.isArray(obj.data)) return null;
    return obj;
  } catch (_) {
    return null;
  }
}

/* ---------------------------
   Normalização de colunas
--------------------------- */
function normKey(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function findValue(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const cNorm = normKey(c);
    const found = keys.find(k => normKey(k).includes(cNorm));
    if (found !== undefined && row[found] !== undefined && row[found] !== null && String(row[found]).trim() !== '') {
      return String(row[found]).trim();
    }
  }
  return '';
}

function normalizeTicket(row) {
  const numero = findValue(row, ['numero', 'número', 'ticket', 'id', 'codigo']);
  let numeroFix = numero;
  if (/^\d+\.0+$/.test(numeroFix)) numeroFix = numeroFix.split('.')[0];

  const abertoEm = findValue(row, ['aberto em', 'abertoem', 'data abertura', 'data de abertura', 'criado em', 'criado']);
  const departamento = findValue(row, ['departamento', 'depto', 'setor', 'area']);
  const solicitante = findValue(row, ['usuario solicitante', 'usuário solicitante', 'solicitante', 'requester']);
  const servico = findValue(row, ['servico', 'serviço', 'service', 'modulo', 'módulo']);
  const assunto = findValue(row, ['assunto', 'titulo', 'título', 'subject', 'descricao', 'descrição']);
  const responsavel = findValue(row, ['responsavel', 'responsável', 'atribuido', 'atribuído', 'assigned']);
  const categoria = findValue(row, ['categoria', 'category', 'tipo']);
  const ultimaAcao = findValue(row, ['data da última ação', 'data da ultima acao', 'última ação', 'ultima acao', 'ultimaacao', 'atualizado em', 'updated']);
  const status = findValue(row, ['status', 'situacao', 'situação', 'estado']);
  const cliente = findValue(row, ['cliente (pessoa)', 'cliente pessoa', 'cliente']);
  let priorizado = findValue(row, ['ticket priorizado', 'priorizado', 'prioritario', 'prioritário', 'priority']);

  priorizado = /sim/i.test(priorizado) ? 'Sim' : (/nao|não/i.test(priorizado) ? 'Nao' : (priorizado ? priorizado : 'Nao'));

  return {
    numero: numeroFix,
    abertoEm,
    departamento,
    solicitante,
    cliente,
    servico,
    assunto,
    responsavel,
    categoria,
    ultimaAcao,
    status,
    priorizado,
    _raw: row
  };
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  const s = String(v).trim();
  if (!s) return null;

  // Excel serial
  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = parseFloat(s);
    if (serial > 1 && serial < 100000) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const ms = epoch.getTime() + serial * 24 * 60 * 60 * 1000;
      const d = new Date(ms);
      if (!isNaN(d.getTime()) && d.getFullYear() > 1990 && d.getFullYear() < 2100) return d;
    }
  }

  // BR dd/mm/yyyy hh:mm
  const br = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (br) {
    let [, dd, mm, yy, hh = '0', mi = '0', ss = '0'] = br;
    yy = yy.length === 2 ? '20' + yy : yy;
    const d = new Date(parseInt(yy), parseInt(mm) - 1, parseInt(dd), parseInt(hh), parseInt(mi), parseInt(ss));
    return isNaN(d.getTime()) ? null : d;
  }

  // ISO-ish
  const iso = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (iso) {
    const [, yy, mm, dd, hh = '0', mi = '0', ss = '0'] = iso;
    const d = new Date(parseInt(yy), parseInt(mm) - 1, parseInt(dd), parseInt(hh), parseInt(mi), parseInt(ss));
    return isNaN(d.getTime()) ? null : d;
  }

  const nd = new Date(s);
  return isNaN(nd.getTime()) ? null : nd;
}

function fmtDate(v) {
  const d = parseDate(v);
  if (!d) return v ? String(v) : '-';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  return `${dd}/${mm}/${yy} ${hh}:${mi}`;
}

function daysOpen(ticket) {
  const d = parseDate(ticket.abertoEm);
  if (!d) return null;
  const today = new Date();
  const diff = Math.floor((today - d) / (1000 * 60 * 60 * 24));
  return isNaN(diff) ? null : diff;
}

function isClosedStatus(status) {
  const s = String(status || '').toLowerCase();
  return s.includes('fechado') || s.includes('resolvido') || s.includes('cancelado');
}

function isAwaitingStatus(status) {
  const s = String(status || '').toLowerCase();
  return s.includes('aguardando');
}

/* ---------------------------
   Categoria (classificação)
--------------------------- */
function normStr(v) { return String(v || '').trim().toLowerCase(); }

// Classificação simples para relatório gerencial.
function classifyCategory(ticket) {
  const c = normStr(ticket?.categoria);
  if (!c) return 'adeq_duvida';
  if (c.includes('problema')) return 'problema';
  if (c.includes('duvida') || c.includes('dúvida')) return 'adeq_duvida';
  if (c.includes('adequ')) return 'adeq_duvida';
  return 'adeq_duvida';
}

function quickCategory(mode) {
  activeKPI = 'all';
  document.querySelectorAll('.kpi').forEach(k => k.classList.remove('active'));
  $('kpiTotal').classList.add('active');

  if (mode === 'all') {
    activeCategoryMode = null;
  } else {
    activeCategoryMode = (activeCategoryMode === mode) ? null : mode;
  }
  applyFilters();
}

let activeCategoryMode = null; // 'problema' | 'adeq_duvida' | null

function updateCategoryKpis() {
  const pF = (filteredTickets || []).filter(t => !isClosedStatus(t.status) && classifyCategory(t) === 'problema').length;
  const aF = (filteredTickets || []).filter(t => !isClosedStatus(t.status) && classifyCategory(t) !== 'problema').length;

  $('kCatProblema').textContent = pF;
  $('kCatAdeqDuvida').textContent = aF;

  $('kpiCatProblema')?.classList.toggle('active', activeCategoryMode === 'problema');
  $('kpiCatAdeqDuvida')?.classList.toggle('active', activeCategoryMode === 'adeq_duvida');
  $('kpiCatAll')?.classList.toggle('active', activeCategoryMode === null);
}

/* ---------------------------
   Aging (contagem + filtro)
--------------------------- */
function bucketMatch(days, bucket) {
  if (days === null || days === undefined) return false;
  if (bucket === '0-3') return days >= 0 && days <= 3;
  if (bucket === '4-7') return days >= 4 && days <= 7;
  if (bucket === '8-15') return days >= 8 && days <= 15;
  if (bucket === '15+') return days >= 16;
  return false;
}

function setAging(bucket) {
  activeAgingBucket = (activeAgingBucket === bucket) ? null : bucket;
  applyFilters();
}

function updateAgingCards() {
  const base = (filteredTickets || []).filter(t => !isClosedStatus(t.status));
  let c03 = 0, c47 = 0, c815 = 0, c15 = 0;

  base.forEach(t => {
    const d = daysOpen(t);
    if (bucketMatch(d, '0-3')) c03++;
    else if (bucketMatch(d, '4-7')) c47++;
    else if (bucketMatch(d, '8-15')) c815++;
    else if (bucketMatch(d, '15+')) c15++;
  });

  $('aging03').textContent = c03;
  $('aging47').textContent = c47;
  $('aging815').textContent = c815;
  $('aging15').textContent = c15;

  const cards = document.querySelectorAll('#agingCards .agingCard');
  cards.forEach(card => {
    const b = card.getAttribute('data-bucket');
    card.classList.toggle('active', activeAgingBucket === b);
  });
}

/* ---------------------------
   Filtros e opções
--------------------------- */
function buildFilterOptions() {
  const yearSet = new Set();
  const depSet = new Set();
  const clientSet = new Set();
  const catSet = new Set();

  allTickets.forEach(t => {
    const d = parseDate(t.abertoEm);
    if (d) yearSet.add(String(d.getFullYear()));
    if (t.departamento) depSet.add(t.departamento);
    if (t.cliente) clientSet.add(t.cliente);
    if (t.categoria) catSet.add(t.categoria);
  });

  const yearSel = $('yearFilter');
  yearSel.innerHTML = `<option value="">Todos</option>` + [...yearSet].sort((a, b) => b.localeCompare(a)).map(y => `<option value="${y}">${y}</option>`).join('');

  const monthSel = $('monthFilter');
  const monthNames = ['01 - Jan', '02 - Fev', '03 - Mar', '04 - Abr', '05 - Mai', '06 - Jun', '07 - Jul', '08 - Ago', '09 - Set', '10 - Out', '11 - Nov', '12 - Dez'];
  monthSel.innerHTML = `<option value="">Todos</option>` + monthNames.map(m => `<option value="${m.slice(0, 2)}">${m}</option>`).join('');

  const depSel = $('departmentFilter');
  depSel.innerHTML = `<option value="">Todos</option>` + [...depSet].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(d => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join('');

  const cliSel = $('clientFilter');
  cliSel.innerHTML = `<option value="">Todos</option>` + [...clientSet].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');

  const catSel = $('categoryFilter');
  if (catSel) {
    catSel.innerHTML = `<option value="">Todas</option>` + [...catSet].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  }
}

function buildStatusMultiSelect() {
  const stSet = new Set();
  allTickets.forEach(t => { if (t.status) stSet.add(t.status); });
  const statuses = [...stSet].sort((a, b) => a.localeCompare(b, 'pt-BR'));

  const menu = $('statusMSMenu');
  menu.innerHTML = '';

  const frag = document.createDocumentFragment();

  statuses.forEach(st => {
    const id = 'st_' + Math.random().toString(36).slice(2);
    const row = document.createElement('div');
    row.className = 'ms-item';
    row.innerHTML = `<input type="checkbox" id="${id}"><span>${escapeHtml(st)}</span>`;
    row.onclick = (e) => {
      if (e.target.tagName !== 'INPUT') {
        row.querySelector('input').checked = !row.querySelector('input').checked;
      }
      const checked = row.querySelector('input').checked;
      if (checked) statusSelected.add(st); else statusSelected.delete(st);
      updateStatusMSLabel();
      applyFilters();
    };
    frag.appendChild(row);
  });

  const actions = document.createElement('div');
  actions.className = 'ms-actions';
  actions.innerHTML = `<button type="button" onclick="selectAllStatus(true); event.stopPropagation();">Marcar todos</button>
                       <button type="button" onclick="selectAllStatus(false); event.stopPropagation();">Limpar</button>`;
  frag.appendChild(actions);

  menu.appendChild(frag);
  updateStatusMSLabel();
}

function selectAllStatus(select) {
  const items = $('statusMSMenu').querySelectorAll('.ms-item');
  items.forEach(it => {
    const st = it.querySelector('span').textContent;
    it.querySelector('input').checked = !!select;
    if (select) statusSelected.add(st); else statusSelected.delete(st);
  });
  updateStatusMSLabel();
  applyFilters();
}

function toggleMS(id) {
  const el = $(id);
  el.classList.toggle('open');
  ['statusMS'].forEach(other => {
    if (other !== id) $(other).classList.remove('open');
  });
}

document.addEventListener('click', (e) => {
  const ms = $('statusMS');
  if (ms && !ms.contains(e.target)) {
    ms.classList.remove('open');
  }
});

function updateStatusMSLabel() {
  const val = $('statusMSValue');
  if (statusSelected.size === 0) {
    val.textContent = 'Todos';
  } else if (statusSelected.size === 1) {
    val.textContent = [...statusSelected][0];
  } else {
    val.textContent = `${statusSelected.size} selecionados`;
  }
}

/* Debounce for search */
let _debounce = null;
function applyFiltersDebounced() {
  clearTimeout(_debounce);
  _debounce = setTimeout(() => applyFilters(), 180);
}

function clearAllFilters(resetKPI = true) {
  $('searchInput').value = '';
  $('departmentFilter').value = '';
  $('clientFilter').value = '';
  if ($('categoryFilter')) $('categoryFilter').value = '';
  $('priorityFilter').value = '';
  $('yearFilter').value = '';
  $('monthFilter').value = '';
  $('onlyOpen').value = '';

  activeAgingBucket = null;
  activeCategoryMode = null;

  statusSelected.clear();
  const items = $('statusMSMenu')?.querySelectorAll?.('.ms-item input') || [];
  items.forEach(i => i.checked = false);
  updateStatusMSLabel();

  if (resetKPI) activeKPI = 'all';
  document.querySelectorAll('.kpi').forEach(k => k.classList.remove('active'));
  $('kpiTotal').classList.add('active');

  applyFilters();
}

function setKPIFilter(kpi) {
  activeKPI = kpi;
  document.querySelectorAll('.kpi').forEach(k => k.classList.remove('active'));
  const map = {
    all: 'kpiTotal',
    open: 'kpiOpen',
    awaiting: 'kpiAwait',
    prioritized: 'kpiPrio',
    overdue: 'kpiSLA',
    resolved: 'kpiResolved'
  };
  const id = map[kpi] || 'kpiTotal';
  $(id).classList.add('active');
  applyFilters();
}

/* ---------------------------
   Apply + Update UI
--------------------------- */
function applyFilters() {
  if (!Array.isArray(allTickets) || allTickets.length === 0) {
    filteredTickets = [];
    renderTable();
    updateEverything();
    return;
  }

  const q = $('searchInput').value.trim().toLowerCase();
  const dep = $('departmentFilter').value;
  const cli = $('clientFilter').value;
  const cat = $('categoryFilter') ? $('categoryFilter').value : '';
  const prio = $('priorityFilter').value;
  const year = $('yearFilter').value;
  const month = $('monthFilter').value;
  const onlyOpen = $('onlyOpen').value === '1';

  filteredTickets = allTickets.filter(t => {
    if (activeKPI === 'open' && isClosedStatus(t.status)) return false;
    if (activeKPI === 'resolved' && !isClosedStatus(t.status)) return false;
    if (activeKPI === 'awaiting' && (isClosedStatus(t.status) || !isAwaitingStatus(t.status))) return false;
    if (activeKPI === 'prioritized' && (isClosedStatus(t.status) || String(t.priorizado).toLowerCase() !== 'sim')) return false;
    if (activeKPI === 'overdue') {
      if (isClosedStatus(t.status)) return false;
      const d = daysOpen(t);
      if (d === null || d <= SLA_PROXY_DAYS) return false;
    }

    if (onlyOpen && isClosedStatus(t.status)) return false;

    if (statusSelected.size > 0 && !statusSelected.has(t.status)) return false;

    if (dep && t.departamento !== dep) return false;
    if (cli && t.cliente !== cli) return false;
    if (cat && t.categoria !== cat) return false;

    if (activeCategoryMode) {
      const cls = classifyCategory(t);
      if (activeCategoryMode === 'problema' && cls !== 'problema') return false;
      if (activeCategoryMode === 'adeq_duvida' && cls === 'problema') return false;
    }

    if (prio) {
      const p = String(t.priorizado || '').toLowerCase() === 'sim' ? 'Sim' : 'Nao';
      if (p !== prio) return false;
    }

    if (year || month) {
      const d = parseDate(t.abertoEm);
      if (!d) return false;
      if (year && String(d.getFullYear()) !== year) return false;
      if (month && String(d.getMonth() + 1).padStart(2, '0') !== month) return false;
    }

    if (activeAgingBucket) {
      if (isClosedStatus(t.status)) return false;
      const d = daysOpen(t);
      if (!bucketMatch(d, activeAgingBucket)) return false;
    }

    if (q) {
      const hay = [
        t.numero, t.assunto, t.solicitante, t.cliente, t.servico,
        t.status, t.departamento, t.responsavel, t.categoria
      ].map(x => String(x || '').toLowerCase());
      if (!hay.some(x => x.includes(q))) return false;
    }

    return true;
  });

  updateEverything();
}

/* ---------------------------
   Update Everything
--------------------------- */
function updateEverything() {
  updateKpis();
  updateCategoryKpis();
  updateSidebar();
  updateChips();
  renderTable();
  updateCharts();
  updateLineChart();
  updateAgingCards();
  updateDiretoriaPage();
}

function updateKpis() {
  const total = allTickets.length;
  const open = allTickets.filter(t => !isClosedStatus(t.status)).length;
  const resolved = total - open;
  const resolvedPct = total ? Math.round((resolved / total) * 100) : 0;

  const awaiting = allTickets.filter(t => !isClosedStatus(t.status) && isAwaitingStatus(t.status)).length;
  const prioOpen = allTickets.filter(t => !isClosedStatus(t.status) && String(t.priorizado).toLowerCase() === 'sim').length;

  const overdue = allTickets.filter(t => {
    if (isClosedStatus(t.status)) return false;
    const d = daysOpen(t);
    return d !== null && d > SLA_PROXY_DAYS;
  }).length;

  $('kTotal').textContent = total;
  $('kOpen').textContent = open;
  $('kResolved').textContent = resolved;
  $('kResolvedPct').textContent = resolvedPct + '%';
  $('kAwait').textContent = awaiting;
  $('kPrio').textContent = prioOpen;
  $('kOverdue').textContent = overdue;

  $('sideBase').textContent = String(filteredTickets.length || 0);

  const openF = filteredTickets.filter(t => !isClosedStatus(t.status)).length;
  const awaitF = filteredTickets.filter(t => !isClosedStatus(t.status) && isAwaitingStatus(t.status)).length;
  const prioF = filteredTickets.filter(t => !isClosedStatus(t.status) && String(t.priorizado).toLowerCase() === 'sim').length;
  const overdueF = filteredTickets.filter(t => {
    if (isClosedStatus(t.status)) return false;
    const d = daysOpen(t); return d !== null && d > SLA_PROXY_DAYS;
  }).length;

  $('sideOpen').textContent = openF;
  $('sideAwait').textContent = awaitF;
  $('sidePrio').textContent = prioF;
  $('sideOverdue').textContent = overdueF;

  $('kTotalHint').textContent = $('fileLabel').textContent.includes('Nenhum') ? 'Carregue o Excel' : `Arquivo: ${$('fileLabel').textContent}`;
}

function updateSidebar() {
  const prio = filteredTickets
    .filter(t => !isClosedStatus(t.status) && String(t.priorizado).toLowerCase() === 'sim')
    .sort((a, b) => (daysOpen(b) || 0) - (daysOpen(a) || 0))
    .slice(0, 12);

  const await = filteredTickets
    .filter(t => !isClosedStatus(t.status) && isAwaitingStatus(t.status))
    .sort((a, b) => (daysOpen(b) || 0) - (daysOpen(a) || 0))
    .slice(0, 12);

  $('sidePrioListCount').textContent = prio.length;
  $('sideAwaitListCount').textContent = await.length;

  $('prioList').innerHTML = prio.length ? prio.map(t => miniTicketHtml(t, 'prio')).join('') : `<li class="small" style="color:var(--muted); padding:4px 2px;">Nenhum na base filtrada.</li>`;
  $('awaitList').innerHTML = await.length ? await.map(t => miniTicketHtml(t, 'wait')).join('') : `<li class="small" style="color:var(--muted); padding:4px 2px;">Nenhum na base filtrada.</li>`;

  $('prioList').querySelectorAll?.('.ticket-mini')?.forEach(el => {
    el.addEventListener('click', () => showTicketModal(el.getAttribute('data-num')));
  });
  $('awaitList').querySelectorAll?.('.ticket-mini')?.forEach(el => {
    el.addEventListener('click', () => showTicketModal(el.getAttribute('data-num')));
  });
}

function miniTicketHtml(t, type) {
  const d = daysOpen(t);
  const slaBad = (d !== null && d > SLA_PROXY_DAYS);
  const badgeClass = type === 'prio' ? 'prio' : 'wait';
  const badgeText = type === 'prio' ? 'Prior.' : 'Aguard.';
  const extra = slaBad ? `<span class="badge sla">SLA</span>` : '';
  return `
    <li class="ticket-mini" data-num="${escapeHtml(t.numero)}">
      <div class="top">
        <div class="num">#${escapeHtml(t.numero)}</div>
        <div style="display:flex; gap:8px; align-items:center;">
          <span class="badge ${badgeClass}">${badgeText}</span>
          ${extra}
        </div>
      </div>
      <div class="title">${escapeHtml(t.assunto || '-')}</div>
      <div class="meta">
        <span>${escapeHtml(t.cliente || '-')}</span>
        <span>•</span>
        <span>${escapeHtml(t.status || '-')}</span>
        <span>•</span>
        <span>${d !== null ? (d + 'd') : '-'}</span>
      </div>
    </li>
  `;
}

function updateChips() {
  const chips = [];
  const add = (label, onRemove) => {
    const id = 'chip_' + Math.random().toString(36).slice(2);
    chips.push({ id, label, onRemove });
  };

  const q = $('searchInput').value.trim();
  if (q) add(`Busca: ${q}`, () => { $('searchInput').value = ''; applyFilters(); });

  if (statusSelected.size > 0) add(`Status: ${[...statusSelected].join(', ')}`, () => { statusSelected.clear(); $('statusMSMenu').querySelectorAll('.ms-item input').forEach(i => i.checked = false); updateStatusMSLabel(); applyFilters(); });

  const dep = $('departmentFilter').value;
  if (dep) add(`Departamento: ${dep}`, () => { $('departmentFilter').value = ''; applyFilters(); });

  const cli = $('clientFilter').value;
  if (cli) add(`Cliente: ${cli}`, () => { $('clientFilter').value = ''; applyFilters(); });

  const cat = $('categoryFilter') ? $('categoryFilter').value : '';
  if (cat) add(`Categoria: ${cat}`, () => { $('categoryFilter').value = ''; applyFilters(); });

  const pr = $('priorityFilter').value;
  if (pr) add(`Priorizado: ${pr}`, () => { $('priorityFilter').value = ''; applyFilters(); });

  const year = $('yearFilter').value;
  if (year) add(`Ano: ${year}`, () => { $('yearFilter').value = ''; applyFilters(); });

  const month = $('monthFilter').value;
  if (month) add(`Mês: ${month}`, () => { $('monthFilter').value = ''; applyFilters(); });

  const onlyOpen = $('onlyOpen').value === '1';
  if (onlyOpen) add(`Somente abertos`, () => { $('onlyOpen').value = ''; applyFilters(); });

  if (activeAgingBucket) {
    add(`Aging: ${activeAgingBucket}`, () => { activeAgingBucket = null; applyFilters(); });
  }

  if (activeKPI && activeKPI !== 'all') {
    const map = { open: 'KPI: Em aberto', resolved: 'KPI: Fechados', awaiting: 'KPI: Aguardando', prioritized: 'KPI: Priorizados', overdue: 'KPI: Fora do prazo' };
    add(map[activeKPI] || 'KPI', () => { setKPIFilter('all'); });
  }

  const container = $('activeChips');
  container.innerHTML = chips.map(c => `<span class="chip" id="${c.id}">${escapeHtml(c.label)} <span class="x">×</span></span>`).join('');
  chips.forEach(c => {
    const el = $(c.id);
    el.addEventListener('click', () => c.onRemove());
  });

  $('filterMeta').textContent = chips.length ? `Filtros ativos: ${chips.length}` : '';
}

function renderTable() {
  const tbody = $('ticketTableBody');
  $('tableMeta').textContent = `Linhas: ${filteredTickets.length}`;

  if (!filteredTickets.length) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding:26px; color:var(--muted);">Nenhum ticket para exibir (verifique filtros).</td></tr>`;
    return;
  }

  // Ordenação (gerencial):
  // 1) Abertos antes de fechados
  // 2) Priorizados (abertos) antes
  // 3) Mais recentes no topo (data de abertura; fallback última ação)
  // 4) Desempate por número (desc)
  const data = [...filteredTickets].sort((a, b) => {
    const aClosed = isClosedStatus(a.status);
    const bClosed = isClosedStatus(b.status);
    if (aClosed !== bClosed) return aClosed ? 1 : -1;

    const aPrio = (!aClosed && String(a.priorizado).toLowerCase() === 'sim') ? 1 : 0;
    const bPrio = (!bClosed && String(b.priorizado).toLowerCase() === 'sim') ? 1 : 0;
    if (aPrio !== bPrio) return bPrio - aPrio;

    const aOpen = parseDate(a.abertoEm)?.getTime() ?? 0;
    const bOpen = parseDate(b.abertoEm)?.getTime() ?? 0;
    if (aOpen !== bOpen) return bOpen - aOpen;

    const aKey = parseDate(a.ultimaAcao)?.getTime() ?? 0;
    const bKey = parseDate(b.ultimaAcao)?.getTime() ?? 0;
    if (aKey !== bKey) return bKey - aKey;

    const an = Number(String(a.numero).replace(/\D/g, '')) || 0;
    const bn = Number(String(b.numero).replace(/\D/g, '')) || 0;
    return bn - an;
  }).slice(0, 800);

  tbody.innerHTML = data.map(t => {
    const d = daysOpen(t);
    const isPrio = String(t.priorizado).toLowerCase() === 'sim' && !isClosedStatus(t.status);
    const ageClass = (d !== null && d > 15 && !isClosedStatus(t.status)) ? 'age high' : 'age';

    const st = statusBadge(t.status);
    const pr = isPrio ? `<span class="badge2 prio-sim">Sim</span>` : `<span class="badge2 prio-nao">${escapeHtml(t.priorizado || 'Nao')}</span>`;
    const age = (d !== null && !isClosedStatus(t.status)) ? `<span class="badge2 ${ageClass}">${d}d</span>` : `<span class="badge2 prio-nao">—</span>`;

    return `
      <tr class="row" onclick="showTicketModal('${escapeJs(t.numero)}')">
        <td><span class="tnum">#${escapeHtml(t.numero)}</span></td>
        <td>${escapeHtml(fmtDate(t.abertoEm))}</td>
        <td>${escapeHtml(t.departamento || '-')}</td>
        <td><div class="wrap2" title="${escapeHtml(t.solicitante || '')}">${escapeHtml(t.solicitante || '-')}</div></td>
        <td><div class="wrap2" title="${escapeHtml(t.cliente || '')}">${escapeHtml(t.cliente || '-')}</div></td>
        <td><div class="wrap2" title="${escapeHtml(t.servico || '')}">${escapeHtml(t.servico || '-')}</div></td>
        <td><div class="wrap2" title="${escapeHtml(t.assunto || '')}">${escapeHtml(t.assunto || '-')}</div></td>
        <td>${st}</td>
        <td>${escapeHtml(fmtDate(t.ultimaAcao))}</td>
        <td>${pr}</td>
        <td><div class="wrap2" title="${escapeHtml(t.responsavel || '')}">${escapeHtml(t.responsavel || '-')}</div></td>
        <td>${age}</td>
      </tr>
    `;
  }).join('');

  setupTopTableScroll();
}

/* ---------------------------
   Diretoria (Problema vs Adequação/Dúvida)
--------------------------- */
function setDiretoriaView(view) {
  diretoriaView = view || 'all';
  updateDiretoriaPage();
}

function updateDiretoriaPage() {
  const page = $('pageDiretoria');
  if (!page || page.style.display === 'none') return;

  const base = (filteredTickets || []).filter(t => !isClosedStatus(t.status));
  const problema = base.filter(t => classifyCategory(t) === 'problema');
  const adeqduv = base.filter(t => classifyCategory(t) !== 'problema');

  $('dirOpenTotal').textContent = base.length;
  $('dirProblema').textContent = problema.length;
  $('dirAdeqDuvida').textContent = adeqduv.length;

  const prioOpen = base.filter(t => String(t.priorizado).toLowerCase() === 'sim').length;
  const awaitOpen = base.filter(t => isAwaitingStatus(t.status)).length;
  $('dirPrioOpen').textContent = prioOpen;
  $('dirAwaitOpen').textContent = awaitOpen;

  const probOverdue = problema.filter(t => {
    const d = daysOpen(t);
    return d !== null && d > SLA_PROXY_DAYS;
  }).length;
  $('dirProblemaOverdue').textContent = probOverdue;

  const chipAll = $('dirChipAll'), chipP = $('dirChipProblema'), chipAD = $('dirChipAdeqDuvida');
  [chipAll, chipP, chipAD].forEach(el => el && el.classList.remove('active'));
  if (diretoriaView === 'problema') chipP?.classList.add('active');
  else if (diretoriaView === 'adeq_duvida') chipAD?.classList.add('active');
  else chipAll?.classList.add('active');

  $('dirKpiProblema')?.classList.toggle('active', diretoriaView === 'problema');
  $('dirKpiAdeq')?.classList.toggle('active', diretoriaView === 'adeq_duvida');
  $('dirKpiTotal')?.classList.toggle('active', diretoriaView === 'all');

  let viewData = base;
  if (diretoriaView === 'problema') viewData = problema;
  if (diretoriaView === 'adeq_duvida') viewData = adeqduv;

  $('dirTableMeta').textContent = `Linhas: ${viewData.length}`;

  const meta = [];
  if (diretoriaView === 'problema') meta.push('Filtro: Problema');
  if (diretoriaView === 'adeq_duvida') meta.push('Filtro: Adequação/Dúvida');
  $('dirTableFilterMeta').textContent = meta.join(' • ');

  renderDiretoriaTable(viewData);
}

function renderDiretoriaTable(list) {
  const tbody = $('dirTicketTableBody');
  if (!tbody) return;

  if (!list || !list.length) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding:26px; color:var(--muted);">Nenhum ticket em aberto para a visão selecionada.</td></tr>`;
    setupDirTopTableScroll();
    return;
  }

  // Ordenação (visão diretoria): priorizados primeiro e mais recentes no topo
  const data = [...list].sort((a, b) => {
    const aPrio = String(a.priorizado).toLowerCase() === 'sim' ? 1 : 0;
    const bPrio = String(b.priorizado).toLowerCase() === 'sim' ? 1 : 0;
    if (aPrio !== bPrio) return bPrio - aPrio;

    const aOpen = parseDate(a.abertoEm)?.getTime() ?? 0;
    const bOpen = parseDate(b.abertoEm)?.getTime() ?? 0;
    if (aOpen !== bOpen) return bOpen - aOpen;

    const aKey = parseDate(a.ultimaAcao)?.getTime() ?? 0;
    const bKey = parseDate(b.ultimaAcao)?.getTime() ?? 0;
    if (aKey !== bKey) return bKey - aKey;

    const an = Number(String(a.numero).replace(/\D/g, '')) || 0;
    const bn = Number(String(b.numero).replace(/\D/g, '')) || 0;
    return bn - an;
  }).slice(0, 800);

  tbody.innerHTML = data.map(t => {
    const d = daysOpen(t);
    const pr = (String(t.priorizado).toLowerCase() === 'sim') ? `<span class="badge2 prio-sim">Sim</span>` : `<span class="badge2 prio-nao">Não</span>`;
    const ageClass = (d !== null && d > 15) ? 'age high' : 'age';
    const age = d !== null ? `<span class="badge2 ${ageClass}">${d}d</span>` : `<span class="badge2 prio-nao">—</span>`;
    const cat = classifyCategory(t) === 'problema' ? `<span class="badge2 st-erro">Problema</span>` : `<span class="badge2 st-wait">Adeq/Dúvida</span>`;
    const st = statusBadge(t.status);

    return `
      <tr class="row" onclick="showTicketModal('${escapeJs(t.numero)}')">
        <td><span class="tnum">#${escapeHtml(t.numero)}</span></td>
        <td>${escapeHtml(fmtDate(t.abertoEm))}</td>
        <td>${cat}</td>
        <td>${escapeHtml(t.departamento || '-')}</td>
        <td><div class="wrap2" title="${escapeHtml(t.cliente || '')}">${escapeHtml(t.cliente || '-')}</div></td>
        <td><div class="wrap2" title="${escapeHtml(t.assunto || '')}">${escapeHtml(t.assunto || '-')}</div></td>
        <td>${st}</td>
        <td><div class="wrap2" title="${escapeHtml(t.responsavel || '')}">${escapeHtml(t.responsavel || '-')}</div></td>
        <td>${pr}</td>
        <td>${age}</td>
      </tr>
    `;
  }).join('');

  setupDirTopTableScroll();
}

function setupDirTopTableScroll() {
  const wrap = $('dirTableWrap');
  const top = $('dirTableXScrollTop');
  const inner = $('dirTableXScrollTopInner');
  if (!wrap || !top || !inner) return;

  const table = $('dirTicketTable');
  if (!table) return;

  inner.style.width = table.scrollWidth + 'px';

  top.onscroll = () => { wrap.scrollLeft = top.scrollLeft; };
  wrap.onscroll = () => { top.scrollLeft = wrap.scrollLeft; };

  top.scrollLeft = wrap.scrollLeft || 0;
}

function statusBadge(status) {
  const s = String(status || '').toLowerCase();
  let cls = 'st-fechado';
  if (s.includes('novo')) cls = 'st-novo';
  else if (s.includes('aguardando')) cls = 'st-aguardando';
  else if (s.includes('andamento') || s.includes('aberto')) cls = 'st-andamento';
  else if (s.includes('pausado')) cls = 'st-pausado';
  else if (s.includes('resolvido')) cls = 'st-resolvido';
  else if (s.includes('cancelado')) cls = 'st-cancelado';
  else if (s.includes('fechado')) cls = 'st-fechado';
  return `<span class="badge2 ${cls}">${escapeHtml(status || '-')}</span>`;
}

/* ---------------------------
   Charts
--------------------------- */
function countBy(field, onlyOpen = false) {
  const map = new Map();
  (filteredTickets || []).forEach(t => {
    if (onlyOpen && isClosedStatus(t.status)) return;
    const key = (t[field] || '').trim() || '(vazio)';
    map.set(key, (map.get(key) || 0) + 1);
  });
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function renderBarList(containerId, entries, limit = 10, clickHandler = null, colorMode = 'default') {
  const container = $(containerId);
  container.innerHTML = '';

  if (!entries.length) {
    container.innerHTML = `<div class="small" style="color:var(--muted); padding:6px 0;">Sem dados para exibir.</div>`;
    return;
  }

  const max = Math.max(...entries.map(e => e[1]));
  const shown = entries.slice(0, limit);

  shown.forEach(([k, v]) => {
    const item = document.createElement('div');
    item.className = 'baritem';
    item.innerHTML = `
      <div class="k" title="${escapeHtml(k)}">${escapeHtml(k)}</div>
      <div class="bar"><div class="fill"></div></div>
      <div class="v">${v}</div>
    `;

    const pct = max ? Math.round((v / max) * 100) : 0;
    const fill = item.querySelector('.fill');
    fill.style.width = pct + '%';

    if (colorMode === 'status') {
      fill.style.background = statusFill(k);
    } else if (colorMode === 'department') {
      fill.style.background = deptFill(k);
    } else if (colorMode === 'client') {
      fill.style.background = 'linear-gradient(90deg, rgba(124,58,237,.85), rgba(37,99,235,.65))';
    } else if (colorMode === 'category') {
      fill.style.background = 'linear-gradient(90deg, rgba(22,163,74,.85), rgba(37,99,235,.55))';
    }

    if (clickHandler) {
      item.addEventListener('click', () => clickHandler(k));
    }
    container.appendChild(item);
  });
}

function statusFill(label) {
  const s = String(label || '').toLowerCase();
  if (s.includes('novo')) return 'linear-gradient(90deg, rgba(37,99,235,.92), rgba(59,130,246,.62))';
  if (s.includes('aguardando')) return 'linear-gradient(90deg, rgba(245,158,11,.92), rgba(251,191,36,.58))';
  if (s.includes('andamento') || s.includes('aberto')) return 'linear-gradient(90deg, rgba(22,163,74,.90), rgba(34,197,94,.55))';
  if (s.includes('pausado')) return 'linear-gradient(90deg, rgba(124,58,237,.90), rgba(167,139,250,.55))';
  if (s.includes('resolvido')) return 'linear-gradient(90deg, rgba(59,130,246,.88), rgba(147,197,253,.55))';
  if (s.includes('cancelado')) return 'linear-gradient(90deg, rgba(239,68,68,.92), rgba(252,165,165,.55))';
  if (s.includes('fechado')) return 'linear-gradient(90deg, rgba(100,116,139,.90), rgba(148,163,184,.55))';
  return 'linear-gradient(90deg, rgba(37,99,235,.85), rgba(59,130,246,.65))';
}
function deptFill(_) {
  return 'linear-gradient(90deg, rgba(37,99,235,.85), rgba(124,58,237,.55))';
}

function updateCharts() {
  const dept = countBy('departamento', false);
  renderBarList('departmentChart', dept, 10, (k) => {
    if (k === '(vazio)') return;
    $('departmentFilter').value = k;
    applyFilters();
  }, 'department');

  const st = countBy('status', false);
  renderBarList('statusChart', st, 10, (k) => {
    if (k === '(vazio)') return;
    toggleStatusFromChart(k);
  }, 'status');

  const cat = countBy('categoria', false);
  renderBarList('categoryChart', cat, 10, (k) => {
    $('searchInput').value = k === '(vazio)' ? '' : k;
    applyFilters();
  }, 'category');

  const open = countBy('cliente', true);
  renderBarList('clientOpenChart', open, 8, (k) => {
    if (k === '(vazio)') return;
    $('clientFilter').value = k;
    applyFilters();
  }, 'client');
}

function toggleStatusFromChart(statusLabel) {
  if (statusSelected.has(statusLabel)) statusSelected.delete(statusLabel);
  else statusSelected.add(statusLabel);

  const items = $('statusMSMenu').querySelectorAll('.ms-item');
  items.forEach(it => {
    const st = it.querySelector('span').textContent;
    if (st === statusLabel) {
      it.querySelector('input').checked = statusSelected.has(statusLabel);
    }
  });
  updateStatusMSLabel();
  applyFilters();
}

/* Line chart - Chart.js */
function updateLineChart() {
  const canvas = $('lineChart');
  if (!canvas) return;

  if (typeof Chart === 'undefined') {
    return;
  }

  const map = new Map();
  filteredTickets.forEach(t => {
    const d = parseDate(t.abertoEm);
    if (!d || d.getFullYear() < 2000) return;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    map.set(key, (map.get(key) || 0) + 1);
  });

  const keys = [...map.keys()].sort();
  const lastKeys = keys.slice(-24);
  const labels = lastKeys.map(k => {
    const [yy, mm] = k.split('-');
    const mnames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    return `${mnames[parseInt(mm, 10) - 1]}/${yy.slice(2)}`;
  });
  const values = lastKeys.map(k => map.get(k));

  if (lineChartInstance) {
    lineChartInstance.data.labels = labels;
    lineChartInstance.data.datasets[0].data = values;
    lineChartInstance.update();
    return;
  }

  lineChartInstance = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Tickets',
        data: values,
        tension: 0.28,
        fill: true,
        borderWidth: 2,
        pointRadius: 3,
        pointHoverRadius: 5,
        borderColor: 'rgba(37,99,235,.95)',
        backgroundColor: 'rgba(37,99,235,.10)'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => ` ${ctx.parsed.y} ticket(s)` } }
      },
      scales: {
        x: { grid: { color: 'rgba(148,163,184,.25)' }, ticks: { color: '#475569', maxRotation: 0, autoSkip: true } },
        y: { grid: { color: 'rgba(148,163,184,.25)' }, ticks: { color: '#475569' } }
      },
      onClick: (evt, elements) => {
        if (!elements?.length) return;
        const idx = elements[0].index;
        const clickedKey = lastKeys[idx];
        if (!clickedKey) return;

        const [yy, mm] = clickedKey.split('-');
        $('yearFilter').value = yy;
        $('monthFilter').value = mm;
        applyFilters();
      }
    }
  });
}

/* ---------------------------
   Modal / Ticket detail (Painel)
--------------------------- */
let _currentModalTicket = null;

function showTicketModal(ticketNumber) {
  const t = allTickets.find(x => String(x.numero) === String(ticketNumber));
  if (!t) return;

  _currentModalTicket = t;

  $('modalSubject').textContent = t.assunto || 'Sem assunto';
  const dOpen = daysOpen(t);
  const ageText = (!isClosedStatus(t.status) && dOpen !== null) ? `${dOpen} dia(s) em aberto` : '—';
  const meta = [
    `<span><strong>Ticket</strong> #${escapeHtml(t.numero)}</span>`,
    `<span>•</span>`,
    `<span>${statusBadge(t.status)}</span>`,
    `<span>•</span>`,
    `<span>${escapeHtml(ageText)}</span>`
  ].join(' ');
  $('modalMeta').innerHTML = meta;

  const grid = $('modalDetailGrid');
  grid.innerHTML = [
    dField('Número', '#' + t.numero),
    dField('Aberto em', fmtDate(t.abertoEm)),
    dField('Departamento', t.departamento || '-'),
    dField('Usuário solicitante', t.solicitante || '-'),
    dField('Cliente (Pessoa)', t.cliente || '-'),
    dField('Serviço', t.servico || '-'),
    dField('Responsável', t.responsavel || '-'),
    dField('Categoria', t.categoria || '-'),
    dField('Data da última ação', fmtDate(t.ultimaAcao)),
    dField('Status', t.status || '-'),
    dField('Ticket priorizado', (String(t.priorizado).toLowerCase() === 'sim' ? 'Sim' : 'Não')),
    dField('Tempo em aberto', (!isClosedStatus(t.status) && dOpen !== null) ? (dOpen + ' dia(s)') : '—')
  ].join('');

  $('modalOverlay').classList.add('open');
}

function dField(label, value) {
  return `<div class="dcard"><div class="dl">${escapeHtml(label)}</div><div class="dv">${escapeHtml(String(value ?? '-'))}</div></div>`;
}

function closeModal(e) {
  if (e && e.target && e.target !== $('modalOverlay')) return;
  $('modalOverlay').classList.remove('open');
}

function copyTicket() {
  if (!_currentModalTicket) return;
  const t = _currentModalTicket;
  const lines = [
    `Ticket #${t.numero}`,
    `Assunto: ${t.assunto || '-'}`,
    `Cliente: ${t.cliente || '-'}`,
    `Departamento: ${t.departamento || '-'}`,
    `Status: ${t.status || '-'}`,
    `Prioridade: ${(String(t.priorizado).toLowerCase() === 'sim' ? 'Sim' : 'Não')}`,
    `Aberto em: ${fmtDate(t.abertoEm)}`,
    `Última ação: ${fmtDate(t.ultimaAcao)}`,
    `Responsável: ${t.responsavel || '-'}`
  ].join('\n');

  navigator.clipboard?.writeText(lines)
    .then(() => alert('Resumo copiado para a área de transferência.'))
    .catch(() => alert('Não foi possível copiar automaticamente. Seu navegador pode bloquear.'));
}

function openTicketPageFromModal() {
  if (!_currentModalTicket) return;
  const num = _currentModalTicket.numero;
  closeModal();
  showPage('ticket');
  $('ticketSearchInput').value = num;
  searchTicket();
}

/* ---------------------------
   Modal / Ticket detail (Comparação)
   - Mostra tudo que mudou naquele ticket (antes/depois)
--------------------------- */
function showCompareDiffModal(ticketNumber) {
  const num = String(ticketNumber);

  const oldT = compareOldMap.get(num) || null;
  const newT = compareNewMap.get(num) || null;

  // linhas do diff apenas desse ticket
  const changes = (compareDiffRows || []).filter(r => String(r.numero) === num && r.tipo === 'ALTERADO');

  // tipo macro (se existe só de um lado)
  let macroType = 'ALTERADO';
  if (oldT && !newT) macroType = 'REMOVIDO';
  if (!oldT && newT) macroType = 'NOVO';

  const subject = (newT?.assunto || oldT?.assunto || `Ticket #${num}`);
  $('modalSubject').textContent = subject || 'Sem assunto';

  const statusText = newT?.status || oldT?.status || '-';
  const openDays = newT ? daysOpen(newT) : (oldT ? daysOpen(oldT) : null);
  const ageText = (openDays !== null && statusText && !isClosedStatus(statusText)) ? `${openDays} dia(s) em aberto` : '—';

  const badgeTipo = macroType === 'NOVO'
    ? `<span class="badge2 st-andamento">NOVO</span>`
    : (macroType === 'REMOVIDO'
      ? `<span class="badge2 st-cancelado">REMOVIDO</span>`
      : `<span class="badge2 st-aguardando">ALTERADO</span>`);

  $('modalMeta').innerHTML = [
    `<span><strong>Ticket</strong> #${escapeHtml(num)}</span>`,
    `<span>•</span>`,
    badgeTipo,
    `<span>•</span>`,
    `<span>${statusBadge(statusText)}</span>`,
    `<span>•</span>`,
    `<span>${escapeHtml(ageText)}</span>`
  ].join(' ');

  const grid = $('modalDetailGrid');

  // helper para snapshot
  const snapBlock = (title, t, sideLabel) => {
    if (!t) {
      return `
        <div class="dcard" style="grid-column:1/-1;">
          <div class="dl">${escapeHtml(title)}</div>
          <div class="dv">${escapeHtml(sideLabel)}: não existe neste arquivo.</div>
        </div>
      `;
    }

    const dOpen = daysOpen(t);
    const tempo = (!isClosedStatus(t.status) && dOpen !== null) ? (dOpen + ' dia(s)') : '—';

    return `
      <div class="dcard" style="grid-column:1/-1; border-color: rgba(37,99,235,.22); background: color-mix(in srgb, var(--panel2) 88%, transparent);">
        <div class="dl">${escapeHtml(title)}</div>
        <div class="dv" style="margin-top:8px; font-weight:800;">
          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px;">
            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Assunto</div><div style="margin-top:6px;">${escapeHtml(t.assunto || '-')}</div></div>
            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Status</div><div style="margin-top:6px;">${escapeHtml(t.status || '-')}</div></div>

            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Aberto em</div><div style="margin-top:6px;">${escapeHtml(fmtDate(t.abertoEm))}</div></div>
            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Última ação</div><div style="margin-top:6px;">${escapeHtml(fmtDate(t.ultimaAcao))}</div></div>

            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Cliente</div><div style="margin-top:6px;">${escapeHtml(t.cliente || '-')}</div></div>
            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Departamento</div><div style="margin-top:6px;">${escapeHtml(t.departamento || '-')}</div></div>

            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Serviço</div><div style="margin-top:6px;">${escapeHtml(t.servico || '-')}</div></div>
            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Responsável</div><div style="margin-top:6px;">${escapeHtml(t.responsavel || '-')}</div></div>

            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Categoria</div><div style="margin-top:6px;">${escapeHtml(t.categoria || '-')}</div></div>
            <div><div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Priorizado</div><div style="margin-top:6px;">${escapeHtml(String(t.priorizado || 'Nao'))}</div></div>

            <div style="grid-column:1/-1;">
              <div style="font-size:11px; color:var(--muted); font-weight:900; text-transform:uppercase;">Tempo em aberto</div>
              <div style="margin-top:6px;">${escapeHtml(tempo)}</div>
            </div>
          </div>
        </div>
      </div>
    `;
  };

  const changesBlock = () => {
    if (macroType === 'NOVO') {
      return `
        <div class="dcard" style="grid-column:1/-1; border-color: rgba(22,163,74,.22);">
          <div class="dl">Alterações detectadas</div>
          <div class="dv">Ticket novo: não existe comparação de campos (não existia no Excel velho).</div>
        </div>
      `;
    }
    if (macroType === 'REMOVIDO') {
      return `
        <div class="dcard" style="grid-column:1/-1; border-color: rgba(239,68,68,.25);">
          <div class="dl">Alterações detectadas</div>
          <div class="dv">Ticket removido: não existe comparação de campos (não existe no Excel novo).</div>
        </div>
      `;
    }

    if (!changes.length) {
      return `
        <div class="dcard" style="grid-column:1/-1;">
          <div class="dl">Alterações detectadas</div>
          <div class="dv">Nenhuma mudança de campo registrada para este ticket (pode ter entrado na lista por regra/normalização).</div>
        </div>
      `;
    }

    const rows = changes.map(c => {
      return `
        <tr>
          <td style="padding:8px 10px; border-bottom:1px solid var(--border); font-weight:900; color:var(--slate);">${escapeHtml(c.campo)}</td>
          <td style="padding:8px 10px; border-bottom:1px solid var(--border); color:var(--muted);"><div class="wrap2" title="${escapeHtml(c.antes)}">${escapeHtml(c.antes)}</div></td>
          <td style="padding:8px 10px; border-bottom:1px solid var(--border); color:var(--slate);"><div class="wrap2" title="${escapeHtml(c.depois)}">${escapeHtml(c.depois)}</div></td>
        </tr>
      `;
    }).join('');

    return `
      <div class="dcard" style="grid-column:1/-1;">
        <div class="dl">Alterações detectadas</div>
        <div class="dv" style="margin-top:8px;">
          <div style="font-size:12px; color:var(--muted); font-weight:800; margin-bottom:8px;">
            Total de campos alterados: <strong>${changes.length}</strong>
          </div>
          <div style="border:1px solid var(--border); border-radius:12px; overflow:hidden;">
            <table style="width:100%; border-collapse:collapse; min-width: 560px;">
              <thead>
                <tr>
                  <th style="text-align:left; padding:9px 10px; font-size:11px; letter-spacing:.35px; text-transform:uppercase; color:var(--muted); background: color-mix(in srgb, var(--panel2) 88%, transparent); border-bottom:1px solid var(--border);">Campo</th>
                  <th style="text-align:left; padding:9px 10px; font-size:11px; letter-spacing:.35px; text-transform:uppercase; color:var(--muted); background: color-mix(in srgb, var(--panel2) 88%, transparent); border-bottom:1px solid var(--border);">Antes</th>
                  <th style="text-align:left; padding:9px 10px; font-size:11px; letter-spacing:.35px; text-transform:uppercase; color:var(--muted); background: color-mix(in srgb, var(--panel2) 88%, transparent); border-bottom:1px solid var(--border);">Depois</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  };

  // monta grid completo
  grid.innerHTML = `
    ${snapBlock('Snapshot — Excel velho', oldT, 'Velho')}
    ${snapBlock('Snapshot — Excel novo', newT, 'Novo')}
    ${changesBlock()}
  `;

  // marca modal como aberto
  $('modalOverlay').classList.add('open');

  // define o "current" para botões padrão (copiar/abrir)
  // preferimos o ticket "novo", se existir; senão velho.
  _currentModalTicket = newT || oldT || null;
}

/* ---------------------------
   Ticket Page search
--------------------------- */
function searchTicket() {
  const num = $('ticketSearchInput').value.trim();
  const panel = $('ticketDetailPanel');
  const grid = $('ticketDetailGrid');
  if (!num) {
    panel.style.display = 'none';
    return;
  }

  const t = allTickets.find(x => String(x.numero) === String(num));
  if (!t) {
    panel.style.display = 'block';
    $('ticketDetailTitle').textContent = 'Ticket não encontrado';
    $('ticketDetailSubtitle').textContent = 'Verifique o número e tente novamente.';
    grid.innerHTML = '';
    return;
  }

  panel.style.display = 'block';
  $('ticketDetailTitle').textContent = `#${t.numero} — ${t.assunto || 'Sem assunto'}`;
  $('ticketDetailSubtitle').textContent = `${t.status || '-'} • Cliente: ${t.cliente || '-'} • Departamento: ${t.departamento || '-'}`;

  const dOpen = daysOpen(t);
  grid.innerHTML = [
    dField('Assunto', t.assunto || '-'),
    dField('Status', t.status || '-'),
    dField('Prioridade', (String(t.priorizado).toLowerCase() === 'sim' ? 'Sim' : 'Não')),
    dField('Aberto em', fmtDate(t.abertoEm)),
    dField('Última ação', fmtDate(t.ultimaAcao)),
    dField('Tempo em aberto', (!isClosedStatus(t.status) && dOpen !== null) ? (dOpen + ' dia(s)') : '—'),
    dField('Cliente (Pessoa)', t.cliente || '-'),
    dField('Usuário solicitante', t.solicitante || '-'),
    dField('Departamento', t.departamento || '-'),
    dField('Serviço', t.servico || '-'),
    dField('Responsável', t.responsavel || '-'),
    dField('Categoria', t.categoria || '-')
  ].join('');
}

/* ---------------------------
   Export filtered
--------------------------- */
function exportFiltered() {
  if (!filteredTickets.length) {
    alert('Não há dados filtrados para exportar.');
    return;
  }
  if (typeof XLSX === 'undefined') {
    alert('XLSX não carregou. Verifique conexão com internet.');
    return;
  }
  const exportData = filteredTickets.map(t => ({
    Numero: t.numero,
    AbertoEm: t.abertoEm,
    Departamento: t.departamento,
    Solicitante: t.solicitante,
    Cliente: t.cliente,
    Servico: t.servico,
    Assunto: t.assunto,
    Responsavel: t.responsavel,
    Categoria: t.categoria,
    UltimaAcao: t.ultimaAcao,
    Status: t.status,
    Priorizado: t.priorizado
  }));
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Tickets');
  XLSX.writeFile(wb, `tickets_filtrados_${new Date().toISOString().split('T')[0]}.xlsx`);
}

/* ---------------------------
   Utils
--------------------------- */
function escapeHtml(str) {
  return String(str ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", "&#039;");
}
function escapeJs(str) {
  return String(str ?? '').replaceAll("\\", "\\\\").replaceAll("'", "\\'");
}

/* ---------------------------
   Comparar Excel (velho vs novo)
--------------------------- */
function triggerCompareOld() { $('compareOldInput')?.click(); }
function triggerCompareNew() { $('compareNewInput')?.click(); }

function handleCompareFile(event, which) {
  const file = event.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    alert('Biblioteca XLSX não carregou.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'dd/mm/yyyy' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: 'dd/mm/yyyy' });

      const tickets = rows.map(normalizeTicket).filter(t => String(t.numero || '').trim() !== '');

      if (which === 'old') {
        compareOldTickets = tickets;
        $('cmpOldName').textContent = file.name;
      } else {
        compareNewTickets = tickets;
        $('cmpNewName').textContent = file.name;
      }

      // limpa resultados para evitar confusão
      compareDiffRows = [];
      compareOldMap = new Map();
      compareNewMap = new Map();
      renderCompareTable();
      updateCompareKpis();

    } catch (err) {
      console.error(err);
      alert('Falha ao ler o arquivo de comparação. Verifique se é um Excel válido.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function runCompare() {
  if (!compareOldTickets.length || !compareNewTickets.length) {
    alert('Carregue o Excel velho e o Excel novo antes de comparar.');
    return;
  }

  // guarda mapas globais para abrir modal depois
  compareOldMap = new Map(compareOldTickets.map(t => [String(t.numero).trim(), t]));
  compareNewMap = new Map(compareNewTickets.map(t => [String(t.numero).trim(), t]));

  const allNums = new Set([...compareOldMap.keys(), ...compareNewMap.keys()]);
  const diffs = [];

  const fields = [
    { key: 'status', label: 'Status' },
    { key: 'responsavel', label: 'Responsável' },
    { key: 'categoria', label: 'Categoria' },
    { key: 'priorizado', label: 'Ticket priorizado' },
    { key: 'assunto', label: 'Assunto' },
    { key: 'cliente', label: 'Cliente' },
    { key: 'departamento', label: 'Departamento' },
    { key: 'servico', label: 'Serviço' },
    { key: 'abertoEm', label: 'Aberto em' },
    { key: 'ultimaAcao', label: 'Data da última ação' }
  ];

  allNums.forEach(num => {
    const o = compareOldMap.get(num);
    const n = compareNewMap.get(num);

    if (o && !n) {
      diffs.push({ numero: num, tipo: 'REMOVIDO', campo: '—', antes: 'Existe no velho', depois: 'Não existe no novo' });
      return;
    }
    if (!o && n) {
      diffs.push({ numero: num, tipo: 'NOVO', campo: '—', antes: 'Não existe no velho', depois: 'Existe no novo' });
      return;
    }

    // ambos
    fields.forEach(f => {
      const a = String(o[f.key] ?? '').trim();
      const b = String(n[f.key] ?? '').trim();
      if (a !== b) {
        diffs.push({ numero: num, tipo: 'ALTERADO', campo: f.label, antes: a || '—', depois: b || '—' });
      }
    });
  });

  const rank = { ALTERADO: 0, NOVO: 1, REMOVIDO: 2 };
  compareDiffRows = diffs.sort((x, y) => {
    const rx = rank[x.tipo] ?? 9;
    const ry = rank[y.tipo] ?? 9;
    if (rx !== ry) return rx - ry;
    return String(x.numero).localeCompare(String(y.numero), 'pt-BR', { numeric: true });
  });

  compareFilterMode = 'all';
  updateCompareKpis();
  renderCompareTable();
  setupCompareTopTableScroll();
}

function updateCompareKpis() {
  const added = compareDiffRows.filter(r => r.tipo === 'NOVO').length;
  const removed = compareDiffRows.filter(r => r.tipo === 'REMOVIDO').length;
  const changed = compareDiffRows.filter(r => r.tipo === 'ALTERADO').length;

  $('cmpAdded').textContent = added;
  $('cmpRemoved').textContent = removed;
  $('cmpChanged').textContent = changed;

  const total = new Set(compareDiffRows.map(r => r.numero)).size;
  $('cmpTotal').textContent = total;

  $('cmpKpiAll')?.classList.toggle('active', compareFilterMode === 'all');
  $('cmpKpiChanged')?.classList.toggle('active', compareFilterMode === 'changed');
}

function setCompareFilter(mode) {
  compareFilterMode = mode || 'all';
  updateCompareKpis();
  renderCompareTable();
  setupCompareTopTableScroll();
}

function renderCompareTable() {
  const tbody = $('cmpTableBody');
  if (!tbody) return;

  const rows = (compareDiffRows || []);

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:26px; color:var(--muted);">Nenhuma diferença carregada. Clique em “Comparar”.</td></tr>`;
    return;
  }

  // ===== Agrupa por ticket =====
  const groups = new Map(); // num -> {numero, macroTipo's, changes[]}
  for (const r of rows) {
    const num = String(r.numero);
    if (!groups.has(num)) {
      groups.set(num, { numero: num, tipoMacro: null, changes: [] });
    }

    const g = groups.get(num);

    // Se for NOVO/REMOVIDO, isso vira tipo macro do ticket
    if (r.tipo === 'NOVO' || r.tipo === 'REMOVIDO') {
      g.tipoMacro = r.tipo;
    } else if (r.tipo === 'ALTERADO') {
      g.changes.push(r);
      if (!g.tipoMacro) g.tipoMacro = 'ALTERADO';
    }
  }

  // ===== Monta lista ordenada =====
  const orderRank = { ALTERADO: 0, NOVO: 1, REMOVIDO: 2 };
  const data = [...groups.values()].sort((a, b) => {
    const ra = orderRank[a.tipoMacro] ?? 9;
    const rb = orderRank[b.tipoMacro] ?? 9;
    if (ra !== rb) return ra - rb;
    return a.numero.localeCompare(b.numero, 'pt-BR', { numeric: true });
  });

  // ===== Filtro "somente alterados" =====
  const filtered = (compareFilterMode === 'changed')
    ? data.filter(g => g.tipoMacro === 'ALTERADO')
    : data;

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:26px; color:var(--muted);">Nenhum ticket alterado para exibir nesse filtro.</td></tr>`;
    return;
  }

  // ===== Render =====
  tbody.innerHTML = filtered.slice(0, 2000).map(g => {
    const cls = g.tipoMacro === 'ALTERADO' ? 'st-wait' : (g.tipoMacro === 'NOVO' ? 'st-ok' : 'st-erro');

    // Box de comparações (somente para ALTERADO)
    let boxHtmlAntes = '';
    let boxHtmlDepois = '';

    if (g.tipoMacro === 'ALTERADO') {
      boxHtmlAntes = `
    <div class="cmpBox">
      ${g.changes.map(c => `
        <div class="cmpRow">
          <div class="cmpCampo">${escapeHtml(c.campo)}</div>
          <div class="cmpVal cmpAntes" title="${escapeHtml(c.antes)}">${escapeHtml(c.antes)}</div>
        </div>
      `).join('')}
    </div>
  `;

      boxHtmlDepois = `
    <div class="cmpBox">
      ${g.changes.map(c => `
        <div class="cmpRow">
          <div class="cmpCampo">${escapeHtml(c.campo)}</div>
          <div class="cmpVal cmpDepois" title="${escapeHtml(c.depois)}">${escapeHtml(c.depois)}</div>
        </div>
      `).join('')}
    </div>
  `;
    } else {
      const msg = g.tipoMacro === 'NOVO'
        ? 'Ticket existe apenas no Excel novo.'
        : 'Ticket existe apenas no Excel velho.';
      boxHtmlAntes = `<div class="cmpBox cmpBoxInfo">${msg}</div>`;
      boxHtmlDepois = `<div class="cmpBox cmpBoxInfo">${msg}</div>`;
    }


    const count = (g.tipoMacro === 'ALTERADO') ? (g.changes?.length || 0) : 0;

    return `
      <tr class="row" onclick="showCompareDiffModal('${escapeJs(g.numero)}')">
        <td style="width:140px;"><span class="tnum">#${escapeHtml(g.numero)}</span></td>
        <td style="width:140px;">
          <span class="badge2 ${cls}">${escapeHtml(g.tipoMacro)}</span>
          ${g.tipoMacro === 'ALTERADO' ? `<div class="small" style="margin-top:6px; color:var(--muted); font-weight:700;">${count} campo(s)</div>` : ``}
        </td>
        <td>${boxHtmlAntes}</td>
        <td>${boxHtmlDepois}</td>

      </tr>
    `;
  }).join('');
}


function setupCompareTopTableScroll() {
  const wrap = $('cmpTableWrap');
  const top = $('cmpTableXScrollTop');
  const inner = $('cmpTableXScrollTopInner');
  if (!wrap || !top || !inner) return;

  const table = $('cmpTable');
  if (!table) return;

  inner.style.width = table.scrollWidth + 'px';
  top.onscroll = () => { wrap.scrollLeft = top.scrollLeft; };
  wrap.onscroll = () => { top.scrollLeft = wrap.scrollLeft; };
  top.scrollLeft = wrap.scrollLeft || 0;
}

function exportCompareCSV() {
  if (!compareDiffRows || !compareDiffRows.length) {
    alert('Nenhuma diferença para exportar.');
    return;
  }
  const header = ['numero', 'tipo', 'campo', 'antes', 'depois'];
  const lines = [header.join(';')].concat(compareDiffRows.map(r => [
    r.numero, r.tipo, r.campo,
    String(r.antes || '').replaceAll(';', ','),
    String(r.depois || '').replaceAll(';', ',')
  ].join(';')));

  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `comparacao_tickets_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---------------------------
   Boot
--------------------------- */
function init() {
  // default KPI selection
  $('kpiTotal').classList.add('active');

  // tema
  setTheme(loadTheme());

  // load cache if exists
  const cache = loadCache();
  if (cache?.data?.length) {
    allTickets = cache.data;
    setFileLabel(cache.fileName || 'cache', allTickets.length);
    buildFilterOptions();
    buildStatusMultiSelect();
    clearAllFilters(false);
    updateEverything();
    setupTopTableScroll();
  } else {
    $('cacheFile').textContent = 'nenhum arquivo carregado';
    $('cacheCount').textContent = '0';
  }

  window.addEventListener('resize', () => {
    if (lineChartInstance) lineChartInstance.resize();
    setupTopTableScroll();
    setupCompareTopTableScroll();
  });
}

init();
