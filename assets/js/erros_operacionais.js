let all = [];
let filtered = [];
let activeKpi = 'all';

function $(id) { return document.getElementById(id); }

function toggleTheme() {
  const app = $('app');
  const next = app.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  app.setAttribute('data-theme', next);
  try { localStorage.setItem('occ_theme', next); } catch (_) { }
}
(function bootTheme() {
  try {
    const t = localStorage.getItem('occ_theme');
    if (t) $('app').setAttribute('data-theme', t);
  } catch (_) { }
})();

function triggerFile() { $('fileInput').click(); }

function normKey(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}
function findValue(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const cn = normKey(c);
    const found = keys.find(k => normKey(k).includes(cn));
    if (found !== undefined && row[found] !== undefined && row[found] !== null && String(row[found]).trim() !== '')
      return String(row[found]).trim();
  }
  return '';
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = String(v).trim();
  if (!s) return null;

  // Excel serial (às vezes vem como número)
  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = parseFloat(s);
    if (serial > 1 && serial < 100000) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      const ms = epoch.getTime() + serial * 86400000;
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

function normalizeRow(row) {
  const ticket = findValue(row, ['ticket', 'chamado', 'id', 'numero', 'número']);
  let ticketFix = ticket;
  if (/^\d+\.0+$/.test(ticketFix)) ticketFix = ticketFix.split('.')[0];

  const departamento = findValue(row, ['departamento', 'depto', 'setor', 'área', 'area']);
  const responsavel = findValue(row, ['responsavel', 'responsável', 'responsavel ti', 'owner', 'atribuido', 'atribuído']);
  const tipo = findValue(row, ['tipo']);
  const dataAbert = findValue(row, ['data abert', 'data abertura', 'aberto em', 'abertura']);
  const titulo = findValue(row, ['titulo / descricao', 'título / descrição', 'titulo', 'título', 'descricao', 'descrição']);
  const aliare = findValue(row, ['aliare']);
  const prioridade = findValue(row, ['prioridade', 'prioridad']);
  const impacto = findValue(row, ['pimpacto', 'impacto']);
  const validacao = findValue(row, ['validacao ti', 'validação ti', 'observacao ti', 'observação ti', 'validacao', 'validação']);
  const statusAtual = findValue(row, ['status at', 'status atual', 'status']);
  const tipoOc = findValue(row, ['tipo de ocorrencia', 'tipo de ocorrência', 'ocorrencia', 'ocorrência']);

  return {
    ticket: ticketFix,
    departamento,
    responsavel,
    tipo,
    dataAbert,
    titulo,
    aliare,
    prioridade,
    impacto,
    validacao,
    statusAtual,
    tipoOc,
    _raw: row
  };
}

function isDone(status) {
  const s = String(status || '').toLowerCase();
  return s.includes('resolvido') || s.includes('fechado') || s.includes('encerr') || s.includes('conclu');
}
function isErroOperacional(tipoOc) {
  const s = String(tipoOc || '').toLowerCase();
  return s.includes('erro operacional');
}
function isHighImpact(impacto) {
  const s = String(impacto || '').toLowerCase();
  return s.includes('3') || s.includes('alta') || s.includes('4') || s.includes('urgente');
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    alert('Biblioteca XLSX não carregou. Verifique conexão com internet ou use uma cópia local.');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'dd/mm/yyyy' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: 'dd/mm/yyyy' });

      all = rows.map(normalizeRow).filter(r => String(r.ticket || '').trim() !== '');
      cacheSave(file.name, all);
      $('fileTag').textContent = `${file.name} • ${all.length} linhas`;

      buildOptions();
      clearAll(false);
      applyFilters();
    } catch (err) {
      console.error(err);
      alert('Falha ao ler o Excel. Confirme se a primeira planilha contém os dados.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function buildOptions() {
  const setTipoOc = new Set(), setImpacto = new Set(), setStatus = new Set(), setDept = new Set();
  all.forEach(r => {
    if (r.tipoOc) setTipoOc.add(r.tipoOc);
    if (r.impacto) setImpacto.add(r.impacto);
    if (r.statusAtual) setStatus.add(r.statusAtual);
    if (r.departamento) setDept.add(r.departamento);
  });

  $('tipoOc').innerHTML = `<option value="">Todos</option>` + [...setTipoOc].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  $('impacto').innerHTML = `<option value="">Todos</option>` + [...setImpacto].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  $('statusAt').innerHTML = `<option value="">Todos</option>` + [...setStatus].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  $('dept').innerHTML = `<option value="">Todos</option>` + [...setDept].sort((a, b) => a.localeCompare(b, 'pt-BR')).map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
}

function setKpi(k) {
  activeKpi = k;
  ['kpiAll', 'kpiOpen', 'kpiDone', 'kpiErrOp', 'kpiHigh'].forEach(id => $(id).classList.remove('active'));
  const map = { all: 'kpiAll', open: 'kpiOpen', done: 'kpiDone', errop: 'kpiErrOp', high: 'kpiHigh' };
  $(map[k] || 'kpiAll').classList.add('active');
  applyFilters();
}

function clearAll(resetKpi = true) {
  $('q').value = '';
  $('tipoOc').value = '';
  $('impacto').value = '';
  $('statusAt').value = '';
  $('dept').value = '';
  $('onlyOpen').value = '';
  if (resetKpi) setKpi('all'); else updateUi();
}

let _deb = null;
function applyDebounced() {
  clearTimeout(_deb);
  _deb = setTimeout(applyFilters, 180);
}

function applyFilters() {
  const q = $('q').value.trim().toLowerCase();
  const tipoOc = $('tipoOc').value;
  const impacto = $('impacto').value;
  const statusAt = $('statusAt').value;
  const dept = $('dept').value;
  const onlyOpen = $('onlyOpen').value === '1';

  filtered = all.filter(r => {
    if (activeKpi === 'open' && isDone(r.statusAtual)) return false;
    if (activeKpi === 'done' && !isDone(r.statusAtual)) return false;
    if (activeKpi === 'errop' && !isErroOperacional(r.tipoOc)) return false;
    if (activeKpi === 'high' && !isHighImpact(r.impacto)) return false;

    if (onlyOpen && isDone(r.statusAtual)) return false;

    if (tipoOc && r.tipoOc !== tipoOc) return false;
    if (impacto && r.impacto !== impacto) return false;
    if (statusAt && r.statusAtual !== statusAt) return false;
    if (dept && r.departamento !== dept) return false;

    if (q) {
      const hay = [
        r.ticket, r.titulo, r.departamento, r.responsavel, r.tipoOc, r.impacto, r.statusAtual, r.validacao
      ].map(x => String(x || '').toLowerCase());
      if (!hay.some(x => x.includes(q))) return false;
    }

    return true;
  });

  updateUi();
}

function updateUi() {
  updateChips();
  updateKpis();
  renderTable();
  renderCharts();
}

function updateKpis() {
  const total = all.length;
  const done = all.filter(r => isDone(r.statusAtual)).length;
  const open = total - done;
  const errOp = all.filter(r => isErroOperacional(r.tipoOc)).length;
  const high = all.filter(r => isHighImpact(r.impacto)).length;

  $('kAll').textContent = total;
  $('kOpen').textContent = open;
  $('kDone').textContent = done;
  $('kErrOp').textContent = errOp;
  $('kHigh').textContent = high;

  $('meta').textContent = `Linhas: ${filtered.length}`;
  $('meta2').textContent = activeKpi !== 'all' ? `KPI: ${activeKpi}` : '';
}

function updateChips() {
  const chips = [];
  const add = (label, fn) => chips.push({ label, fn });

  const q = $('q').value.trim();
  if (q) add(`Busca: ${q}`, () => { $('q').value = ''; applyFilters(); });

  if ($('tipoOc').value) add(`Tipo Ocorr.: ${$('tipoOc').value}`, () => { $('tipoOc').value = ''; applyFilters(); });
  if ($('impacto').value) add(`Impacto: ${$('impacto').value}`, () => { $('impacto').value = ''; applyFilters(); });
  if ($('statusAt').value) add(`Status: ${$('statusAt').value}`, () => { $('statusAt').value = ''; applyFilters(); });
  if ($('dept').value) add(`Depto: ${$('dept').value}`, () => { $('dept').value = ''; applyFilters(); });
  if ($('onlyOpen').value === '1') add(`Somente abertos`, () => { $('onlyOpen').value = ''; applyFilters(); });

  if (activeKpi !== 'all') {
    const map = { open: 'KPI: Abertos', done: 'KPI: Resolvidos', errop: 'KPI: Erro Operacional', high: 'KPI: Alto/Urgente' };
    add(map[activeKpi] || `KPI: ${activeKpi}`, () => setKpi('all'));
  }

  const c = $('chips');
  c.innerHTML = chips.map((x, i) => `<span class="chip" data-i="${i}">${esc(x.label)} <span class="x">×</span></span>`).join('');
  c.querySelectorAll('.chip').forEach(el => {
    el.addEventListener('click', () => chips[parseInt(el.getAttribute('data-i'), 10)].fn());
  });
}

function badgeStatus(s) {
  const done = isDone(s);
  return done ? `<span class="badge b-done">${esc(s || '-')}</span>` : `<span class="badge b-open">${esc(s || '-')}</span>`;
}
function badgeImpacto(s) {
  const v = String(s || '').toLowerCase();
  if (!s) return `<span class="badge b-low">—</span>`;
  if (v.includes('4') || v.includes('urgente')) return `<span class="badge b-high">${esc(s)}</span>`;
  if (v.includes('3') || v.includes('alta')) return `<span class="badge b-high">${esc(s)}</span>`;
  if (v.includes('2') || v.includes('media') || v.includes('média')) return `<span class="badge b-med">${esc(s)}</span>`;
  if (v.includes('1') || v.includes('baixa')) return `<span class="badge b-low">${esc(s)}</span>`;
  return `<span class="badge b-med">${esc(s)}</span>`;
}

function renderTable() {
  const tb = $('tb');
  if (!filtered.length) {
    tb.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:26px; color:var(--muted);">Nenhum registro para exibir.</td></tr>`;
    return;
  }

  const data = [...filtered].sort((a, b) => {
    const da = parseDate(a.dataAbert)?.getTime() ?? 0;
    const db = parseDate(b.dataAbert)?.getTime() ?? 0;
    return db - da; // mais recente primeiro
  }).slice(0, 1200);

  tb.innerHTML = data.map(r => `
      <tr class="row" onclick="openModal('${escJs(r.ticket)}')">
        <td><span class="tnum">#${esc(r.ticket)}</span></td>
        <td>${esc(fmtDate(r.dataAbert))}</td>
        <td>${esc(r.departamento || '-')}</td>
        <td>${esc(r.responsavel || '-')}</td>
        <td><div class="wrap2" title="${esc(r.tipoOc || '')}">${esc(r.tipoOc || '-')}</div></td>
        <td>${badgeImpacto(r.impacto)}</td>
        <td>${badgeStatus(r.statusAtual)}</td>
        <td><div class="wrap2" title="${esc(r.titulo || '')}">${esc(r.titulo || '-')}</div></td>
        <td><div class="wrap2" title="${esc(r.validacao || '')}">${esc(r.validacao || '-')}</div></td>
      </tr>
    `).join('');
}

function countBy(field) {
  const m = new Map();
  filtered.forEach(r => {
    const k = (r[field] || '').trim() || '(vazio)';
    m.set(k, (m.get(k) || 0) + 1);
  });
  return [...m.entries()].sort((a, b) => b[1] - a[1]);
}
function renderBar(containerId, entries, limit, onClick, kind) {
  const c = $(containerId);
  c.innerHTML = '';
  if (!entries.length) {
    c.innerHTML = `<div style="color:var(--muted); font-size:12px; padding:6px 0;">Sem dados.</div>`;
    return;
  }
  const max = Math.max(...entries.map(e => e[1]));
  entries.slice(0, limit).forEach(([k, v]) => {
    const item = document.createElement('div');
    item.className = 'baritem';
    item.innerHTML = `<div class="k" title="${esc(k)}">${esc(k)}</div>
                        <div class="bar"><div class="fill"></div></div>
                        <div class="v">${v}</div>`;
    const pct = max ? Math.round((v / max) * 100) : 0;
    const fill = item.querySelector('.fill');
    fill.style.width = pct + '%';

    if (kind === 'impacto') {
      const low = String(k).toLowerCase();
      if (low.includes('4') || low.includes('urgente')) fill.style.background = 'linear-gradient(90deg, rgba(239,68,68,.92), rgba(252,165,165,.55))';
      else if (low.includes('3') || low.includes('alta')) fill.style.background = 'linear-gradient(90deg, rgba(239,68,68,.82), rgba(59,130,246,.20))';
      else if (low.includes('2') || low.includes('media') || low.includes('média')) fill.style.background = 'linear-gradient(90deg, rgba(59,130,246,.88), rgba(147,197,253,.55))';
      else fill.style.background = 'linear-gradient(90deg, rgba(148,163,184,.78), rgba(148,163,184,.25))';
    } else if (kind === 'status') {
      const low = String(k).toLowerCase();
      fill.style.background = low.includes('resolv') || low.includes('fech') ? 'linear-gradient(90deg, rgba(22,163,74,.88), rgba(34,197,94,.55))'
        : 'linear-gradient(90deg, rgba(245,158,11,.92), rgba(251,191,36,.55))';
    } else if (kind === 'tipo') {
      const low = String(k).toLowerCase();
      fill.style.background = low.includes('erro operacional') ? 'linear-gradient(90deg, rgba(124,58,237,.88), rgba(37,99,235,.55))'
        : 'linear-gradient(90deg, rgba(37,99,235,.85), rgba(59,130,246,.62))';
    }

    item.addEventListener('click', () => onClick(k));
    c.appendChild(item);
  });
}

function renderCharts() {
  renderBar('chartTipo', countBy('tipoOc'), 12, (k) => {
    $('tipoOc').value = (k === '(vazio)') ? '' : k;
    applyFilters();
  }, 'tipo');

  renderBar('chartImpacto', countBy('impacto'), 10, (k) => {
    $('impacto').value = (k === '(vazio)') ? '' : k;
    applyFilters();
  }, 'impacto');

  renderBar('chartStatus', countBy('statusAtual'), 10, (k) => {
    $('statusAt').value = (k === '(vazio)') ? '' : k;
    applyFilters();
  }, 'status');
}

// Modal
let _cur = null;
function openModal(ticket) {
  const r = all.find(x => String(x.ticket) === String(ticket));
  if (!r) return;
  _cur = r;

  $('mSubject').textContent = `#${r.ticket} — ${r.titulo || 'Sem título'}`;
  $('mMeta').innerHTML = [
    `<span>${badgeStatus(r.statusAtual)}</span>`,
    `<span>•</span>`,
    `<span>${badgeImpacto(r.impacto)}</span>`,
    `<span>•</span>`,
    `<span>${esc(r.tipoOc || '-')}</span>`
  ].join(' ');

  $('mGrid').innerHTML = [
    d('Ticket', '#' + (r.ticket || '-')),
    d('Data Abertura', fmtDate(r.dataAbert)),
    d('Departamento', r.departamento || '-'),
    d('Responsável', r.responsavel || '-'),
    d('Tipo', r.tipo || '-'),
    d('Aliare', r.aliare || '-'),
    d('Prioridade', r.prioridade || '-'),
    d('Impacto', r.impacto || '-'),
    d('Status Atual', r.statusAtual || '-'),
    d('Tipo de Ocorrência', r.tipoOc || '-'),
    d('Validação / Observação TI', r.validacao || '-'),
    d('Título / Descrição', r.titulo || '-')
  ].join('');

  $('mo').classList.add('open');
}
function d(label, value) {
  return `<div class="dcard"><div class="dl">${esc(label)}</div><div class="dv">${esc(String(value ?? '-'))}</div></div>`;
}
function closeModal(e) {
  if (e && e.target && e.target !== $('mo')) return;
  $('mo').classList.remove('open');
}

// Export
function exportFiltered() {
  if (!filtered.length) {
    alert('Não há dados filtrados para exportar.');
    return;
  }
  if (typeof XLSX === 'undefined') {
    alert('XLSX não carregou. Verifique conexão com internet.');
    return;
  }
  const out = filtered.map(r => ({
    Ticket: r.ticket,
    Departamento: r.departamento,
    Responsavel: r.responsavel,
    DataAbertura: r.dataAbert,
    TituloDescricao: r.titulo,
    TipoOcorrencia: r.tipoOc,
    Impacto: r.impacto,
    StatusAtual: r.statusAtual,
    ValidacaoTI: r.validacao,
    Prioridade: r.prioridade,
    Aliare: r.aliare
  }));
  const ws = XLSX.utils.json_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Ocorrencias');
  XLSX.writeFile(wb, `ocorrencias_filtradas_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// Cache
function cacheSave(fileName, data) {
  try { localStorage.setItem('occ_cache_v1', JSON.stringify({ fileName, data, ts: Date.now() })); } catch (_) { }
}
function cacheLoad() {
  try {
    const raw = localStorage.getItem('occ_cache_v1');
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || !obj.data || !Array.isArray(obj.data)) return null;
    return obj;
  } catch (_) { return null; }
}

// Utils
function esc(str) {
  return String(str ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", "&#039;");
}
function escJs(str) {
  return String(str ?? '').replaceAll("\\", "\\\\").replaceAll("'", "\\'");
}

// Boot
(function init() {
  const cache = cacheLoad();
  if (cache && cache.data && cache.data.length) {
    all = cache.data;
    $('fileTag').textContent = `${cache.fileName || 'cache'} • ${all.length} linhas`;
    buildOptions();
    applyFilters();
  }
})();