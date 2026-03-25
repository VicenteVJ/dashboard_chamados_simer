const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { parse: parseCsv } = require('csv-parse/sync');
const { db, upsertMeta } = require('../database');

const normKey = (s) =>
  String(s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');

const findValue = (row, candidates) => {
  const keys = Object.keys(row || {});
  for (const c of candidates) {
    const cn = normKey(c);
    const found = keys.find((k) => normKey(k).includes(cn));
    if (
      found !== undefined &&
      row[found] !== undefined &&
      row[found] !== null &&
      String(row[found]).trim() !== ''
    ) {
      return String(row[found]).trim();
    }
  }
  return '';
};

const normalizeTicket = (row) => {
  const numero = findValue(row, ['numero', 'número', 'ticket', 'id', 'codigo']);
  let numeroFix = numero;
  if (/^\d+\.0+$/.test(numeroFix)) numeroFix = numeroFix.split('.')[0];

  const abertoEm = findValue(row, [
    'aberto em',
    'abertoem',
    'data abertura',
    'data de abertura',
    'criado em',
    'criado'
  ]);
  const departamento = findValue(row, ['departamento', 'depto', 'setor', 'area', 'área', 'área']);
  const solicitante = findValue(row, ['usuario solicitante', 'usuário solicitante', 'solicitante', 'requester']);
  const servico = findValue(row, ['servico', 'serviço', 'service', 'modulo', 'módulo']);
  const assunto = findValue(row, ['assunto', 'titulo', 'título', 'subject', 'descricao', 'descrição']);
  const responsavel = findValue(row, ['responsavel', 'responsável', 'atribuido', 'atribuído', 'assigned']);
  const categoria = findValue(row, ['categoria', 'category', 'tipo']);
  const ultimaAcao = findValue(row, [
    'data da última ação',
    'data da ultima acao',
    'última ação',
    'ultima acao',
    'ultimaacao',
    'atualizado em',
    'updated'
  ]);
  const status = findValue(row, ['status', 'situacao', 'situação', 'estado']);
  const cliente = findValue(row, ['cliente (pessoa)', 'cliente pessoa', 'cliente']);
  let priorizado = findValue(row, ['ticket priorizado', 'priorizado', 'prioritario', 'prioritário', 'priority']);

  priorizado = /sim/i.test(priorizado)
    ? 'Sim'
    : /nao|não/i.test(priorizado)
      ? 'Nao'
      : priorizado
        ? priorizado
        : 'Nao';

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
    priorizado
  };
};

const normalizeOcorrenciaRow = (row) => {
  const ticket = findValue(row, ['ticket', 'chamado', 'id', 'numero', 'número']);
  let ticketFix = ticket;
  if (/^\d+\.0+$/.test(ticketFix)) ticketFix = ticketFix.split('.')[0];

  const departamento = findValue(row, ['departamento', 'depto', 'setor', 'área', 'area']);
  const responsavel = findValue(row, [
    'responsavel',
    'responsável',
    'responsavel ti',
    'owner',
    'atribuido',
    'atribuído'
  ]);
  const tipo = findValue(row, ['tipo']);
  const dataAbert = findValue(row, ['data abert', 'data abertura', 'aberto em', 'abertura']);
  const titulo = findValue(row, [
    'titulo / descricao',
    'título / descrição',
    'titulo',
    'título',
    'descricao',
    'descrição'
  ]);
  const aliare = findValue(row, ['aliare']);
  const prioridade = findValue(row, ['prioridade', 'prioridad']);
  const impacto = findValue(row, ['pimpacto', 'impacto']);
  const validacao = findValue(row, [
    'validacao ti',
    'validação ti',
    'observacao ti',
    'observação ti',
    'validacao',
    'validação'
  ]);
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
    tipoOc
  };
};

const parseWorkbookToRows = async (buffer, originalName) => {
  const name = String(originalName || '').toLowerCase();
  if (name.endsWith('.csv')) {
    const txt = buffer.toString('utf8');
    return parseCsv(txt, { columns: true, skip_empty_lines: true, bom: true });
  }

  // tentativa 1: exceljs
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    if (!ws) return [];

    const headerRow = ws.getRow(1);
    const headers = (headerRow?.values || []).slice(1).map((h) => String(h ?? '').trim());

    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const obj = {};

      for (let c = 1; c <= headers.length; c++) {
        const key = headers[c - 1] || `col_${c}`;
        const cell = row.getCell(c);
        const v = cell?.text ?? cell?.value ?? '';
        const s = String(v ?? '').trim();
        if (s !== '') obj[key] = s;
      }

      if (Object.keys(obj).length) rows.push(obj);
    }

    return rows;
  } catch (_) {
    // fallback: xlsx
    const wb2 = XLSX.read(buffer, { type: 'buffer', cellDates: true, dateNF: 'dd/mm/yyyy' });
    const sheet = wb2.Sheets[wb2.SheetNames[0]];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: 'dd/mm/yyyy' });
  }
};

const persistTickets = (rows, fileName) => {
  const insert = db.prepare(`
    INSERT INTO tickets (
      numero, abertoEm, departamento, solicitante, cliente, servico, assunto,
      responsavel, categoria, ultimaAcao, status, priorizado
    ) VALUES (
      @numero, @abertoEm, @departamento, @solicitante, @cliente, @servico, @assunto,
      @responsavel, @categoria, @ultimaAcao, @status, @priorizado
    )
  `);

  db.transaction(() => {
    db.prepare('DELETE FROM tickets').run();
    rows.forEach((r) => insert.run(r));
  })();

  upsertMeta('tickets', fileName);
};

const persistOcorrencias = (rows, fileName) => {
  const insert = db.prepare(`
    INSERT INTO ocorrencias (
      ticket, departamento, responsavel, tipo, dataAbert, titulo, aliare,
      prioridade, impacto, validacao, statusAtual, tipoOc
    ) VALUES (
      @ticket, @departamento, @responsavel, @tipo, @dataAbert, @titulo, @aliare,
      @prioridade, @impacto, @validacao, @statusAtual, @tipoOc
    )
  `);

  db.transaction(() => {
    db.prepare('DELETE FROM ocorrencias').run();
    rows.forEach((r) => insert.run(r));
  })();

  upsertMeta('ocorrencias', fileName);
};

const getTickets = () => {
  const fileMeta = db.prepare('SELECT fileName FROM meta WHERE dataset = ?').get('tickets');
  const fileName = fileMeta?.fileName || null;
  const data = db
    .prepare('SELECT numero, abertoEm, departamento, solicitante, cliente, servico, assunto, responsavel, categoria, ultimaAcao, status, priorizado FROM tickets')
    .all();
  return { fileName, count: data.length, data };
};

const getOcorrencias = () => {
  const fileMeta = db.prepare('SELECT fileName FROM meta WHERE dataset = ?').get('ocorrencias');
  const fileName = fileMeta?.fileName || null;
  const data = db
    .prepare('SELECT ticket, departamento, responsavel, tipo, dataAbert, titulo, aliare, prioridade, impacto, validacao, statusAtual, tipoOc FROM ocorrencias')
    .all();
  return { fileName, count: data.length, data };
};

const clearTickets = () => {
  db.prepare('DELETE FROM tickets').run();
  upsertMeta('tickets', null);
};

const clearOcorrencias = () => {
  db.prepare('DELETE FROM ocorrencias').run();
  upsertMeta('ocorrencias', null);
};

const parseTicketsFromBuffer = async (buffer, originalName) => {
  const rows = await parseWorkbookToRows(buffer, originalName);
  const data = rows.map(normalizeTicket).filter((t) => String(t.numero || '').trim() !== '');
  return data;
};

const parseOcorrenciasFromBuffer = async (buffer, originalName) => {
  const rows = await parseWorkbookToRows(buffer, originalName);
  const data = rows.map(normalizeOcorrenciaRow).filter((r) => String(r.ticket || '').trim() !== '');
  return data;
};

module.exports = {
  parseTicketsFromBuffer,
  parseOcorrenciasFromBuffer,
  persistTickets,
  persistOcorrencias,
  getTickets,
  getOcorrencias,
  clearTickets,
  clearOcorrencias
};

