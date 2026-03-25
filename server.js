const path = require('path');
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { parse: parseCsv } = require('csv-parse/sync');

const app = express();

// Upload em memória (não grava em disco)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 20 * 1024 * 1024 // 20MB
  }
});

const PUBLIC_DIR = __dirname;

app.disable('x-powered-by');
app.use(express.static(PUBLIC_DIR));

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
  const departamento = findValue(row, ['departamento', 'depto', 'setor', 'area', 'área']);
  const solicitante = findValue(row, [
    'usuario solicitante',
    'usuário solicitante',
    'solicitante',
    'requester'
  ]);
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
  let priorizado = findValue(row, [
    'ticket priorizado',
    'priorizado',
    'prioritario',
    'prioritário',
    'priority'
  ]);

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
    priorizado,
    _raw: row
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
    tipoOc,
    _raw: row
  };
};

const parseWorkbookToRows = async (file) => {
  const name = String(file?.originalname || '').toLowerCase();

  if (name.endsWith('.csv')) {
    const txt = file.buffer.toString('utf8');
    return parseCsv(txt, { columns: true, skip_empty_lines: true, bom: true });
  }

  // default: xlsx (tentativa 1: exceljs). Se falhar (ex.: .xls), tenta fallback com xlsx
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(file.buffer);

    const ws = wb.worksheets[0];
    if (!ws) return [];

    // 1ª linha = cabeçalho
    const headerRow = ws.getRow(1);
    const headers = (headerRow?.values || [])
      .slice(1)
      .map((h) => String(h ?? '').trim());

    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const obj = {};

      for (let c = 1; c <= headers.length; c++) {
        const key = headers[c - 1] || `col_${c}`;
        const cell = row.getCell(c);
        // tenta manter valor “humano” (similar ao raw:false do SheetJS)
        const v = cell?.text ?? cell?.value ?? '';
        const s = String(v ?? '').trim();
        if (s !== '') obj[key] = s;
      }

      if (Object.keys(obj).length) rows.push(obj);
    }

    return rows;
  } catch (err) {
    // fallback (compatibilidade) - mantém o comportamento anterior do front
    const wb2 = XLSX.read(file.buffer, { type: 'buffer', cellDates: true, dateNF: 'dd/mm/yyyy' });
    const sheet = wb2.Sheets[wb2.SheetNames[0]];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: 'dd/mm/yyyy' });
  }
};

app.post('/api/parse/tickets', upload.single('file'), (req, res) => {
  (async () => {
    const file = req.file;
    if (!file?.buffer) return res.status(400).json({ error: 'Arquivo não enviado (campo: file).' });

    const rows = await parseWorkbookToRows(file);
    const data = rows
      .map(normalizeTicket)
      .filter((t) => String(t.numero || '').trim() !== '');

    return res.json({ fileName: file.originalname, count: data.length, data });
  })().catch((err) => {
    // eslint-disable-next-line no-console
    console.error(err);
    const ext = String(req.file?.originalname || '').toLowerCase().split('.').pop() || '';
    const hint =
      ext === 'xls'
        ? 'Dica: este arquivo é ".xls". O backend usa exceljs para ".xlsx" e pode não suportar ".xls". Converta para ".xlsx" ou use ".csv".'
        : undefined;
    return res.status(400).json({
      error: 'Falha ao ler o arquivo. Confirme se é um Excel/CSV válido.',
      detail: err?.message ? String(err.message) : String(err),
      hint
    });
  });
});

app.post('/api/parse/ocorrencias', upload.single('file'), (req, res) => {
  (async () => {
    const file = req.file;
    if (!file?.buffer) return res.status(400).json({ error: 'Arquivo não enviado (campo: file).' });

    const rows = await parseWorkbookToRows(file);
    const data = rows
      .map(normalizeOcorrenciaRow)
      .filter((r) => String(r.ticket || '').trim() !== '');

    return res.json({ fileName: file.originalname, count: data.length, data });
  })().catch((err) => {
    // eslint-disable-next-line no-console
    console.error(err);
    const ext = String(req.file?.originalname || '').toLowerCase().split('.').pop() || '';
    const hint =
      ext === 'xls'
        ? 'Dica: este arquivo é ".xls". O backend usa exceljs para ".xlsx" e pode não suportar ".xls". Converta para ".xlsx" ou use ".csv".'
        : undefined;
    return res.status(400).json({
      error: 'Falha ao ler o arquivo. Confirme se é um Excel/CSV válido.',
      detail: err?.message ? String(err.message) : String(err),
      hint
    });
  });
});

// SPA-ish fallback: servir portal
app.get('/', (_req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'index.html'));
});

const PORT = process.env.PORT ? Number(process.env.PORT) : 3000;
app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`Dashboard rodando em http://localhost:${PORT}`);
});

