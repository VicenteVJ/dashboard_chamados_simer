const busboy = require('busboy');
const { normalizeOcorrenciaRow, parseWorkbookToRows } = require('./parse-shared');

const parseMultipart = (event) =>
  new Promise((resolve, reject) => {
    const contentType =
      event.headers['content-type'] || event.headers['Content-Type'] || '';
    const bb = busboy({ headers: { 'content-type': contentType } });

    let fileBuffer = null;
    let fileName = '';

    bb.on('file', (_fieldname, stream, info) => {
      fileName = info.filename || '';
      const chunks = [];
      stream.on('data', (chunk) => chunks.push(chunk));
      stream.on('end', () => {
        fileBuffer = Buffer.concat(chunks);
      });
    });

    bb.on('finish', () => resolve({ buffer: fileBuffer, originalname: fileName }));
    bb.on('error', reject);

    const encoding = event.isBase64Encoded ? 'base64' : 'binary';
    bb.end(Buffer.from(event.body, encoding));
  });

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  try {
    const file = await parseMultipart(event);
    if (!file.buffer) {
      return {
        statusCode: 400,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ error: 'Arquivo não enviado (campo: file).' })
      };
    }

    const rows = await parseWorkbookToRows(file.buffer, file.originalname);
    const data = rows.map(normalizeOcorrenciaRow).filter((r) => String(r.ticket || '').trim() !== '');

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileName: file.originalname, count: data.length, data })
    };
  } catch (err) {
    console.error(err);
    const ext = String(event.headers?.['x-filename'] || '').toLowerCase().split('.').pop() || '';
    const hint =
      ext === 'xls'
        ? 'Dica: este arquivo é ".xls". Converta para ".xlsx" ou use ".csv".'
        : undefined;
    return {
      statusCode: 400,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        error: 'Falha ao ler o arquivo. Confirme se é um Excel/CSV válido.',
        detail: err?.message ? String(err.message) : String(err),
        hint
      })
    };
  }
};
