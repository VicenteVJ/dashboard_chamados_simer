const multipart = require('lambda-multipart-parser');
const {
  parseOcorrenciasFromBuffer,
  persistOcorrencias
} = require('../../lib/pipeline');

const json = (statusCode, payload) => ({
  statusCode,
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(payload)
});

exports.handler = async (event) => {
  try {
    if (event.httpMethod === 'OPTIONS') return json(200, { ok: true });

    const form = await multipart.parse(event);

    const files = Array.isArray(form?.files) ? form.files : [];
    const file =
      files.find((f) => f.fieldname === 'file' || f.fieldname === 'files') ||
      files[0];

    const persistRaw = form?.persist ?? form?.fields?.persist ?? '1';
    const shouldPersist = !(persistRaw === '0' || persistRaw === 0 || persistRaw === false || persistRaw === 'false');

    if (!file?.content) {
      return json(400, { error: 'Arquivo não enviado.' });
    }

    const buffer = file.content;
    const originalName = file.filename || file.fileName || 'arquivo.xlsx';

    const data = await parseOcorrenciasFromBuffer(buffer, originalName);

    if (shouldPersist) {
      persistOcorrencias(data, originalName);
    }

    return json(200, { fileName: originalName, count: data.length, data });
  } catch (err) {
    console.error(err);
    return json(400, {
      error: 'Falha ao processar o arquivo.',
      detail: err?.message ? String(err.message) : String(err)
    });
  }
};

