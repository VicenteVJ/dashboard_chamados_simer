const { clearOcorrencias } = require('../../lib/pipeline');

const json = (statusCode, payload) => ({
  statusCode,
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(payload)
});

exports.handler = async (event) => {
  try {
    if (event.httpMethod === 'OPTIONS') return json(200, { ok: true });
    clearOcorrencias();
    return json(200, { ok: true });
  } catch (err) {
    console.error(err);
    return json(500, { error: 'Falha ao limpar ocorrências.', detail: err?.message ? String(err.message) : String(err) });
  }
};

