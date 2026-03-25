const { getOcorrencias } = require('../../lib/pipeline');

const json = (statusCode, payload) => ({
  statusCode,
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(payload)
});

exports.handler = async () => {
  try {
    const payload = getOcorrencias();
    return json(200, payload);
  } catch (err) {
    console.error(err);
    return json(500, { error: 'Falha ao carregar ocorrências.', detail: err?.message ? String(err.message) : String(err) });
  }
};

