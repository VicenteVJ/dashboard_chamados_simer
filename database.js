const path = require('path');
const Database = require('better-sqlite3');

// Em server normal, process.cwd() costuma ser o root do projeto.
// Em Netlify Functions, isso também ajuda a manter o arquivo `database.db` no root.
const DB_PATH = process.env.DB_PATH || path.join(process.cwd(), 'database.db');

// better-sqlite3 cria o arquivo automaticamente se não existir.
const db = new Database(DB_PATH);

db.pragma('journal_mode = WAL');

db.exec(`
  CREATE TABLE IF NOT EXISTS tickets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero TEXT,
    abertoEm TEXT,
    departamento TEXT,
    solicitante TEXT,
    cliente TEXT,
    servico TEXT,
    assunto TEXT,
    responsavel TEXT,
    categoria TEXT,
    ultimaAcao TEXT,
    status TEXT,
    priorizado TEXT
  );

  CREATE TABLE IF NOT EXISTS ocorrencias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ticket TEXT,
    departamento TEXT,
    responsavel TEXT,
    tipo TEXT,
    dataAbert TEXT,
    titulo TEXT,
    aliare TEXT,
    prioridade TEXT,
    impacto TEXT,
    validacao TEXT,
    statusAtual TEXT,
    tipoOc TEXT
  );

  CREATE TABLE IF NOT EXISTS meta (
    dataset TEXT PRIMARY KEY,
    fileName TEXT,
    updatedAt TEXT
  );
`);

const upsertMeta = (dataset, fileName) => {
  db.prepare(`
    INSERT INTO meta(dataset, fileName, updatedAt)
    VALUES (@dataset, @fileName, @updatedAt)
    ON CONFLICT(dataset) DO UPDATE SET
      fileName = excluded.fileName,
      updatedAt = excluded.updatedAt
  `).run({ dataset, fileName: fileName || null, updatedAt: new Date().toISOString() });
};

module.exports = {
  db,
  upsertMeta
};

