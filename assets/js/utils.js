// utils.js - Funções utilitárias compartilhadas

/**
 * Seletor simples para elementos por ID
 * @param {string} id - ID do elemento
 * @returns {HTMLElement|null}
 */
function $(id) {
  return document.getElementById(id);
}

/**
 * Normaliza uma string para busca (remove acentos, pontuação, etc.)
 * @param {string} s - String a normalizar
 * @returns {string}
 */
function normKey(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Encontra um valor em uma linha do Excel baseado em candidatos
 * @param {object} row - Linha do Excel
 * @param {string[]} candidates - Possíveis nomes de coluna
 * @returns {string}
 */
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

/**
 * Parse de data, incluindo serial do Excel
 * @param {*} v - Valor a parsear
 * @returns {Date|null}
 */
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
      const ms = epoch.getTime() + serial * 86400000;
      const d = new Date(ms);
      return isNaN(d.getTime()) ? null : d;
    }
  }

  // Tentar parse normal
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Toggle de tema claro/escuro
 * @param {string} appId - ID do elemento raiz
 * @param {string} storageKey - Chave para localStorage
 */
function toggleTheme(appId = 'app', storageKey = 'theme') {
  const app = $(appId);
  const next = app.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  app.setAttribute('data-theme', next);
  try { localStorage.setItem(storageKey, next); } catch (_) { }
}

/**
 * Inicializa tema do localStorage
 * @param {string} appId - ID do elemento raiz
 * @param {string} storageKey - Chave para localStorage
 */
function initTheme(appId = 'app', storageKey = 'theme') {
  try {
    const t = localStorage.getItem(storageKey);
    if (t) $(appId).setAttribute('data-theme', t);
  } catch (_) { }
}

/**
 * Trigger para input file
 * @param {string} inputId - ID do input file
 */
function triggerFile(inputId = 'fileInput') {
  $(inputId).click();
}

// Namespace estável para uso quando algum arquivo sobrescrever nomes globais
try {
  window.DashUtils = window.DashUtils || {};
  window.DashUtils.$ = $;
  window.DashUtils.normKey = normKey;
  window.DashUtils.findValue = findValue;
  window.DashUtils.parseDate = parseDate;
  window.DashUtils.toggleTheme = toggleTheme;
  window.DashUtils.initTheme = initTheme;
  window.DashUtils.triggerFile = triggerFile;
} catch (_) {
  // ignore (ambiente sem window)
}

// Exportar para uso em módulos
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { $, normKey, findValue, parseDate, toggleTheme, initTheme, triggerFile };
}