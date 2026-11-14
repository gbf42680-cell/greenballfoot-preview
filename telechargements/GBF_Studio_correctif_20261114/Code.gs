/**
 * GreenBallFoot — Simulateur public, lecture seule, 09/11/2026.
 * À déployer dans un projet Apps Script AUTONOME (pas dans un projet PROD existant).
 * Aucune fonction d'écriture. Une seule fonction de données exposée au navigateur : getCataloguePublic().
 */
const GBF_PRODUCTION_ID_ = '__GBF_PRODUCTION_ID_A_REMPLACER__';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Simulateur')
    .setTitle('GBF Studio de Production')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Les seuls champs transmis au navigateur sont les matchs présélectionnés, date/heure, IG et cotes. */
function getCataloguePublic() {
  try {
    // Cache 60 s : limite les lectures quand plusieurs visiteurs ouvrent la page.
    const cache = CacheService.getScriptCache();
    const stored = cache.get('gbf-catalogue-public-v2');
    if (stored) return JSON.parse(stored);
    const ss = SpreadsheetApp.openById(GBF_PRODUCTION_ID_);
    const nationalBook = gbfOddsIndex_(ss.getSheetByName('Combimax2'), 5, 4, 3);
    const europeBook = gbfOddsIndex_(ss.getSheetByName('CombiMax Europe'), 11, 4, 5);
    const national = gbfSelection_(ss.getSheetByName('Choix des matchs'), 'national', nationalBook);
    const europe = gbfSelection_(ss.getSheetByName('Choix des matchs Europe'), 'europe', europeBook);
    const generatedAt = Utilities.formatDate(new Date(), 'Europe/Paris', "dd/MM/yyyy HH:mm:ss");
    const payload = {
      schema: 'gbf-preselection-v1',
      snapshot: 'GreenBallFoot — PRODUCTION · synchro ' + generatedAt + ' (Paris)',
      generatedAt: generatedAt,
      national: national,
      europe: europe
    };
    cache.put('gbf-catalogue-public-v2', JSON.stringify(payload), 60);
    return payload;
  } catch (error) {
    // Pas de nom de fichier interne, ID de classeur, ni trace technique exposés au public.
    throw new Error('La lecture du catalogue GreenBallFoot est momentanément indisponible.');
  }
}

function gbfSelection_(sheet, circuit, book) {
  if (!sheet) throw new Error('Onglet source absent');
  const count = Math.min(Math.max(sheet.getLastRow() - 1, 0), 250);
  if (!count) return [];
  const cols = circuit === 'national' ? 17 : 19;
  const grid = sheet.getRange(2, 1, count, cols).getDisplayValues();
  const dateColumn = circuit === 'national' ? 2 : 3;
  const rawDates = sheet.getRange(2, dateColumn, count, 1).getValues();
  const seen = {};
  const rows = [];
  grid.forEach(function(r, i) {
    const sel = String(r[circuit === 'national' ? 13 : 14] || '').trim().toUpperCase();
    if (sel !== 'ME' && sel !== 'CM' && sel !== 'S') return;
    const name = String(r[0] || '').trim().slice(0, 160);
    if (!name || seen[gbfNorm_(name)]) return;
    const colsOdds = circuit === 'national' ? [14, 15, 16] : [16, 17, 18];
    let odds = colsOdds.map(function(col) { return gbfNum_(r[col]); });
    let oddsSource = 'Choix des matchs';
    // Pour un match CM, les cotes réellement inscrites dans CombiMax priment.
    const bookOdds = sel === 'CM' ? book[gbfNorm_(name)] : null;
    if (bookOdds) { odds = bookOdds.slice(); oddsSource = 'CombiMax'; }
    // Conserver les matchs présélectionnés même si les cotes ne sont pas encore saisies.
    // null signale « cote à renseigner » ; aucune cote fictive n'est transmise.
    odds = odds.map(function(v) { return v > 1 && v < 1000 ? v : null; });
    const rawDate = rawDates[i] ? rawDates[i][0] : '';
    const date = gbfDate_(rawDate, r[dateColumn - 1]);
    const hour = String(r[circuit === 'national' ? 2 : 3] || '').trim().slice(0, 8);
    const ig = gbfNum_(r[12]);
    const key = gbfNorm_(name);
    seen[key] = true;
    rows.push({
      id: (circuit === 'national' ? 'n-' : 'e-') + gbfHash_(key),
      name: name, date: date, hour: hour,
      ig: isFinite(ig) ? Math.max(0, ig) : 0,
      selection: sel, odds: odds,
      europe: String(r[circuit === 'national' ? 3 : 1] || '').slice(0, 30),
      oddsSource: oddsSource
    });
  });
  return rows.slice(0, 100);
}

function gbfOddsIndex_(sheet, rowStart, nRows, matchCol) {
  if (!sheet) throw new Error('Onglet CombiMax absent');
  const rows = sheet.getRange(rowStart, matchCol, nRows, 4).getDisplayValues();
  const oddsMap = {};
  rows.forEach(function(r) {
    const name = gbfNorm_(r[0]);
    const odds = r.slice(1, 4).map(gbfNum_);
    if (name && odds.every(function(x) { return x > 1 && x < 1000; })) oddsMap[name] = odds;
  });
  return oddsMap;
}

function gbfNum_(v) {
  const match = String(v === null || v === undefined ? '' : v).replace(/\s/g, '').replace(',', '.').match(/[0-9]+(?:\.[0-9]+)?/);
  return match ? Number(match[0]) : NaN;
}
function gbfNorm_(text) { return String(text || '').normalize('NFKC').replace(/[\u2012\u2013\u2014\-]/g, '-').replace(/\s+/g, ' ').trim().toLowerCase(); }
function gbfHash_(s) {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  return (h >>> 0).toString(36);
}
function gbfDate_(v, displayed) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()))
    return Utilities.formatDate(v, 'Europe/Paris', 'dd/MM/yyyy');
  const t = String(displayed || v || '').trim();
  let m = t.match(/^(\d{1,2})[\/-](\d{1,2})(?:[\/-](\d{2,4}))?/);
  if (!m) return '';
  let year = m[3] ? (m[3].length === 2 ? '20' + m[3] : m[3]) : String(new Date().getFullYear());
  return m[1].padStart(2, '0') + '/' + m[2].padStart(2, '0') + '/' + year;
}