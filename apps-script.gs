/**
 * EBD Animada 91 anos — recebe os resultados do quiz e grava na planilha.
 * Cole este código em Extensões > Apps Script da planilha.
 *
 * Cada telefone ocupa UMA linha. Se a pessoa jogar de novo, os dados
 * são atualizados e fica valendo a MAIOR pontuação.
 */

const SHEET_NAME = 'Participantes';
const HEADERS = ['ID', 'Nome', 'Sobrenome', 'Telefone', 'Telefone (só números)', 'Perfil',
                 'Pontos', 'Acertos', 'Data/Hora do quiz', 'Jogadas', 'Última atualização'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // evita duas gravações ao mesmo tempo

    const p = JSON.parse(e.postData.contents);
    const digits = String(p.telefone || '').replace(/\D/g, '');
    if (!p.nome || digits.length < 10) return json({ ok: false, erro: 'dados inválidos' });

    const sheet = getSheet();
    const pontos = Number(p.pontos) || 0;
    const acertos = Number(p.acertos) || 0;
    const quando = p.dataHora ? new Date(p.dataHora) : new Date();
    const agora = new Date();

    // procura o telefone na coluna E
    const last = sheet.getLastRow();
    let row = -1;
    if (last > 1) {
      const tels = sheet.getRange(2, 5, last - 1, 1).getValues();
      const i = tels.findIndex(r => String(r[0]) === digits);
      if (i >= 0) row = i + 2;
    }

    if (row === -1) {
      sheet.appendRow([p.id || '', p.nome, p.sobrenome || '', p.telefone || '', digits,
                       p.perfil || '', pontos, acertos, quando, 1, agora]);
    } else {
      const cur = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      const melhor = pontos > Number(cur[6]);
      sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
        cur[0] || p.id || '', p.nome, p.sobrenome || '', p.telefone || '', digits, p.perfil || '',
        melhor ? pontos : cur[6], melhor ? acertos : cur[7], melhor ? quando : cur[8],
        (Number(cur[9]) || 1) + 1, agora
      ]]);
    }
    CacheService.getScriptCache().remove('ranking'); // ranking reflete o novo resultado na hora
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * GET ?action=ranking[&me=81987654321]
 * Ranking público: só nome + inicial do sobrenome, perfil e pontos. NUNCA o telefone.
 * "me" (opcional) marca a linha do próprio participante, sem revelar a de mais ninguém.
 * Sem parâmetros: só confirma que o script está publicado.
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action !== 'ranking') return json({ ok: true, msg: 'Script da EBD funcionando' });

  const me = String(params.me || '').replace(/\D/g, '');
  const rows = getRankingRows(); // [{nome, perfil, pontos, quando, tel}]
  const lista = rows.map(r => ({ nome: r.nome, perfil: r.perfil, pontos: r.pontos, me: !!me && r.tel === me }));
  return json({ ok: true, ranking: lista });
}

function getRankingRows() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ranking');
  if (hit) return JSON.parse(hit);

  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => r[1] && r[4])
    .map(r => ({
      nome: publicName(r[1], r[2]),
      perfil: String(r[5] || ''),
      pontos: Number(r[6]) || 0,
      quando: r[8] instanceof Date ? r[8].getTime() : Date.parse(r[8]) || 0,
      tel: String(r[4])
    }))
    .sort((a, b) => b.pontos - a.pontos || a.quando - b.quando);

  try { cache.put('ranking', JSON.stringify(rows), 15); } catch (err) {} // 15 s de cache
  return rows;
}

function publicName(nome, sobrenome) {
  const ini = String(sobrenome || '').trim().split(' ').pop().charAt(0).toLocaleUpperCase('pt-BR');
  return String(nome).trim() + (ini ? ' ' + ini + '.' : '');
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@'); // telefone como texto
  }
  return sh;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
