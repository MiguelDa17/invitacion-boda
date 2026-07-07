const GUESTS_SPREADSHEET_ID = '1thqdVGV7PtosBbt0G6wAi_F444elK34ZnUPOq27Oyqk';
const GUESTS_SHEET_NAME = 'Respuestas de formulario 1';

const RESPONSES_SPREADSHEET_ID = '1vJC8NBQlSDARAAA7gPBO27FbHTlt4d1egrXqUc08lpU';
const RESPONSES_SHEET_NAME = 'Hoja 1';

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.action === 'searchGuest') {
    const payload = {
      ok: true,
      matches: searchGuestGroups(params.q || '')
    };

    return outputJSON(payload, params.callback);
  }

  return outputJSON({ ok: false, error: 'Accion no soportada' }, params.callback);
}

function doPost(e) {
  try {
    const data = parsePostData(e);
    const sheet = getOrCreateSheet(RESPONSES_SHEET_NAME);
    ensureHeaders(sheet, [
      'fecha_respuesta',
      'grupo_id',
      'nombres_invitados',
      'cantidad_invitados',
      'nombre_buscado',
      'nombre',
      'invitados',
      'confirmacion',
      'asistencia'
    ]);

    sheet.appendRow([
      new Date(),
      data.grupo_id || '',
      data.nombres_invitados || '',
      data.cantidad_invitados || data.invitados || '',
      data.nombre_buscado || '',
      data.nombre || '',
      data.invitados || data.cantidad_invitados || '',
      data.confirmacion || data.asistencia || '',
      data.asistencia || data.confirmacion || ''
    ]);

    return outputJSON({ ok: true });
  } catch (err) {
    return outputJSON({ ok: false, error: err.message });
  }
}

function searchGuestGroups(query) {
  const normalizedQuery = normalizeText(query);
  if (normalizedQuery.length < 3) return [];

  const sheet = getGuestsSpreadsheet().getSheetByName(GUESTS_SHEET_NAME);
  if (!sheet) throw new Error('No existe la pestaña ' + GUESTS_SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(String);
  const indexes = {
    grupoId: findHeaderIndex(headers, 'Grupo_ID'),
    names: findHeaderIndex(headers, 'Nombres completos invitados'),
    quantity: findHeaderIndex(headers, 'Cantidad'),
    status: findHeaderIndex(headers, 'Estado')
  };

  if (indexes.names === -1) {
    throw new Error('Falta la columna requerida "Nombres completos invitados" en ' + GUESTS_SHEET_NAME);
  }

  const queryTokens = tokenize(normalizedQuery);

  return values.slice(1)
    .map((row, rowIndex) => buildGuestGroup(row, indexes, rowIndex + 2))
    .filter(group => group && (!group.status || normalizeText(group.status) === 'activo'))
    .map(group => {
      const match = scoreGuestGroup(group, normalizedQuery, queryTokens);

      return {
        group: {
          grupoId: group.grupoId,
          names: group.names,
          quantity: group.quantity,
          score: match.score,
          matchReason: match.matchReason
        },
        score: match.score
      };
    })
    .filter(result => result.score >= 70)
    .sort((a, b) => b.score - a.score)
    .slice(0, 5)
    .map(result => result.group);
}

function buildGuestGroup(row, indexes, sheetRowNumber) {
  const namesCell = String(row[indexes.names] || '').trim();
  if (!namesCell) return null;

  const names = namesCell
    .split(/\r?\n/)
    .map(name => name.trim())
    .filter(Boolean);

  if (!names.length) return null;

  const rawQuantity = indexes.quantity >= 0 ? Number(row[indexes.quantity]) : 0;

  return {
    grupoId: indexes.grupoId >= 0 ? String(row[indexes.grupoId] || '').trim() : 'FILA_' + sheetRowNumber,
    names: names,
    quantity: rawQuantity > 0 ? rawQuantity : names.length,
    status: indexes.status >= 0 ? String(row[indexes.status] || '').trim() : ''
  };
}

function scoreGuestGroup(group, normalizedQuery, queryTokens) {
  if (!queryTokens.length) return noMatch();

  return group.names
    .map(name => scoreSingleName(normalizeText(name), normalizedQuery, queryTokens))
    .sort((a, b) => b.score - a.score)[0] || noMatch();
}

function scoreSingleName(normalizedName, normalizedQuery, queryTokens) {
  const nameTokens = tokenize(normalizedName);

  if (normalizedName.includes(normalizedQuery)) {
    return { score: 100, matchReason: 'phrase_exact_same_guest' };
  }

  if (queryTokens.length >= 2) {
    if (queryTokens.every(token => tokenMatchesExact(token, nameTokens))) {
      return { score: 90, matchReason: 'all_tokens_same_guest' };
    }

    if (queryTokens.every(token => tokenMatchesPrefix(token, nameTokens))) {
      return { score: 80, matchReason: 'all_tokens_prefix_same_guest' };
    }

    if (queryTokens.every(token => tokenMatchesFuzzy(token, nameTokens))) {
      return { score: 70, matchReason: 'all_tokens_fuzzy_same_guest' };
    }

    return noMatch();
  }

  const token = queryTokens[0];

  if (nameTokens.some(nameToken => nameToken === token)) {
    return { score: 95, matchReason: 'single_token_exact' };
  }

  if (nameTokens.some(nameToken => nameToken.startsWith(token))) {
    return { score: 85, matchReason: 'single_token_prefix' };
  }

  if (token.length >= 5 && nameTokens.some(nameToken => nameToken.includes(token))) {
    return { score: 78, matchReason: 'single_token_partial' };
  }

  if (tokenMatchesFuzzy(token, nameTokens)) {
    return { score: 70, matchReason: 'single_token_fuzzy' };
  }

  return noMatch();
}

function tokenize(text) {
  return normalizeText(text)
    .split(' ')
    .filter(token => token.length >= 2);
}

function tokenMatchesExact(token, nameTokens) {
  return nameTokens.some(nameToken => nameToken === token || nameToken.includes(token));
}

function tokenMatchesPrefix(token, nameTokens) {
  return nameTokens.some(nameToken => nameToken.startsWith(token));
}

function tokenMatchesFuzzy(token, nameTokens) {
  return token.length >= 5 && nameTokens.some(nameToken => isSmallTypo(token, nameToken));
}

function noMatch() {
  return { score: 0, matchReason: 'no_match' };
}

function isSmallTypo(queryToken, nameToken) {
  if (queryToken.length < 5) return false;
  if (Math.abs(queryToken.length - nameToken.length) > 2) return false;
  const maxDistance = queryToken.length <= 6 ? 1 : 2;
  return levenshteinDistance(queryToken, nameToken) <= maxDistance;
}

function levenshteinDistance(a, b) {
  const matrix = [];

  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

function parsePostData(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      // Sigue con e.parameter por compatibilidad con formularios tradicionales.
    }
  }

  return (e && e.parameter) || {};
}

function getGuestsSpreadsheet() {
  return SpreadsheetApp.openById(GUESTS_SPREADSHEET_ID);
}

function getResponsesSpreadsheet() {
  return SpreadsheetApp.openById(RESPONSES_SPREADSHEET_ID);
}

function getOrCreateSheet(name) {
  const spreadsheet = getResponsesSpreadsheet();
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureHeaders(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = current.some(Boolean);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function findHeaderIndex(headers, expectedHeader) {
  const expected = normalizeText(expectedHeader);
  return headers.findIndex(header => normalizeText(header) === expected);
}

function normalizeText(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9ñ\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function outputJSON(payload, callback) {
  if (callback) {
    const safeCallback = String(callback).match(/^[\w.$]+$/) ? callback : 'callback';
    return ContentService
      .createTextOutput(safeCallback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
