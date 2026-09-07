const GUESTS_SPREADSHEET_ID = '1thqdVGV7PtosBbt0G6wAi_F444elK34ZnUPOq27Oyqk';
const GUESTS_SHEET_NAME = 'Respuestas de formulario 1';

const RESPONSES_SPREADSHEET_ID = '1vJC8NBQlSDARAAA7gPBO27FbHTlt4d1egrXqUc08lpU';
const RESPONSES_SHEET_NAME = 'Hoja 1';
const SUMMARY_SHEET_NAME = 'Resumen RSVP';
const CONFIRMED_SHEET_NAME = 'Confirmados';
const DECLINED_SHEET_NAME = 'No asisten';
const PENDING_SHEET_NAME = 'Pendientes';
const TOTALS_SHEET_NAME = 'Totales RSVP';
const RSVP_APP_VERSION = '2026-09-06-rsvp-person-status-v3';
let guestGroupsCache = null;

const RESPONSE_HEADERS = [
  'fecha_respuesta',
  'grupo_id',
  'nombres_invitados',
  'cantidad_invitados',
  'nombre_buscado',
  'nombre',
  'invitados',
  'confirmacion',
  'asistencia',
  'confirmaciones_invitados',
  'asistentes_confirmados',
  'no_asisten',
  'total_confirmados',
  'estado_grupo',
  'cantidad_asisten',
  'cantidad_no_asisten',
  'pendientes_invitados',
  'cantidad_pendientes'
];

const SUMMARY_HEADERS = [
  'grupo_id',
  'nombres_invitados',
  'cantidad_invitados',
  'estado_grupo',
  'asistentes_confirmados',
  'cantidad_asisten',
  'no_asisten',
  'cantidad_no_asisten',
  'pendientes_invitados',
  'cantidad_pendientes',
  'nombre_buscado',
  'fecha_respuesta'
];

const STATUS_HEADERS = [
  'nombre_invitado',
  'grupo_id',
  'grupo_invitacion',
  'cantidad_grupo',
  'estado',
  'nombre_buscado',
  'fecha_respuesta'
];

function doGet(e) {
  resetExecutionCache();
  const params = (e && e.parameter) || {};

  if (params.action === 'health') {
    return outputJSON({
      ok: true,
      version: RSVP_APP_VERSION,
      guestsSpreadsheetId: GUESTS_SPREADSHEET_ID,
      guestsSheetName: GUESTS_SHEET_NAME,
      responsesSpreadsheetId: RESPONSES_SPREADSHEET_ID,
      responsesSheetName: RESPONSES_SHEET_NAME
    }, params.callback);
  }

  if (params.action === 'searchGuest') {
    const searchResult = searchGuestGroupsResult(params.q || '');
    const payload = {
      ok: true,
      version: RSVP_APP_VERSION,
      matches: searchResult.matches,
      closedMatches: searchResult.closedMatches,
      allMatchesClosed: searchResult.allMatchesClosed
    };

    return outputJSON(payload, params.callback);
  }

  if (params.action === 'submitRsvp') {
    try {
      return outputJSON(saveRsvpResponse(parseSubmitData(params)), params.callback);
    } catch (err) {
      return outputJSON({ ok: false, error: err.message }, params.callback);
    }
  }

  return outputJSON({ ok: false, error: 'Accion no soportada' }, params.callback);
}

function doPost(e) {
  resetExecutionCache();
  return outputJSON(saveRsvpResponse(parsePostData(e)));
}

function saveRsvpResponse(data) {
  let lock = null;

  try {
    if (!hasRsvpPayload(data)) {
      return { ok: false, error: 'Solicitud RSVP vacía' };
    }

    if (typeof LockService !== 'undefined') {
      lock = LockService.getScriptLock();
      lock.waitLock(10000);
    }

    const sheet = prepareResponsesSheet({ fast: true });
    const accepted = buildAcceptedResponseRecord(data);
    if (!accepted.ok) return accepted;

    appendRecordByHeaders(sheet, accepted.record);

    const totals = refreshRsvpSummary({ fast: true });

    return {
      ok: true,
      summaryUpdated: true,
      acceptedGuests: accepted.acceptedGuests,
      ignoredGuests: accepted.ignoredGuests,
      totals: totals
    };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function setupRsvpWorkbook() {
  repairLegacyRsvpResponses();
}

function validateRsvpWorkbook() {
  const guestsSheet = getGuestsSpreadsheet().getSheetByName(GUESTS_SHEET_NAME);
  if (!guestsSheet) throw new Error('No existe la pestaña de invitados ' + GUESTS_SHEET_NAME);

  const groups = getGuestGroups();
  const totals = repairLegacyRsvpResponses();

  Logger.log('Validación OK');
  Logger.log('Grupos de invitados: ' + groups.length);
  Logger.log('Histórico de respuestas: ' + RESPONSES_SHEET_NAME);
  Logger.log('Resumen organizado: ' + SUMMARY_SHEET_NAME);
  Logger.log('Pestañas de seguimiento: ' + CONFIRMED_SHEET_NAME + ', ' + DECLINED_SHEET_NAME + ', ' + PENDING_SHEET_NAME);
  Logger.log('Totales: ' + TOTALS_SHEET_NAME);
  Logger.log('Personas que van: ' + totals.personasQueVan);
  Logger.log('Personas que no van: ' + totals.personasQueNoVan);
  Logger.log('Personas pendientes: ' + totals.personasPendientes);
}

function repairLegacyRsvpResponses() {
  resetExecutionCache();
  const sheet = getOrCreateSheet(RESPONSES_SHEET_NAME);
  normalizeLegacyResponseSheet(sheet);
  ensureHeaders(sheet, RESPONSE_HEADERS);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const records = values.slice(1)
    .map(row => rowToRecord(headers, row))
    .filter(hasMeaningfulHistoricRecord)
    .map(record => buildResponseRecord(record));

  rewriteSheetWithRecords(sheet, RESPONSE_HEADERS, records, { tabColor: '#6E1F2C' });
  const totals = refreshRsvpSummary();

  Logger.log('Histórico RSVP reparado');
  Logger.log('Filas históricas conservadas: ' + records.length);
  Logger.log('Personas que van: ' + totals.personasQueVan);
  Logger.log('Personas que no van: ' + totals.personasQueNoVan);
  Logger.log('Personas pendientes: ' + totals.personasPendientes);

  return totals;
}

function actualizarRsvpHistorico() {
  return repairLegacyRsvpResponses();
}

function searchGuestGroups(query) {
  return searchGuestGroupsResult(query).matches;
}

function resetExecutionCache() {
  guestGroupsCache = null;
}

function searchGuestGroupsResult(query) {
  const normalizedQuery = normalizeText(query);
  if (normalizedQuery.length < 3) {
    return { matches: [], closedMatches: 0, allMatchesClosed: false };
  }

  const queryTokens = tokenize(normalizedQuery);
  const responseStatesByGroup = getResponseStatesByGroup();
  const openMatches = [];
  let closedMatches = 0;

  getGuestGroups()
    .filter(group => group && (!group.status || normalizeText(group.status) === 'activo'))
    .forEach(group => {
      const searchGroup = buildSearchGuestGroup(group, responseStatesByGroup[group.grupoId]);
      const pendingMatch = scoreNames(searchGroup.pendingNames, normalizedQuery, queryTokens);
      const anyMatch = scoreNames(group.names, normalizedQuery, queryTokens);

      if (pendingMatch.score >= 70) {
        openMatches.push({
          group: Object.assign({}, searchGroup, {
            score: pendingMatch.score,
            matchReason: pendingMatch.matchReason
          }),
          score: pendingMatch.score
        });
        return;
      }

      if (anyMatch.score >= 70) {
        closedMatches++;
      }
    });

  openMatches.sort((a, b) => b.score - a.score);

  return {
    matches: openMatches
      .slice(0, 5)
      .map(result => result.group),
    closedMatches: closedMatches,
    allMatchesClosed: closedMatches > 0 && openMatches.length === 0
  };
}

function getClosedResponseGroupIds() {
  const latestResponses = getResponseStatesByGroup();
  const closed = {};

  Object.keys(latestResponses).forEach(groupId => {
    const state = latestResponses[groupId];
    const answered = Object.keys(state.byName || {}).length;
    const pending = Math.max(0, Number(state.quantity || 0) - answered);

    if (answered > 0 && pending === 0) {
      closed[groupId] = true;
    }
  });

  return closed;
}

function getGuestGroups() {
  if (guestGroupsCache) return guestGroupsCache;

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

  guestGroupsCache = values.slice(1)
    .map((row, rowIndex) => buildGuestGroup(row, indexes, rowIndex + 2))
    .filter(Boolean);

  return guestGroupsCache;
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

function buildResponseRecord(data) {
  const officialGroup = getOfficialGroupForRecord(data);
  const invitedNames = officialGroup
    ? officialGroup.names
    : splitNames(data.nombres_invitados || data.nombre);
  const quantity = toPositiveNumber(
    officialGroup ? officialGroup.quantity : data.cantidad_invitados || data.invitados,
    invitedNames.length
  );
  const groupId = officialGroup ? officialGroup.grupoId : data.grupo_id || '';
  const normalizedData = Object.assign({}, data, {
    grupo_id: groupId,
    nombres_invitados: invitedNames.join('\n'),
    cantidad_invitados: quantity
  });
  const confirmations = getRecordConfirmations(normalizedData);

  let yesNames = splitNames(data.asistentes_confirmados);
  let noNames = splitNames(data.no_asisten);

  if (confirmations.length) {
    yesNames = confirmations
      .filter(item => item.asistencia === 'Sí')
      .map(item => item.nombre)
      .filter(Boolean);
    noNames = confirmations
      .filter(item => item.asistencia === 'No')
      .map(item => item.nombre)
      .filter(Boolean);
  }

  yesNames = uniqueNames(yesNames);
  noNames = uniqueNames(noNames);

  let yesCount = yesNames.length;
  if (!yesCount && data.total_confirmados !== undefined) {
    yesCount = toPositiveNumber(data.total_confirmados, 0);
  }

  const noCount = noNames.length;
  let pendingNames = splitNames(data.pendientes_invitados);
  if (!pendingNames.length) {
    pendingNames = computePendingNames(invitedNames, yesNames, noNames);
  }

  const computedPendingCount = Math.max(0, quantity - yesCount - noCount);
  const pendingCount = pendingNames.length || computedPendingCount;
  const groupStatus = getGroupStatus(quantity, yesCount, noCount, pendingCount);
  const attendance = getAttendanceStatus(quantity, yesCount, noCount, pendingCount);
  const serializedConfirmations = JSON.stringify(
    yesNames
      .map(name => ({ nombre: name, asistencia: 'Sí' }))
      .concat(noNames.map(name => ({ nombre: name, asistencia: 'No' })))
  );

  return {
    fecha_respuesta: data.fecha_respuesta || new Date(),
    grupo_id: groupId,
    nombres_invitados: invitedNames.join('\n'),
    cantidad_invitados: quantity,
    nombre_buscado: data.nombre_buscado || '',
    nombre: data.nombre || invitedNames[0] || '',
    invitados: data.invitados || yesCount,
    confirmacion: data.confirmacion || attendance,
    asistencia: attendance,
    confirmaciones_invitados: serializedConfirmations,
    asistentes_confirmados: yesNames.join('\n'),
    no_asisten: noNames.join('\n'),
    total_confirmados: yesCount,
    estado_grupo: groupStatus,
    cantidad_asisten: yesCount,
    cantidad_no_asisten: noCount,
    pendientes_invitados: pendingNames.join('\n'),
    cantidad_pendientes: pendingCount
  };
}

function buildAcceptedResponseRecord(data) {
  const groupId = String(data.grupo_id || '').trim();
  const officialGroup = findGuestGroupById(groupId);
  const invitedNames = officialGroup
    ? officialGroup.names
    : splitNames(data.nombres_invitados || data.nombre);
  const quantity = toPositiveNumber(
    officialGroup ? officialGroup.quantity : data.cantidad_invitados || data.invitados,
    invitedNames.length
  );
  const incomingConfirmations = getIncomingConfirmations(data, invitedNames);

  if (!groupId) {
    return { ok: false, error: 'No se recibió el ID del grupo de invitación.' };
  }

  if (!officialGroup) {
    return { ok: false, error: 'No encontramos este grupo en la lista oficial de invitados. Comunícate con los novios.' };
  }

  if (!incomingConfirmations.length) {
    return { ok: false, error: 'Selecciona Sí o No para al menos una persona pendiente.' };
  }

  const responseState = getResponseStatesByGroup()[groupId] || null;
  const statusMap = getGroupStatusMap(
    { names: invitedNames, quantity: quantity },
    responseState
  );
  const acceptedConfirmations = [];
  const ignoredGuests = [];

  incomingConfirmations.forEach(item => {
    const matchedName = matchKnownGuestName(item.nombre, invitedNames) || item.nombre;
    const key = normalizeText(matchedName);

    if (statusMap[key]) {
      ignoredGuests.push(matchedName);
      return;
    }

    acceptedConfirmations.push({
      nombre: matchedName,
      asistencia: item.asistencia
    });
    statusMap[key] = item.asistencia;
  });

  if (!acceptedConfirmations.length) {
    return {
      ok: false,
      error: 'Esta persona ya tiene respuesta registrada. Si necesitas hacer algún cambio, comunícate con los novios.',
      ignoredGuests: ignoredGuests
    };
  }

  const yesNames = invitedNames.filter(name => statusMap[normalizeText(name)] === 'Sí');
  const noNames = invitedNames.filter(name => statusMap[normalizeText(name)] === 'No');
  const pendingNames = invitedNames.filter(name => !statusMap[normalizeText(name)]);
  const pendingCount = Math.max(0, quantity - yesNames.length - noNames.length);
  const acceptedYesNames = acceptedConfirmations
    .filter(item => item.asistencia === 'Sí')
    .map(item => item.nombre);
  const acceptedNoNames = acceptedConfirmations
    .filter(item => item.asistencia === 'No')
    .map(item => item.nombre);
  const currentConfirmations = invitedNames
    .filter(name => statusMap[normalizeText(name)])
    .map(name => ({
      nombre: name,
      asistencia: statusMap[normalizeText(name)]
    }));

  const acceptedData = Object.assign({}, data, {
    nombres_invitados: invitedNames.join('\n'),
    cantidad_invitados: quantity,
    confirmaciones_invitados: JSON.stringify(currentConfirmations),
    asistentes_confirmados: yesNames.join('\n'),
    no_asisten: noNames.join('\n'),
    total_confirmados: yesNames.length,
    invitados: acceptedYesNames.length,
    confirmacion: getAttendanceStatus(acceptedConfirmations.length, acceptedYesNames.length, acceptedNoNames.length, 0),
    asistencia: getGroupStatus(quantity, yesNames.length, noNames.length, pendingCount),
    pendientes_invitados: pendingNames.join('\n'),
    cantidad_pendientes: pendingCount,
    estado_grupo: getGroupStatus(quantity, yesNames.length, noNames.length, pendingCount)
  });

  return {
    ok: true,
    record: buildResponseRecord(acceptedData),
    acceptedGuests: acceptedConfirmations.map(item => item.nombre),
    ignoredGuests: ignoredGuests
  };
}

function findGuestGroupById(groupId) {
  if (!groupId) return null;
  return getGuestGroups().find(group => group.grupoId === groupId) || null;
}

function getIncomingConfirmations(data, invitedNames) {
  let confirmations = parseGuestConfirmations(data.confirmaciones_invitados);

  if (!confirmations.length) {
    confirmations = splitNames(data.asistentes_confirmados)
      .map(name => ({ nombre: name, asistencia: 'Sí' }))
      .concat(splitNames(data.no_asisten).map(name => ({ nombre: name, asistencia: 'No' })));
  }

  if (!confirmations.length) {
    const attendance = normalizeAttendance(data.asistencia || data.confirmacion);
    if (attendance) {
      confirmations = invitedNames.map(name => ({ nombre: name, asistencia: attendance }));
    }
  }

  return confirmations
    .map(item => ({
      nombre: matchKnownGuestName(item.nombre, invitedNames) || item.nombre,
      asistencia: normalizeAttendance(item.asistencia)
    }))
    .filter(item => item.nombre && item.asistencia);
}

function refreshRsvpSummary(options) {
  const settings = options || {};
  const summarySheet = getOrCreateSheet(SUMMARY_SHEET_NAME);
  const responseStatesByGroup = getResponseStatesByGroup();
  const records = getGuestGroups()
    .filter(group => group && (!group.status || normalizeText(group.status) === 'activo'))
    .map(group => buildSummaryRecord(group, responseStatesByGroup[group.grupoId]));

  const totals = calculateRsvpTotals(records);

  rewriteSheetWithRecords(summarySheet, SUMMARY_HEADERS, records, { tabColor: '#6E1F2C', fast: settings.fast });
  refreshStatusSheets(records, settings);
  refreshTotalsSheet(totals, settings);

  return totals;
}

function calculateRsvpTotals(records) {
  const totalInvitados = records.reduce((sum, record) => sum + Number(record.cantidad_invitados || 0), 0);
  const personasQueVan = records.reduce((sum, record) => sum + Number(record.cantidad_asisten || 0), 0);
  const personasQueNoVan = records.reduce((sum, record) => sum + Number(record.cantidad_no_asisten || 0), 0);
  const personasPendientes = records.reduce((sum, record) => sum + Number(record.cantidad_pendientes || 0), 0);
  const invitacionesCompletas = records.filter(record => {
    const answered = Number(record.cantidad_asisten || 0) + Number(record.cantidad_no_asisten || 0);
    return answered > 0 && Number(record.cantidad_pendientes || 0) === 0;
  }).length;
  const invitacionesParciales = records.filter(record => {
    const answered = Number(record.cantidad_asisten || 0) + Number(record.cantidad_no_asisten || 0);
    return answered > 0 && Number(record.cantidad_pendientes || 0) > 0;
  }).length;
  const invitacionesSinRespuesta = records.filter(record => {
    const answered = Number(record.cantidad_asisten || 0) + Number(record.cantidad_no_asisten || 0);
    return answered === 0 && Number(record.cantidad_pendientes || 0) > 0;
  }).length;

  return {
    totalInvitados: totalInvitados,
    personasQueVan: personasQueVan,
    personasQueNoVan: personasQueNoVan,
    personasPendientes: personasPendientes,
    totalGrupos: records.length,
    invitacionesCompletas: invitacionesCompletas,
    invitacionesParciales: invitacionesParciales,
    invitacionesSinRespuesta: invitacionesSinRespuesta
  };
}

function refreshTotalsSheet(totals, options) {
  const settings = options || {};
  const sheet = getOrCreateSheet(TOTALS_SHEET_NAME);
  const shouldFormat = !settings.fast || !sheet.getLastRow() || !sheet.getLastColumn();

  if (!shouldFormat) {
    resetSheetContents(sheet);
  } else {
    resetSheet(sheet);
  }

  const rows = [
    ['Métrica', 'Total', 'Detalle'],
    ['Invitados totales', totals.totalInvitados, 'Cupos registrados en la base de invitados'],
    ['Personas que van', totals.personasQueVan, 'Confirmaron Sí'],
    ['Personas que no van', totals.personasQueNoVan, 'Confirmaron No'],
    ['Personas pendientes', totals.personasPendientes, 'Aún falta respuesta'],
    ['', '', ''],
    ['Invitaciones totales', totals.totalGrupos, 'Grupos registrados en la base'],
    ['Invitaciones completas', totals.invitacionesCompletas, 'Todos los invitados del grupo respondieron'],
    ['Invitaciones parciales', totals.invitacionesParciales, 'Respondió una parte del grupo'],
    ['Invitaciones sin respuesta', totals.invitacionesSinRespuesta, 'Nadie del grupo ha respondido']
  ];

  if (shouldFormat) {
    sheet.getRange(1, 1, 1, 3).merge();
  }
  sheet.getRange(1, 1).setValue('Totales RSVP');
  sheet.getRange(2, 1).setValue('Última actualización');
  sheet.getRange(2, 2).setValue(new Date());
  sheet.getRange(4, 1, rows.length, 3).setValues(rows);

  if (shouldFormat) {
    formatTotalsSheet(sheet);
  }
}

function refreshStatusSheets(summaryRecords, options) {
  const settings = options || {};

  const grouped = buildStatusSheetRecords(summaryRecords);

  rewriteSheetWithRecords(
    getOrCreateSheet(CONFIRMED_SHEET_NAME),
    STATUS_HEADERS,
    grouped.confirmados,
    { tabColor: '#2F6B4F', headerColor: '#2F6B4F', fast: settings.fast }
  );
  rewriteSheetWithRecords(
    getOrCreateSheet(DECLINED_SHEET_NAME),
    STATUS_HEADERS,
    grouped.noAsisten,
    { tabColor: '#5E2E37', headerColor: '#5E2E37', fast: settings.fast }
  );
  rewriteSheetWithRecords(
    getOrCreateSheet(PENDING_SHEET_NAME),
    STATUS_HEADERS,
    grouped.pendientes,
    { tabColor: '#A06A2A', headerColor: '#A06A2A', fast: settings.fast }
  );
}

function buildStatusSheetRecords(summaryRecords) {
  const grouped = {
    confirmados: [],
    noAsisten: [],
    pendientes: []
  };

  summaryRecords.forEach(record => {
    splitNames(record.asistentes_confirmados).forEach(name => {
      grouped.confirmados.push(buildStatusRecord(record, name, 'Confirmado'));
    });

    splitNames(record.no_asisten).forEach(name => {
      grouped.noAsisten.push(buildStatusRecord(record, name, 'No asiste'));
    });

    splitNames(record.pendientes_invitados).forEach(name => {
      grouped.pendientes.push(buildStatusRecord(record, name, 'Pendiente'));
    });
  });

  return grouped;
}

function buildStatusRecord(summaryRecord, name, status) {
  return {
    nombre_invitado: name,
    grupo_id: summaryRecord.grupo_id,
    grupo_invitacion: summaryRecord.nombres_invitados,
    cantidad_grupo: summaryRecord.cantidad_invitados,
    estado: status,
    nombre_buscado: summaryRecord.nombre_buscado,
    fecha_respuesta: summaryRecord.fecha_respuesta
  };
}

function buildSummaryRecord(group, responseState) {
  const statusMap = getGroupStatusMap(group, responseState);
  const yesNames = group.names.filter(name => statusMap[normalizeText(name)] === 'Sí');
  const noNames = group.names.filter(name => statusMap[normalizeText(name)] === 'No');
  const pendingNames = group.names.filter(name => !statusMap[normalizeText(name)]);
  const quantity = group.quantity || group.names.length;
  const pendingCount = Math.max(0, quantity - yesNames.length - noNames.length);
  const latestDate = responseState ? responseState.latestDate : '';

  return {
    grupo_id: group.grupoId,
    nombres_invitados: group.names.join('\n'),
    cantidad_invitados: quantity,
    estado_grupo: getGroupStatus(quantity, yesNames.length, noNames.length, pendingCount),
    asistentes_confirmados: yesNames.join('\n'),
    cantidad_asisten: yesNames.length,
    no_asisten: noNames.join('\n'),
    cantidad_no_asisten: noNames.length,
    pendientes_invitados: pendingNames.join('\n'),
    cantidad_pendientes: pendingCount,
    nombre_buscado: responseState ? responseState.nombreBuscado : '',
    fecha_respuesta: latestDate
  };
}

function buildSearchGuestGroup(group, responseState) {
  const statusMap = getGroupStatusMap(group, responseState);
  const guestStatuses = group.names.map(name => ({
    nombre: name,
    asistencia: statusMap[normalizeText(name)] || '',
    confirmado: Boolean(statusMap[normalizeText(name)])
  }));
  const pendingNames = guestStatuses
    .filter(item => !item.confirmado)
    .map(item => item.nombre);

  return {
    grupoId: group.grupoId,
    names: group.names,
    quantity: group.quantity,
    guestStatuses: guestStatuses,
    pendingNames: pendingNames,
    pendingCount: pendingNames.length,
    answeredCount: guestStatuses.length - pendingNames.length,
    isClosed: pendingNames.length === 0
  };
}

function getGroupStatusMap(group, responseState) {
  const statusMap = {};
  if (!responseState || !responseState.byName) return statusMap;

  group.names.forEach(name => {
    const key = normalizeText(name);
    if (responseState.byName[key]) {
      statusMap[key] = responseState.byName[key].asistencia;
    }
  });

  return statusMap;
}

function scoreNames(names, normalizedQuery, queryTokens) {
  if (!names || !names.length) return noMatch();

  return names
    .map(name => scoreSingleName(normalizeText(name), normalizedQuery, queryTokens))
    .sort((a, b) => b.score - a.score)[0] || noMatch();
}

function getLatestResponsesByGroup() {
  return getResponseStatesByGroup();
}

function getResponseStatesByGroup() {
  const sheet = getResponsesSpreadsheet().getSheetByName(RESPONSES_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const states = {};

  values.slice(1).forEach(row => {
    const record = rowToRecord(headers, row);
    const groupId = String(record.grupo_id || '').trim();
    if (!groupId) return;

    updateResponseState(states, record);
  });

  return states;
}

function updateResponseState(states, record) {
  const groupId = String(record.grupo_id || '').trim();
  if (!groupId) return;

  const timestamp = toTimestamp(record.fecha_respuesta);
  const confirmations = getRecordConfirmations(record);
  if (!confirmations.length) return;

  if (!states[groupId]) {
    states[groupId] = {
      grupoId: groupId,
      invitedNames: splitNames(record.nombres_invitados || record.nombre),
      quantity: toPositiveNumber(record.cantidad_invitados || record.invitados, 0),
      byName: {},
      latestTimestamp: 0,
      latestDate: '',
      nombreBuscado: ''
    };
  }

  const state = states[groupId];
  if (!state.invitedNames.length) {
    state.invitedNames = splitNames(record.nombres_invitados || record.nombre);
  }
  if (!state.quantity) {
    state.quantity = toPositiveNumber(record.cantidad_invitados || record.invitados, state.invitedNames.length);
  }

  confirmations.forEach(item => {
    if (!item.asistencia) return;

    const matchedName = matchKnownGuestName(item.nombre, state.invitedNames);
    const key = normalizeText(matchedName || item.nombre);
    const current = state.byName[key];

    if (!current || timestamp >= current.timestamp) {
      state.byName[key] = {
        nombre: matchedName || item.nombre,
        asistencia: item.asistencia,
        timestamp: timestamp,
        fechaRespuesta: record.fecha_respuesta || ''
      };
    }
  });

  if (timestamp >= state.latestTimestamp) {
    state.latestTimestamp = timestamp;
    state.latestDate = record.fecha_respuesta || '';
    state.nombreBuscado = record.nombre_buscado || state.nombreBuscado || '';
  }
}

function getRecordConfirmations(record) {
  const existingNames = splitNames(record.nombres_invitados || record.nombre);
  const officialGroup = existingNames.length ? null : getOfficialGroupForRecord(record);
  const invitedNames = officialGroup
    ? officialGroup.names
    : existingNames;
  const quantity = toPositiveNumber(
    officialGroup ? officialGroup.quantity : record.cantidad_invitados || record.invitados,
    invitedNames.length
  );
  const confirmations = normalizeConfirmations(
    parseGuestConfirmations(record.confirmaciones_invitados),
    invitedNames
  );
  if (confirmations.length) return confirmations;

  const yesNames = splitNames(record.asistentes_confirmados)
    .map(name => ({ nombre: name, asistencia: 'Sí' }));
  const noNames = splitNames(record.no_asisten)
    .map(name => ({ nombre: name, asistencia: 'No' }));
  const explicitNames = normalizeConfirmations(yesNames.concat(noNames), invitedNames);
  if (explicitNames.length) return explicitNames;

  return inferLegacyConfirmations(record, invitedNames, quantity);
}

function matchKnownGuestName(name, invitedNames) {
  const normalizedName = normalizeText(name);
  return invitedNames.find(invitedName => normalizeText(invitedName) === normalizedName) || '';
}

function normalizeConfirmations(confirmations, invitedNames) {
  const seen = {};

  return (confirmations || [])
    .map(item => {
      const matchedName = matchKnownGuestName(item.nombre, invitedNames) || item.nombre;

      return {
        nombre: matchedName,
        asistencia: normalizeAttendance(item.asistencia)
      };
    })
    .filter(item => item.nombre && item.asistencia)
    .filter(item => {
      const key = normalizeText(item.nombre);
      if (!key || seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

function inferLegacyConfirmations(record, invitedNames, quantity) {
  if (!invitedNames.length) return [];

  const status = normalizeRsvpStatus(record.asistencia || record.confirmacion);
  if (!status) return [];

  const matchedNames = matchLegacyNamedGuests(record, invitedNames);
  const attendingCount = getLegacyAttendingCount(record, quantity);

  if (status === 'Sí') {
    if (isWholeGroupLegacyYes(attendingCount, quantity, invitedNames)) {
      return buildConfirmationItems(invitedNames, 'Sí');
    }

    if (matchedNames.length) {
      return buildConfirmationItems(limitMatchedNames(matchedNames, attendingCount), 'Sí');
    }

    return [];
  }

  if (status === 'No') {
    if (quantity <= 1 || invitedNames.length === 1 || !matchedNames.length) {
      return buildConfirmationItems(invitedNames, 'No');
    }

    return buildConfirmationItems(matchedNames, 'No');
  }

  if (status === 'Parcial') {
    if (matchedNames.length) {
      return buildConfirmationItems(limitMatchedNames(matchedNames, attendingCount), 'Sí');
    }
  }

  return [];
}

function getOfficialGroupForRecord(record) {
  const groups = getGuestGroups();
  const groupId = String(record.grupo_id || '').trim();

  if (groupId) {
    const exactGroup = groups.find(group => group.grupoId === groupId);
    if (exactGroup) return exactGroup;
  }

  const names = splitNames(record.nombres_invitados);
  if (names.length) {
    const normalizedNames = normalizeText(names.join('\n'));
    const byNames = groups.find(group => normalizeText(group.names.join('\n')) === normalizedNames);
    if (byNames) return byNames;
  }

  const sources = getLegacyNameSources(record);
  const candidates = [];

  sources.forEach(source => {
    const normalizedSource = normalizeText(source);
    const tokens = tokenize(normalizedSource);
    if (!tokens.length) return;

    groups.forEach(group => {
      const match = scoreNames(group.names, normalizedSource, tokens);
      if (match.score >= 90) {
        candidates.push({
          group: group,
          score: match.score
        });
      }
    });
  });

  candidates.sort((a, b) => b.score - a.score);
  if (!candidates.length) return null;

  const best = candidates[0];
  const sameBest = candidates.filter(candidate => candidate.score === best.score);
  return sameBest.length === 1 ? best.group : null;
}

function matchLegacyNamedGuests(record, invitedNames) {
  const candidates = [];

  getLegacyNameSources(record).forEach(source => {
    const normalizedSource = normalizeText(source);
    const tokens = tokenize(normalizedSource);
    if (!tokens.length) return;

    invitedNames.forEach(name => {
      const match = scoreSingleName(normalizeText(name), normalizedSource, tokens);
      if (match.score >= 70) {
        candidates.push({
          name: name,
          score: match.score
        });
      }
    });
  });

  candidates.sort((a, b) => b.score - a.score);

  return uniqueNames(candidates.map(candidate => candidate.name));
}

function getLegacyNameSources(record) {
  return [
    record.nombre,
    record.nombre_buscado,
    record.asistentes_confirmados,
    record.no_asisten
  ]
    .map(value => String(value || '').trim())
    .filter(Boolean);
}

function getLegacyAttendingCount(record, quantity) {
  const values = [
    record.total_confirmados,
    record.cantidad_asisten,
    record.invitados
  ];

  for (let index = 0; index < values.length; index++) {
    const count = Number(values[index]);
    if (count > 0) return count;
  }

  return quantity === 1 ? 1 : 0;
}

function isWholeGroupLegacyYes(attendingCount, quantity, invitedNames) {
  if (quantity <= 1 || invitedNames.length === 1) return true;
  if (attendingCount >= quantity) return true;
  if (quantity <= invitedNames.length && attendingCount >= invitedNames.length) return true;
  return false;
}

function limitMatchedNames(names, expectedCount) {
  if (!expectedCount || expectedCount >= names.length) return names;
  return names.slice(0, expectedCount);
}

function buildConfirmationItems(names, asistencia) {
  return uniqueNames(names).map(name => ({
    nombre: name,
    asistencia: asistencia
  }));
}

function prepareResponsesSheet(options) {
  const settings = options || {};
  const sheet = getOrCreateSheet(RESPONSES_SHEET_NAME);
  normalizeLegacyResponseSheet(sheet);
  ensureHeaders(sheet, RESPONSE_HEADERS);

  if (!settings.fast) {
    formatRsvpSheet(sheet, sheet.getLastColumn(), { tabColor: '#6E1F2C' });
  }

  return sheet;
}

function normalizeLegacyResponseSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (!lastRow || !lastColumn) return;

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(String);
  if (isCanonicalResponseSheet(headers)) return;

  const records = values.slice(1)
    .map(row => rowToRsvpData(headers, row))
    .filter(hasMeaningfulHistoricRecord)
    .map(record => buildResponseRecord(record));

  rewriteSheetWithRecords(sheet, RESPONSE_HEADERS, records, { tabColor: '#6E1F2C' });
}

function isCanonicalResponseSheet(headers) {
  const normalizedHeaders = headers.map(normalizeText);

  return RESPONSE_HEADERS.every((header, index) => (
    normalizedHeaders[index] === normalizeText(header)
  ));
}

function rowToRsvpData(headers, row) {
  const record = {};

  headers.forEach((header, index) => {
    const key = canonicalResponseKey(header);
    if (!key) return;

    record[key] = row[index];
  });

  return record;
}

function canonicalResponseKey(header) {
  const key = normalizeText(header).replace(/\s+/g, '_');

  if (!key || key === 'dashboard') return '';
  if (key === 'timestamp') return 'fecha_respuesta';
  if (key === 'nombre') return 'nombre';
  if (key === 'invitados') return 'invitados';
  if (key === 'asistencia') return 'asistencia';
  if (key === 'confirmacion') return 'confirmacion';
  if (key === 'confirmaciones') return 'confirmaciones_invitados';
  if (key.indexOf('confirmaciones') === 0) return 'confirmaciones_invitados';

  return RESPONSE_HEADERS.indexOf(key) >= 0 ? key : '';
}

function hasMeaningfulHistoricRecord(record) {
  if (
    record.grupo_id ||
    record.nombres_invitados ||
    record.nombre ||
    record.nombre_buscado ||
    record.asistentes_confirmados ||
    record.no_asisten
  ) {
    return true;
  }

  if (normalizeAttendance(record.asistencia || record.confirmacion)) return true;
  return parseGuestConfirmations(record.confirmaciones_invitados).length > 0;
}

function hasRsvpPayload(data) {
  return hasMeaningfulHistoricRecord(data);
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
  return nameTokens.some(nameToken => nameToken === token);
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
  if (queryToken.charAt(0) !== nameToken.charAt(0)) return false;
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

function parseGuestConfirmations(rawValue) {
  if (!rawValue) return [];

  let parsed = rawValue;
  if (!Array.isArray(parsed)) {
    try {
      parsed = JSON.parse(String(rawValue));
    } catch (err) {
      return [];
    }
  }

  if (!Array.isArray(parsed)) return [];

  return parsed
    .map(item => ({
      nombre: String((item && (item.nombre || item.name)) || '').trim(),
      asistencia: normalizeAttendance(item && (item.asistencia || item.attending))
    }))
    .filter(item => item.nombre);
}

function normalizeAttendance(value) {
  if (value === true) return 'Sí';
  if (value === false) return 'No';

  const normalized = normalizeText(value);
  if (['si', 'sí', 'yes', 'asiste', 'confirmado', 'confirmada'].indexOf(normalized) >= 0) {
    return 'Sí';
  }
  if (['no', 'no asiste', 'rechazado', 'rechazada'].indexOf(normalized) >= 0) {
    return 'No';
  }

  return '';
}

function normalizeRsvpStatus(value) {
  const attendance = normalizeAttendance(value);
  if (attendance) return attendance;

  const normalized = normalizeText(value);
  if (['parcial', 'partial'].indexOf(normalized) >= 0) return 'Parcial';

  return '';
}

function getAttendanceStatus(quantity, yesCount, noCount, pendingCount) {
  if (pendingCount > 0) return '';
  if (quantity > 0 && yesCount >= quantity) return 'Sí';
  if (quantity > 0 && noCount >= quantity) return 'No';
  if (yesCount > 0 && noCount > 0) return 'Parcial';
  if (yesCount > 0) return 'Sí';
  if (noCount > 0) return 'No';
  return '';
}

function getGroupStatus(quantity, yesCount, noCount, pendingCount) {
  if (pendingCount > 0 && (yesCount > 0 || noCount > 0)) return 'Parcial con pendientes';
  if (pendingCount > 0) return 'Pendiente';
  if (quantity > 0 && yesCount >= quantity) return 'Todos asisten';
  if (quantity > 0 && noCount >= quantity) return 'No asiste el grupo';
  if (yesCount > 0 && noCount > 0) return 'Asistencia parcial';
  if (yesCount > 0) return 'Asisten';
  if (noCount > 0) return 'No asisten';
  return 'Pendiente';
}

function splitNames(value) {
  return String(value || '')
    .split(/\r?\n/)
    .map(name => name.trim())
    .filter(Boolean);
}

function uniqueNames(names) {
  const seen = {};
  return names.filter(name => {
    const key = normalizeText(name);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function computePendingNames(invitedNames, yesNames, noNames) {
  const answered = {};
  yesNames.concat(noNames).forEach(name => {
    answered[normalizeText(name)] = true;
  });

  return invitedNames.filter(name => !answered[normalizeText(name)]);
}

function toPositiveNumber(value, fallback) {
  const number = Number(value);
  return number > 0 ? number : Number(fallback || 0);
}

function toTimestamp(value) {
  if (value instanceof Date) return value.getTime();

  const timestamp = new Date(value).getTime();
  return isNaN(timestamp) ? 0 : timestamp;
}

function rowToRecord(headers, row) {
  const record = {};

  headers.forEach((header, index) => {
    const value = row[index];
    const key = normalizeText(header).replace(/\s+/g, '_');

    record[header] = value;
    record[key] = value;
  });

  return record;
}

function appendRecordByHeaders(sheet, record) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const row = headers.map(header => getRecordValue(record, header));

  sheet.appendRow(row);
}

function rewriteSheetWithRecords(sheet, headers, records, options) {
  const settings = options || {};

  if (settings.fast && sheet.getLastRow() && sheet.getLastColumn()) {
    rewriteSheetContents(sheet, headers, records);
    return;
  }

  resetSheet(sheet);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (records.length) {
    const rows = records.map(record => headers.map(header => getRecordValue(record, header)));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  formatRsvpSheet(sheet, headers.length, settings);
  addSheetFilter(sheet);
}

function rewriteSheetContents(sheet, headers, records) {
  resetSheetContents(sheet);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (records.length) {
    const rows = records.map(record => headers.map(header => getRecordValue(record, header)));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function resetSheet(sheet) {
  try {
    sheet.getDataRange().breakApart();
  } catch (err) {
    console.error(err);
  }

  sheet.clear();
}

function resetSheetContents(sheet) {
  if (typeof sheet.clearContents === 'function') {
    sheet.clearContents();
    return;
  }

  sheet.clear();
}

function getRecordValue(record, header) {
  const normalizedHeader = normalizeText(header);
  const keys = Object.keys(record);

  for (let i = 0; i < keys.length; i++) {
    if (normalizeText(keys[i]) === normalizedHeader) {
      return record[keys[i]];
    }
  }

  return '';
}

function formatRsvpSheet(sheet, columnCount, options) {
  if (!columnCount) return;

  const settings = options || {};
  const headerColor = settings.headerColor || '#6E1F2C';
  const tabColor = settings.tabColor || headerColor;

  sheet.setFrozenRows(1);
  sheet.setTabColor(tabColor);
  sheet
    .getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground(headerColor)
    .setFontColor('#FFF9F4');
  sheet.getDataRange()
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontFamily('Arial');
  applyColumnWidths(sheet);
  paintStatusColumns(sheet);
}

function formatTotalsSheet(sheet) {
  sheet.setFrozenRows(4);
  sheet.setTabColor('#2D161B');
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 340);

  sheet.getRange(1, 1)
    .setBackground('#2D161B')
    .setFontColor('#FFF9F4')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  sheet.getRange(2, 1, 1, 2)
    .setBackground('#F6ECE4')
    .setFontColor('#6E1F2C')
    .setFontWeight('bold');

  sheet.getRange(4, 1, 1, 3)
    .setBackground('#6E1F2C')
    .setFontColor('#FFF9F4')
    .setFontWeight('bold');

  sheet.getRange(5, 1, 4, 3)
    .setBackground('#F6ECE4')
    .setFontColor('#2B1D21');

  sheet.getRange(10, 1, 4, 3)
    .setBackground('#F6ECE4')
    .setFontColor('#2B1D21');

  sheet.getRange(5, 2, 4, 1)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(10, 2, 4, 1)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(5, 2).setBackground('#DDEBDD').setFontColor('#2F6B4F');
  sheet.getRange(6, 2).setBackground('#DDEBDD').setFontColor('#2F6B4F');
  sheet.getRange(7, 2).setBackground('#F1DDE0').setFontColor('#6E1F2C');
  sheet.getRange(8, 2).setBackground('#F6E7C9').setFontColor('#7A4B16');

  sheet.getDataRange()
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setFontFamily('Arial');
}

function applyColumnWidths(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const widths = {
    fecha_respuesta: 150,
    grupo_id: 110,
    nombres_invitados: 260,
    cantidad_invitados: 120,
    nombre_buscado: 160,
    nombre: 180,
    invitados: 90,
    confirmacion: 120,
    asistencia: 120,
    confirmaciones_invitados: 260,
    asistentes_confirmados: 240,
    no_asisten: 220,
    total_confirmados: 120,
    estado_grupo: 170,
    cantidad_asisten: 120,
    cantidad_no_asisten: 130,
    pendientes_invitados: 240,
    cantidad_pendientes: 130,
    nombre_invitado: 220,
    grupo_invitacion: 280,
    cantidad_grupo: 110,
    estado: 120
  };

  headers.forEach((header, index) => {
    const key = normalizeText(header).replace(/\s+/g, '_');
    sheet.setColumnWidth(index + 1, widths[key] || 150);
  });
}

function paintStatusColumns(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || !lastColumn) return;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const statusColumn = headers.findIndex(header => {
    const key = normalizeText(header).replace(/\s+/g, '_');
    return key === 'estado_grupo' || key === 'estado';
  }) + 1;

  if (!statusColumn) return;

  const range = sheet.getRange(2, statusColumn, lastRow - 1, 1);
  const values = range.getValues();
  const backgrounds = values.map(row => [getStatusBackground(row[0])]);
  const colors = values.map(row => [getStatusTextColor(row[0])]);

  range
    .setBackgrounds(backgrounds)
    .setFontColors(colors)
    .setFontWeight('bold');
}

function getStatusBackground(status) {
  const normalized = normalizeText(status);

  if (normalized === 'confirmado' || normalized === 'todos asisten' || normalized === 'asisten') {
    return '#DDEBDD';
  }
  if (normalized === 'no asiste' || normalized === 'no asiste el grupo' || normalized === 'no asisten') {
    return '#F1DDE0';
  }
  if (normalized === 'pendiente') {
    return '#F6E7C9';
  }
  if (normalized === 'parcial' || normalized === 'parcial con pendientes' || normalized === 'asistencia parcial') {
    return '#E9D7DC';
  }

  return '#FFF9F4';
}

function getStatusTextColor(status) {
  const normalized = normalizeText(status);

  if (normalized === 'confirmado' || normalized === 'todos asisten' || normalized === 'asisten') {
    return '#2F6B4F';
  }
  if (normalized === 'no asiste' || normalized === 'no asiste el grupo' || normalized === 'no asisten') {
    return '#6E1F2C';
  }
  if (normalized === 'pendiente') {
    return '#7A4B16';
  }
  if (normalized === 'parcial' || normalized === 'parcial con pendientes' || normalized === 'asistencia parcial') {
    return '#5E2E37';
  }

  return '#302125';
}

function addSheetFilter(sheet) {
  try {
    const currentFilter = sheet.getFilter();
    if (currentFilter) currentFilter.remove();
    sheet.getDataRange().createFilter();
  } catch (err) {
    console.error(err);
  }
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

function parseSubmitData(params) {
  if (params && params.payload) {
    try {
      return JSON.parse(String(params.payload));
    } catch (err) {
      throw new Error('No se pudo leer la confirmación enviada.');
    }
  }

  return params || {};
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
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const hasHeaders = current.some(Boolean);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const normalizedCurrent = current.map(normalizeText);
  headers.forEach((header, index) => {
    const normalizedHeader = normalizeText(header);
    if (normalizedCurrent.includes(normalizedHeader)) return;

    if (!current[index]) {
      current[index] = header;
      normalizedCurrent[index] = normalizedHeader;
      return;
    }

    current.push(header);
    normalizedCurrent.push(normalizedHeader);
  });

  sheet.getRange(1, 1, 1, current.length).setValues([current]);
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
