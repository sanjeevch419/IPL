function doGet(e) {
  const name = ((e.parameter || {}).name || '').trim();
  const pin  = ((e.parameter || {}).pin  || '').trim();

  if (!name || !pin) {
    return respond({ ok: false, error: 'Missing name or pin' });
  }

  try {
    const data = loadAppData();

    if (!data) {
      return respond({ ok: false, error: 'No data saved yet. Ask the manager to post a match first.' });
    }

    const pins = data.pins || {};

    const playerPin = pins[name];
    if (!playerPin || playerPin !== pin) {
      return respond({ ok: false, error: 'Incorrect name or PIN. Contact the manager.' });
    }

    const playerRows = (data.rows || []).map(function(row) {
      const bet = row.ratio && row.ratio.bets && row.ratio.bets[name];
      if (!bet) return null;
      return {
        date:   row.date,
        match:  row.match,
        teams:  row.teams,
        winner: row.winner,
        phase:  row.phase || 1,
        ratio: {
          underdog: row.ratio.underdog,
          low:      row.ratio.low,
          high:     row.ratio.high,
          roundNum: row.ratio.roundNum,
          bets: { [name]: bet }
        }
      };
    }).filter(Boolean);

    const rawCleared = data.cleared || {};
    const cleared = {};
    [1, 2, 3, 4, 5, 6].forEach(function(ph) {
      cleared[ph] = { [name]: ((rawCleared[ph] || {})[name] || 0) };
    });

    return respond({
      ok:      true,
      rows:    playerRows,
      mgrPct:  data.mgrPct || 1,
      cleared: cleared
    });

  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function testAuth() {
  const data = loadAppData();
  Logger.log(data ? 'OK' : 'No data');
}

function loadAppData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MLC_APP_DATA');
  if (!sheet) throw new Error('MLC_APP_DATA sheet not found.');
  const chunks = sheet.getRange(1, 1, 20, 1).getValues()
    .map(function(r){ return r[0] || ''; }).join('');
  if (!chunks) return null;
  return JSON.parse(chunks);
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
