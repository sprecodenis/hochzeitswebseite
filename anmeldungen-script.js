// Google Apps Script – Hochzeitsanmeldungen Valeria & Denis
// Anleitung: siehe unten

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Anmeldungen') || ss.getActiveSheet();

    // Kopfzeile beim ersten Eintrag setzen
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Zeitstempel', 'Personen', 'Details', 'Bemerkungen']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }

    const params = e.parameter || {};
    const timestamp = new Date().toLocaleString('de-CH', { timeZone: 'Europe/Zurich' });

    // Personen zusammenfassen
    const persons = [];
    for (let i = 1; i <= 12; i++) {
      if (params['p' + i + '_vorname']) {
        const name    = (params['p' + i + '_vorname'] || '') + ' ' + (params['p' + i + '_nachname'] || '');
        const vegi    = params['p' + i + '_vegi'] === 'ja' ? ' (vegetarisch)' : '';
        const allerg  = params['p' + i + '_allergien'] ? ' | Allergien: ' + params['p' + i + '_allergien'] : '';
        persons.push(name + vegi + allerg);
      }
    }

    sheet.appendRow([
      timestamp,
      persons.length,
      persons.join('\n'),
      params['bemerkungen'] || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/*
── EINRICHTEN (einmalig, 5 Minuten) ──────────────────────────────────────

1. Neue Google-Tabelle erstellen auf sheets.google.com
   → Tabelle umbenennen in "Hochzeitsanmeldungen"

2. Im Menü: Erweiterungen → Apps Script

3. Gesamten bestehenden Code löschen und diesen Code einfügen → Speichern

4. Bereitstellen → Neue Bereitstellung
   → Typ: Web-App
   → Ausführen als: Ich (deine Google-Adresse)
   → Zugriff: Jeder
   → Bereitstellen → Berechtigungen erteilen → URL kopieren

5. Die kopierte URL in index.html einfügen:
   Suche nach "DEIN-APPS-SCRIPT-URL" und ersetze es durch die kopierte URL

──────────────────────────────────────────────────────────────────────────
*/
