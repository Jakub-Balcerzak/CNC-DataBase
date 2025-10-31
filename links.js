/** Links management: compare, sync, mass-check */

function checkDriveLinkStatus(link) {
  if (!link || link === '(brak linku)') return { ok: false, code: 0, status: 'Brak linku' };

  const fileIdMatch = link.match(/[-\w]{25,}/);
  if (!fileIdMatch) return { ok: false, code: 0, status: 'Nieprawidłowy format linku' };

  const fileId = fileIdMatch[0];
  try {
    const file = DriveApp.getFileById(fileId);
    if (!file) return { ok: false, code: 404, status: 'Nie znaleziono pliku' };
    const name = file.getName();
    return { ok: true, code: 200, status: 'OK', name };
  } catch (e) {
    if (String(e).includes('File not found')) return { ok: false, code: 404, status: 'Nie znaleziono pliku' };
    if (String(e).includes('User does not have permission')) return { ok: false, code: 403, status: 'Brak dostępu' };
    return { ok: false, code: 500, status: 'Błąd: ' + e.message };
  }
}

function updateLinks(ss, foundLinks, elementName, newLink, ui) {
  let updatedCount = 0;

  for (const f of foundLinks) {
    const sheet = ss.getSheetByName(f.sheet);
    if (!sheet) continue;
    const col = f.col || 2; // jeśli brak informacji o kolumnie - domyślnie B
    const cell = sheet.getRange(f.row, col);
    const text = cell.getDisplayValue() || elementName;
    const newRich = SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setLinkUrl(newLink)
      .build();
    cell.setRichTextValue(newRich);
    updatedCount++;
  }

  ui.alert(`✅ Zaktualizowano ${updatedCount} linków dla "${elementName}".`);
}

function promptAndCompareLinks() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const elementNameResponse = ui.prompt('Porównaj linki', 'Podaj nazwę elementu:', ui.ButtonSet.OK_CANCEL);
  if (elementNameResponse.getSelectedButton() !== ui.Button.OK) return;

  const elementName = elementNameResponse.getResponseText().trim();
  if (!elementName) return ui.alert('Nie podano nazwy elementu.');

  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC', 'Elementy CNC'];
  const foundLinks = [];

  // Zbierz wszystkie linki — sprawdzamy kolumny A (0) i B (1)
  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const range = sheet.getDataRange();
    const values = range.getValues();
    const richValues = range.getRichTextValues();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c <= 1; c++) {
        const cellVal = String(values[r][c]).trim();
        if (cellVal === elementName) {
          let link = null;
          try {
            link = richValues[r][c]?.getLinkUrl() || '(brak linku)';
          } catch (e) {
            link = '(brak linku)';
          }

          const status = checkDriveLinkStatus(link);
          foundLinks.push({ sheet: sheetName, row: r + 1, col: c + 1, link, ...status });
        }
      }
    }
  }

  if (foundLinks.length === 0) return ui.alert(`Nie znaleziono elementu "${elementName}".`);

  const validLinks = foundLinks.filter(f => f.ok);
  const invalidLinks = foundLinks.filter(f => !f.ok);

  const hasValid = validLinks.length > 0;
  const onlyMissing = invalidLinks.every(f => f.status === 'Brak linku');

  if (invalidLinks.length > 0 && !(hasValid && onlyMissing)) {
    const msg = invalidLinks.map(f => {
      const colLetter = f.col ? String.fromCharCode(64 + f.col) : 'B';
      return `• ${f.sheet}!${colLetter}${f.row} → ${f.status}`;
    }).join('\n');
    const response = ui.prompt(
      'Znaleziono błędne linki',
      `Dla elementu "${elementName}" wykryto błędne linki:\n\n${msg}\n\nPodaj nowy, poprawny link:`,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      const newLink = response.getResponseText().trim();
      if (newLink) updateLinks(ss, foundLinks, elementName, newLink, ui);
    }
    return;
  }

  const allLinks = foundLinks.map(f => f.link).filter(l => l && l !== '(brak linku)');
  const uniqueLinks = [...new Set(allLinks)];

  if (uniqueLinks.length === 0) {
    ui.alert(`❌ Brak jakichkolwiek linków dla "${elementName}".`);
    return;
  }

  if (uniqueLinks.length > 1 || (hasValid && onlyMissing)) {
    let msg = `Znaleziono ${uniqueLinks.length} różne linki (lub brak w niektórych miejscach) dla "${elementName}":\n\n`;
    uniqueLinks.forEach((l, i) => { msg += `${i + 1}. ${l}\n`; });
    msg += `\nWpisz numer linku, który chcesz zachować (lub wklej nowy link):`;

    const response = ui.prompt('Różne lub brakujące linki', msg, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      const userInput = response.getResponseText().trim();
      let selectedLink = null;

      if (/^\d+$/.test(userInput)) {
        const idx = parseInt(userInput, 10) - 1;
        selectedLink = uniqueLinks[idx];
      } else {
        selectedLink = userInput;
      }

      if (selectedLink) {
        updateLinks(ss, foundLinks, elementName, selectedLink, ui);
      }
    }
  } else {
    ui.alert(`✅ Wszystkie linki dla "${elementName}" są poprawne i jednakowe.`);
  }
}

function promptAndSyncLink() {
  const ui = SpreadsheetApp.getUi();

  const resp1 = ui.prompt(
    'Ustaw / edytuj link',
    'Podaj numer katalogowy elementu (np. H_P3300_04):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp1.getSelectedButton() !== ui.Button.OK) return;
  const elementName = resp1.getResponseText().trim();
  if (!elementName) {
    ui.alert('Nie podano numeru katalogowego elementu.');
    return;
  }

  const resp2 = ui.prompt(
    'Nowy link',
    'Podaj pełny adres URL (np. https://drive.google.com/file/d/....):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp2.getSelectedButton() !== ui.Button.OK) return;
  const newLink = resp2.getResponseText().trim();
  if (!newLink) {
    ui.alert('Nie podano linku.');
    return;
  }

  try {
    const count = syncElementLink(elementName, newLink);
    ui.alert('Zakończono', `Podmieniono lub ustawiono link w ${count} komórkach.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function syncElementLink(elementName, link) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC'];
  let totalUpdated = 0;

  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c <= 1; c++) {
        const cellValue = String(values[r][c]).trim();
        if (cellValue === elementName) {
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(cellValue)
            .setLinkUrl(link)
            .build();

          sheet.getRange(r + 1, c + 1).setRichTextValue(richText);
          totalUpdated++;
        }
      }
    }
  }

  return totalUpdated;
}

function massCheckAndFixLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC', 'Elementy CNC'];
  const ui = SpreadsheetApp.getUi();

  const elementLinks = {};

  sheetsToCheck.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const richTexts = sheet.getDataRange().getRichTextValues();

    for (let i = 1; i < data.length; i++) {
      for (let c = 0; c <= 1; c++) {
        const name = data[i][c];
        if (!name) continue;
        const linkObj = richTexts[i][c];
        const link = linkObj && linkObj.getLinkUrl ? linkObj.getLinkUrl() : null;

        if (!elementLinks[name]) elementLinks[name] = [];
        elementLinks[name].push({ sheet: sheetName, row: i + 1, col: c + 1, link });
      }
    }
  });

  const inconsistencies = [];
  const missingLinks = [];

  for (const [name, entries] of Object.entries(elementLinks)) {
    const uniqueLinks = [...new Set(entries.map(e => e.link).filter(l => !!l))];

    if (uniqueLinks.length === 0) {
      missingLinks.push({ name, entries });
    } else if (uniqueLinks.length === 1) {
      const validLink = uniqueLinks[0];
      entries.forEach(e => {
        if (!e.link) {
          const sheet = ss.getSheetByName(e.sheet);
          const col = e.col || 2;
          const cell = sheet.getRange(e.row, col);
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(name)
            .setLinkUrl(validLink)
            .build();
          cell.setRichTextValue(richText);
        }
      });
    } else {
      inconsistencies.push({ name, uniqueLinks, entries });
    }
  }

  if (inconsistencies.length > 0) {
    for (const conflict of inconsistencies) {
      const { name, uniqueLinks, entries } = conflict;
      let msg = `Element "${name}" ma różne linki:\n\n`;
      uniqueLinks.forEach((l, i) => { msg += `${i + 1}. ${l}\n`; });
      msg += `\nWpisz numer (1-${uniqueLinks.length}), który link ma być używany we wszystkich wystąpieniach.`;

      const button = ui.prompt(msg, ui.ButtonSet.OK_CANCEL);
      if (button.getSelectedButton() === ui.Button.OK) {
        const response = button.getResponseText().trim();
        const index = parseInt(response, 10);

        if (!isNaN(index) && index >= 1 && index <= uniqueLinks.length) {
          const chosen = uniqueLinks[index - 1];
          entries.forEach(e => {
            const sheet = ss.getSheetByName(e.sheet);
            const col = e.col || 2;
            const cell = sheet.getRange(e.row, col);
            const richText = SpreadsheetApp.newRichTextValue()
              .setText(name)
              .setLinkUrl(chosen)
              .build();
            cell.setRichTextValue(richText);
          });
        } else {
          ui.alert(`❌ Podano nieprawidłowy numer. Oczekiwano wartości od 1 do ${uniqueLinks.length}.`);
        }
      }
    }
  }

  if (missingLinks.length > 0) {
    let msg = "⚠️ Elementy, które nigdzie nie mają przypisanego linku:\n\n";
    missingLinks.forEach(e => { msg += `• ${e.name} (wystąpień: ${e.entries.length})\n`; });
    ui.alert(msg);
  } else if (inconsistencies.length === 0) {
    ui.alert("✅ Wszystkie elementy mają spójne linki i zostały uzupełnione, jeśli brakowało.");
  }
}
