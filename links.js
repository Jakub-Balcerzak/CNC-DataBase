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

/**
 * Masowe sprawdzanie i synchronizacja linków UCANCAM dla elementów.
 * 
 * LOGIKA:
 * - Zbiera ELEMENTY (nie moduły M.../X...) z kolumny B obu arkuszy
 * - Dla każdego elementu sprawdza kolumnę F (UCANCAM)
 * - Uzupełnia puste komórki jeśli jest jeden spójny link
 * - Pyta użytkownika w przypadku konfliktów
 */
function massCheckAndFixUcancamLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const sheetModuly = ss.getSheetByName('Moduły CNC');
  const sheetZestawy = ss.getSheetByName('Zestawy CNC');
  
  if (!sheetModuly || !sheetZestawy) {
    ui.alert('Błąd', 'Brakuje arkuszy "Moduły CNC" lub "Zestawy CNC".', ui.ButtonSet.OK);
    return;
  }

  // Kolumny (0-based)
  const COL_ELEMENT = 1;        // kolumna B
  const COL_UCANCAM = 5;        // kolumna F
  const COL_UPTODATE = 6;       // kolumna G (UP-TO-DATE?)
  const COL_UCANCAM_LETTER = 'F';
  const COL_UPTODATE_LETTER = 'G';

  // Helper: sprawdza czy nazwa to moduł (M... lub X...)
  function isModule(name) {
    if (!name) return false;
    return /^[MX]/i.test(String(name).trim());
  }

  // Helper: pobiera link z RichText lub null
  function getRichLink(richCell) {
    if (!richCell) return null;
    try {
      if (typeof richCell.getLinkUrl === 'function') {
        return richCell.getLinkUrl() || null;
      }
    } catch (e) {}
    return null;
  }

  // Helper: pobiera wartość checkboxa (true/false/null)
  function getCheckboxValue(sheet, row, col) {
    try {
      const cell = sheet.getRange(row, col);
      const value = cell.getValue();
      
      // Przypadek 1: Sprawdź czy komórka ma checkbox przez Data Validation
      const dataValidation = cell.getDataValidation();
      if (dataValidation) {
        const criteria = dataValidation.getCriteriaType();
        if (criteria === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
          // Jest checkbox - zwróć wartość boolean
          if (value === true) return true;
          if (value === false) return false;
          return null; // checkbox bez wartości
        }
      }
      
      // Przypadek 2: Sprawdź czy wartość to bezpośrednio boolean (TRUE/FALSE wpisane ręcznie)
      if (typeof value === 'boolean') {
        return value;
      }
      
      // Przypadek 3: Sprawdź czy wartość to tekst "TRUE" lub "FALSE"
      if (typeof value === 'string') {
        const upperValue = value.toUpperCase().trim();
        if (upperValue === 'TRUE') return true;
        if (upperValue === 'FALSE') return false;
      }
      
      // Brak checkboxa/wartości - zwróć null
      return null;
    } catch (e) {
      Logger.log(`Błąd odczytu checkboxa w wierszu ${row}, kolumnie ${col}: ${e.message}`);
      return null;
    }
  }

  // Helper: ustawia wartość checkboxa
  function setCheckboxValue(sheet, row, col, value) {
    try {
      const cell = sheet.getRange(row, col);
      
      // Uproszczenie (Optymalizacja #5)
      cell.setValue(Boolean(value));
      return true; // Sukces
      
    } catch (e) {
      // Jeśli wystąpił błąd "typed columns", zwróć false
      if (String(e.message).includes('typed column') || 
          String(e.message).includes('not allowed')) {
        return false; // Typed column - nie można zsynchronizować
      }
      
      // Inny błąd - spróbuj dodać Data Validation
      try {
        const dataValidation = cell.getDataValidation();
        if (!dataValidation) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireCheckbox()
            .setAllowInvalid(false)
            .build();
          cell.setDataValidation(rule);
        }
        cell.setValue(Boolean(value));
        return true; // Sukces
      } catch (e2) {
        return false; // Niepowodzenie
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // ZBIERANIE ELEMENTÓW I ICH LINKÓW UCANCAM + CHECKBOXÓW UP-TO-DATE
  // ═══════════════════════════════════════════════════════════════════════════
  
  const elementLinks = {};     // { "H_M1594_01": [ { sheet, row, col, link }, ... ] }
  const elementCheckboxes = {}; // { "H_M1594_01": [ { sheet, row, col, checked }, ... ] }

  // Cache arkuszy (Optymalizacja #2)
  const sheetsCache = {
    'Moduły CNC': sheetModuly,
    'Zestawy CNC': sheetZestawy
  };

  // Zbierz z arkusza 'Moduły CNC'
  const modData = sheetModuly.getDataRange().getValues();
  const modRich = sheetModuly.getDataRange().getRichTextValues();
  
  Logger.log(`=== ROZPOCZYNAM ZBIERANIE DANYCH ===`);
  Logger.log(`Moduły CNC - wierszy: ${modData.length}`);
  
  // Batch pobranie wszystkich checkboxów (Optymalizacja #3)
  const allModCheckboxValues = modData.length > 1 
    ? sheetModuly.getRange(2, COL_UPTODATE + 1, modData.length - 1, 1).getValues()
    : [];
  
  for (let i = 1; i < modData.length; i++) { // pomijamy nagłówek
    const elementName = String(modData[i][COL_ELEMENT]).trim(); // kolumna B
    
    // Pomijamy puste i moduły
    if (!elementName || isModule(elementName)) continue;
    
    // Inline getRichLink (Optymalizacja #4)
    let link = null;
    const richCell = modRich[i][COL_UCANCAM];
    if (richCell) {
      try {
        link = richCell.getLinkUrl() || null;
      } catch (e) {}
    }
    
    if (!elementLinks[elementName]) elementLinks[elementName] = [];
    elementLinks[elementName].push({
      sheet: 'Moduły CNC',
      row: i + 1,
      col: COL_UCANCAM + 1,
      link: link
    });
    
    // Batch odczyt checkboxa (Optymalizacja #3)
    const value = allModCheckboxValues[i - 1][0];
    let checked = null;
    
    if (typeof value === 'boolean') {
      checked = value;
    } else if (typeof value === 'string') {
      const upperValue = value.toUpperCase().trim();
      if (upperValue === 'TRUE') checked = true;
      else if (upperValue === 'FALSE') checked = false;
    }
    
    if (!elementCheckboxes[elementName]) elementCheckboxes[elementName] = [];
    elementCheckboxes[elementName].push({
      sheet: 'Moduły CNC',
      row: i + 1,
      col: COL_UPTODATE + 1,
      checked: checked
    });
  }

  // Zbierz z arkusza 'Zestawy CNC'
  const zestData = sheetZestawy.getDataRange().getValues();
  const zestRich = sheetZestawy.getDataRange().getRichTextValues();
  
  Logger.log(`Zestawy CNC - wierszy: ${zestData.length}`);
  
  // Batch pobranie wszystkich checkboxów (Optymalizacja #3)
  const allZestCheckboxValues = zestData.length > 1
    ? sheetZestawy.getRange(2, COL_UPTODATE + 1, zestData.length - 1, 1).getValues()
    : [];
  
  for (let i = 1; i < zestData.length; i++) { // pomijamy nagłówek
    const elementName = String(zestData[i][COL_ELEMENT]).trim(); // kolumna B
    
    // Pomijamy puste i moduły
    if (!elementName || isModule(elementName)) continue;
    
    // Inline getRichLink (Optymalizacja #4)
    let link = null;
    const richCell = zestRich[i][COL_UCANCAM];
    if (richCell) {
      try {
        link = richCell.getLinkUrl() || null;
      } catch (e) {}
    }
    
    if (!elementLinks[elementName]) elementLinks[elementName] = [];
    elementLinks[elementName].push({
      sheet: 'Zestawy CNC',
      row: i + 1,
      col: COL_UCANCAM + 1,
      link: link
    });
    
    // Batch odczyt checkboxa (Optymalizacja #3)
    const value = allZestCheckboxValues[i - 1][0];
    let checked = null;
    
    if (typeof value === 'boolean') {
      checked = value;
    } else if (typeof value === 'string') {
      const upperValue = value.toUpperCase().trim();
      if (upperValue === 'TRUE') checked = true;
      else if (upperValue === 'FALSE') checked = false;
    }
    
    if (!elementCheckboxes[elementName]) elementCheckboxes[elementName] = [];
    elementCheckboxes[elementName].push({
      sheet: 'Zestawy CNC',
      row: i + 1,
      col: COL_UPTODATE + 1,
      checked: checked
    });
  }
  
  Logger.log(`=== ZEBRANO ELEMENTÓW: ${Object.keys(elementCheckboxes).length} ===`);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SPRAWDZANIE SPÓJNOŚCI LINKÓW UCANCAM
  // ═══════════════════════════════════════════════════════════════════════════
  
  const inconsistencies = []; // różne linki dla tego samego elementu
  const missingLinks = [];    // elementy bez żadnego linku UCANCAM
  let autoFilledCount = 0;

  for (const [elementName, entries] of Object.entries(elementLinks)) {
    const uniqueLinks = [...new Set(entries.map(e => e.link).filter(l => !!l))];

    if (uniqueLinks.length === 0) {
      // Brak linku nigdzie
      missingLinks.push({ name: elementName, entries });
    } else if (uniqueLinks.length === 1) {
      // Jeden spójny link — uzupełnij puste komórki
      const validLink = uniqueLinks[0];
      for (const e of entries) {
        if (!e.link) {
          const sheet = sheetsCache[e.sheet]; // Użyj cache
          const cell = sheet.getRange(e.row, e.col);
          const currentText = cell.getDisplayValue() || elementName;
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(currentText)
            .setLinkUrl(validLink)
            .build();
          cell.setRichTextValue(richText);
          autoFilledCount++;
        }
      }
    } else {
      // Więcej niż jeden unikalny link → konflikt
      inconsistencies.push({ name: elementName, uniqueLinks, entries });
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SPRAWDZANIE SPÓJNOŚCI CHECKBOXÓW UP-TO-DATE?
  // ═══════════════════════════════════════════════════════════════════════════
  
  const checkboxInconsistencies = []; // różne stany checkboxów dla tego samego elementu
  let autoFilledCheckboxCount = 0;
  let skippedTypedColumns = 0; // licznik typed columns które nie mogły być zsynchronizowane

  for (const [elementName, entries] of Object.entries(elementCheckboxes)) {
    // Filtruj tylko te które mają wartość (true/false), ignoruj null (brak checkboxa)
    const validEntries = entries.filter(e => e.checked !== null);
    
    if (validEntries.length === 0) {
      // Żaden checkbox nie ma wartości - pomijamy
      continue;
    }
    
    const uniqueStates = [...new Set(validEntries.map(e => e.checked))];
    
    if (uniqueStates.length === 1) {
      // Jeden spójny stan — uzupełnij pozostałe checkboxy tym samym stanem
      const validState = uniqueStates[0];
      for (const e of entries) {
        if (e.checked === null) {
          // Checkbox nie ma wartości - spróbuj ustawić
          const sheet = sheetsCache[e.sheet]; // Użyj cache
          const success = setCheckboxValue(sheet, e.row, e.col, validState);
          if (success) {
            autoFilledCheckboxCount++;
          } else {
            skippedTypedColumns++; // Typed column - nie można zsynchronizować
          }
        }
      }
    } else {
      // Więcej niż jeden unikalny stan → konflikt
      checkboxInconsistencies.push({ name: elementName, uniqueStates, entries: validEntries });
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // ROZWIĄZYWANIE KONFLIKTÓW LINKÓW UCANCAM
  // ═══════════════════════════════════════════════════════════════════════════
  
  let fixedConflicts = 0;
  for (const conflict of inconsistencies) {
    const { name, uniqueLinks, entries } = conflict;
    
    // Przygotuj informację o lokalizacjach
    let locationsInfo = entries.map(e => {
      const linkInfo = e.link ? `ma link` : `BRAK linku`;
      return `• ${e.sheet}!${COL_UCANCAM_LETTER}${e.row} (${linkInfo})`;
    }).join('\n');
    
    let msg = `Element "${name}" ma różne linki UCANCAM:\n\n`;
    uniqueLinks.forEach((l, i) => {
      msg += `${i + 1}. ${l}\n`;
    });
    msg += `\nWystąpienia:\n${locationsInfo}\n`;
    msg += `\nWpisz numer (1-${uniqueLinks.length}), który link ma być używany we wszystkich wystąpieniach:`;

    const response = ui.prompt('Konflikt linków UCANCAM', msg, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      const userInput = response.getResponseText().trim();
      const index = parseInt(userInput, 10);

      if (!isNaN(index) && index >= 1 && index <= uniqueLinks.length) {
        const chosenLink = uniqueLinks[index - 1];
        
        // Ustaw wybrany link we wszystkich wystąpieniach
        for (const e of entries) {
          const sheet = sheetsCache[e.sheet]; // Użyj cache
          const cell = sheet.getRange(e.row, e.col);
          const currentText = cell.getDisplayValue() || name;
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(currentText)
            .setLinkUrl(chosenLink)
            .build();
          cell.setRichTextValue(richText);
        }
        fixedConflicts++;
      } else {
        ui.alert(`❌ Podano nieprawidłowy numer. Pomijam element "${name}".`);
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // ROZWIĄZYWANIE KONFLIKTÓW CHECKBOXÓW UP-TO-DATE?
  // ═══════════════════════════════════════════════════════════════════════════
  
  let fixedCheckboxConflicts = 0;
  for (const conflict of checkboxInconsistencies) {
    const { name, uniqueStates, entries } = conflict;
    
    // Przygotuj informację o lokalizacjach
    let locationsInfo = entries.map(e => {
      const stateIcon = e.checked === true ? '✓ zaznaczony' : '☐ odznaczony';
      return `• ${e.sheet}!${COL_UPTODATE_LETTER}${e.row} (${stateIcon})`;
    }).join('\n');
    
    let msg = `Element "${name}" ma różne stany checkboxa UP-TO-DATE?:\n\n`;
    msg += `1. ✓ Zaznaczony (TRUE)\n`;
    msg += `2. ☐ Odznaczony (FALSE)\n`;
    msg += `\nWystąpienia:\n${locationsInfo}\n`;
    msg += `\nWpisz numer (1 lub 2), który stan ma być używany we wszystkich wystąpieniach:`;

    const response = ui.prompt('Konflikt checkboxów UP-TO-DATE?', msg, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      const userInput = response.getResponseText().trim();
      const index = parseInt(userInput, 10);

      if (index === 1 || index === 2) {
        const chosenState = (index === 1); // 1 = true, 2 = false
        
        // Ustaw wybrany stan we wszystkich wystąpieniach (włącznie z tymi bez checkboxa)
        const allEntries = elementCheckboxes[name] || [];
        let successCount = 0;
        for (const e of allEntries) {
          const sheet = sheetsCache[e.sheet]; // Użyj cache
          const success = setCheckboxValue(sheet, e.row, e.col, chosenState);
          if (success) {
            successCount++;
          } else {
            skippedTypedColumns++;
          }
        }
        if (successCount > 0) {
          fixedCheckboxConflicts++;
        }
      } else {
        ui.alert(`❌ Podano nieprawidłowy numer. Pomijam element "${name}".`);
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // RAPORT KOŃCOWY
  // ═══════════════════════════════════════════════════════════════════════════
  
  const summaryLines = [];
  
  summaryLines.push(`📊 Sprawdzono ${Object.keys(elementLinks).length} elementów.`);
  summaryLines.push('');
  
  // Podsumowanie UCANCAM
  summaryLines.push(`📦 LINKI UCANCAM (kolumna F):`);
  if (autoFilledCount > 0) {
    summaryLines.push(`  ✅ Automatycznie uzupełniono ${autoFilledCount} pustych komórek.`);
  }
  if (fixedConflicts > 0) {
    summaryLines.push(`  🔧 Naprawiono konflikty dla ${fixedConflicts} elementów.`);
  }
  if (autoFilledCount === 0 && fixedConflicts === 0 && inconsistencies.length === 0) {
    summaryLines.push(`  ✅ Wszystkie linki są spójne.`);
  }
  
  summaryLines.push('');
  
  // Podsumowanie UP-TO-DATE?
  summaryLines.push(`☑️  CHECKBOXY UP-TO-DATE? (kolumna G):`);
  if (autoFilledCheckboxCount > 0) {
    summaryLines.push(`  ✅ Automatycznie uzupełniono ${autoFilledCheckboxCount} pustych checkboxów.`);
  }
  if (fixedCheckboxConflicts > 0) {
    summaryLines.push(`  🔧 Naprawiono konflikty dla ${fixedCheckboxConflicts} elementów.`);
  }
  if (skippedTypedColumns > 0) {
    summaryLines.push(`  ⚠️ Pominięto ${skippedTypedColumns} komórek z typed columns (wymagana ręczna synchronizacja).`);
  }
  if (autoFilledCheckboxCount === 0 && fixedCheckboxConflicts === 0 && checkboxInconsistencies.length === 0 && skippedTypedColumns === 0) {
    summaryLines.push(`  ✅ Wszystkie checkboxy są spójne.`);
  }
  
  if (missingLinks.length > 0) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Elementy bez linku UCANCAM (${missingLinks.length}):`);
    missingLinks.slice(0, 20).forEach(m => {
      summaryLines.push(`  • ${m.name} (wystąpień: ${m.entries.length})`);
    });
    if (missingLinks.length > 20) {
      summaryLines.push(`  ... i ${missingLinks.length - 20} innych`);
    }
  }

  ui.alert('Sprawdzenie UCANCAM i UP-TO-DATE zakończone', summaryLines.join('\n'), ui.ButtonSet.OK);
}
