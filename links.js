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
 * - Pyta użytkownika w przypadku konfliktów (przez HTML dialog)
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
          const sheet = ss.getSheetByName(e.sheet);
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
          const sheet = sheetsCache[e.sheet];
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
  // ROZWIĄZYWANIE KONFLIKTÓW LINKÓW UCANCAM (HTML DIALOGS)
  // ═══════════════════════════════════════════════════════════════════════════
  
  if (inconsistencies.length > 0) {
    // Zapisz dane konfliktów i inne statystyki w cache do użycia przez HTML dialogi i podsumowanie
    PropertiesService.getScriptProperties().setProperty('ucancamConflicts', JSON.stringify({
      conflicts: inconsistencies,
      COL_UCANCAM_LETTER: COL_UCANCAM_LETTER,
      sheetsCache: Object.keys(sheetsCache)
    }));
    PropertiesService.getScriptProperties().setProperty('ucancamConflictIndex', '0');
    PropertiesService.getScriptProperties().setProperty('autoFilledCount', String(autoFilledCount));
    PropertiesService.getScriptProperties().setProperty('totalElements', String(Object.keys(elementLinks).length));
    PropertiesService.getScriptProperties().setProperty('missingLinks', JSON.stringify(missingLinks));
    
    // Pokaż pierwszy konflikt
    showNextUcancamConflict();
    
    // Nie kontynuuj dalej - reszta zostanie wykonana po rozwiązaniu konfliktów
    return;
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
        const chosenState = (index === 1);
        
        const allEntries = elementCheckboxes[name] || [];
        let successCount = 0;
        for (const e of allEntries) {
          const sheet = sheetsCache[e.sheet];
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
  // AKTUALIZACJA CHECKBOXÓW MODUŁÓW W ZESTAWACH CNC
  // ═══════════════════════════════════════════════════════════════════════════
  
  Logger.log(`=== ROZPOCZYNAM AKTUALIZACJĘ MODUŁÓW ===`);
  
  const updatedModules = updateModuleCheckboxesInZestawy(
    modData,
    allModCheckboxValues,
    zestData,
    sheetsCache,
    COL_ELEMENT,
    COL_UPTODATE,
    isModule
  );
  
  Logger.log(`Zaktualizowano ${updatedModules} modułów w arkuszu Zestawy CNC`);

  // ═══════════════════════════════════════════════════════════════════════════
  // RAPORT KOŃCOWY
  // ═══════════════════════════════════════════════════════════════════════════
  
  showFinalSummary(
    Object.keys(elementLinks).length,
    autoFilledCount,
    0, // fixedConflicts - będzie uzupełnione po rozwiązaniu konfliktów
    inconsistencies.length,
    autoFilledCheckboxCount,
    fixedCheckboxConflicts,
    checkboxInconsistencies.length,
    skippedTypedColumns,
    updatedModules,
    missingLinks
  );
}

/**
 * Pokazuje HTML dialog z następnym konfliktem linków UCANCAM.
 * Jeśli nie ma więcej konfliktów, kontynuuje przetwarzanie checkboxów i pokazuje podsumowanie.
 */
function showNextUcancamConflict() {
  const props = PropertiesService.getScriptProperties();
  const conflictsData = JSON.parse(props.getProperty('ucancamConflicts'));
  const currentIndex = parseInt(props.getProperty('ucancamConflictIndex') || '0');
  
  if (currentIndex >= conflictsData.conflicts.length) {
    // Wszystkie konflikty rozwiązane - kontynuuj przetwarzanie
    continueAfterUcancamConflicts();
    return;
  }
  
  const conflict = conflictsData.conflicts[currentIndex];
  const COL_UCANCAM_LETTER = conflictsData.COL_UCANCAM_LETTER;
  
  // Przygotuj dane dla HTML
  const locations = conflict.entries.map(e => ({
    sheet: e.sheet,
    row: e.row,
    column: COL_UCANCAM_LETTER,
    hasLink: !!e.link
  }));
  
  // Stwórz HTML z template
  const template = HtmlService.createTemplateFromFile('ucancamConflictResolver');
  template.elementName = conflict.name;
  template.uniqueLinks = conflict.uniqueLinks;
  template.locations = locations;
  template.currentConflict = currentIndex + 1;
  template.totalConflicts = conflictsData.conflicts.length;
  
  const html = template.evaluate()
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, `Konflikt linków UCANCAM (${currentIndex + 1}/${conflictsData.conflicts.length})`);
}

/**
 * Funkcja wywoływana z HTML dialog po wyborze linku.
 * Aplikuje wybrany link i pokazuje następny konflikt.
 */
function resolveUcancamConflict(elementName, selectedIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const conflictsData = JSON.parse(props.getProperty('ucancamConflicts'));
  const currentIndex = parseInt(props.getProperty('ucancamConflictIndex') || '0');
  
  const conflict = conflictsData.conflicts[currentIndex];
  const chosenLink = conflict.uniqueLinks[selectedIndex - 1];
  
  // Aplikuj wybrany link do wszystkich wystąpień
  for (const e of conflict.entries) {
    const sheet = ss.getSheetByName(e.sheet);
    const cell = sheet.getRange(e.row, e.col);
    const currentText = cell.getDisplayValue() || elementName;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(currentText)
      .setLinkUrl(chosenLink)
      .build();
    cell.setRichTextValue(richText);
  }
  
  Logger.log(`Rozwiązano konflikt ${currentIndex + 1}: ${elementName} -> link #${selectedIndex}`);
  
  // Przejdź do następnego konfliktu
  props.setProperty('ucancamConflictIndex', String(currentIndex + 1));
  showNextUcancamConflict();
}

/**
 * Kontynuuje przetwarzanie po rozwiązaniu wszystkich konfliktów UCANCAM.
 * Przetwarza checkboxy i pokazuje końcowe podsumowanie.
 */
function continueAfterUcancamConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const sheetModuly = ss.getSheetByName('Moduły CNC');
  const sheetZestawy = ss.getSheetByName('Zestawy CNC');
  
  const COL_ELEMENT = 1;
  const COL_UPTODATE = 6;
  const COL_UPTODATE_LETTER = 'G';
  
  const sheetsCache = {
    'Moduły CNC': sheetModuly,
    'Zestawy CNC': sheetZestawy
  };
  
  function isModule(name) {
    if (!name) return false;
    return /^[MX]/i.test(String(name).trim());
  }
  
  function setCheckboxValue(sheet, row, col, value) {
    try {
      const cell = sheet.getRange(row, col);
      cell.setValue(Boolean(value));
      return true;
    } catch (e) {
      if (String(e.message).includes('typed column') || 
          String(e.message).includes('not allowed')) {
        return false;
      }
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
        return true;
      } catch (e2) {
        return false;
      }
    }
  }
  
  // Pobierz dane o rozwiązanych konfliktach
  const props = PropertiesService.getScriptProperties();
  const conflictsData = JSON.parse(props.getProperty('ucancamConflicts'));
  const fixedConflicts = conflictsData.conflicts.length;
  
  // Pobierz dane które już zostały zebrane
  const modData = sheetModuly.getDataRange().getValues();
  const zestData = sheetZestawy.getDataRange().getValues();
  const allModCheckboxValues = modData.length > 1 
    ? sheetModuly.getRange(2, COL_UPTODATE + 1, modData.length - 1, 1).getValues()
    : [];
  const allZestCheckboxValues = zestData.length > 1
    ? sheetZestawy.getRange(2, COL_UPTODATE + 1, zestData.length - 1, 1).getValues()
    : [];
  
  // Zbierz checkboxy ponownie
  const elementCheckboxes = {};
  
  for (let i = 1; i < modData.length; i++) {
    const elementName = String(modData[i][COL_ELEMENT]).trim();
    if (!elementName || isModule(elementName)) continue;
    
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
  
  for (let i = 1; i < zestData.length; i++) {
    const elementName = String(zestData[i][COL_ELEMENT]).trim();
    if (!elementName || isModule(elementName)) continue;
    
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
  
  // Sprawdź checkboxy
  const checkboxInconsistencies = [];
  let autoFilledCheckboxCount = 0;
  let skippedTypedColumns = 0;
  
  for (const [elementName, entries] of Object.entries(elementCheckboxes)) {
    const validEntries = entries.filter(e => e.checked !== null);
    
    if (validEntries.length === 0) continue;
    
    const uniqueStates = [...new Set(validEntries.map(e => e.checked))];
    
    if (uniqueStates.length === 1) {
      const validState = uniqueStates[0];
      for (const e of entries) {
        if (e.checked === null) {
          const sheet = sheetsCache[e.sheet];
          const success = setCheckboxValue(sheet, e.row, e.col, validState);
          if (success) {
            autoFilledCheckboxCount++;
          } else {
            skippedTypedColumns++;
          }
        }
      }
    } else {
      checkboxInconsistencies.push({ name: elementName, uniqueStates, entries: validEntries });
    }
  }
  
  // Pobierz informacje o brakujących linkach z pierwotnego przetwarzania
  const autoFilledCount = parseInt(props.getProperty('autoFilledCount') || '0');
  const totalElements = parseInt(props.getProperty('totalElements') || '0');
  const missingLinksJson = props.getProperty('missingLinks') || '[]';
  const missingLinks = JSON.parse(missingLinksJson);
  
  // Zapisz dane o checkboxach do przetworzenia
  if (checkboxInconsistencies.length > 0) {
    // Zapisz dane konfliktów checkboxów
    props.setProperty('checkboxConflicts', JSON.stringify({
      conflicts: checkboxInconsistencies,
      COL_UPTODATE_LETTER: COL_UPTODATE_LETTER,
      elementCheckboxes: elementCheckboxes,
      sheetsCache: Object.keys(sheetsCache)
    }));
    props.setProperty('checkboxConflictIndex', '0');
    props.setProperty('checkboxAutoFilledCount', String(autoFilledCheckboxCount));
    props.setProperty('checkboxSkippedTypedColumns', String(skippedTypedColumns));
    
    // Pokaż pierwszy konflikt checkboxów
    showNextCheckboxConflict();
    return;
  }
  
  // Aktualizuj moduły
  const updatedModules = updateModuleCheckboxesInZestawy(
    modData,
    allModCheckboxValues,
    zestData,
    sheetsCache,
    COL_ELEMENT,
    COL_UPTODATE,
    isModule
  );
  
  // Wyczyść cache
  props.deleteProperty('ucancamConflicts');
  props.deleteProperty('ucancamConflictIndex');
  props.deleteProperty('autoFilledCount');
  props.deleteProperty('totalElements');
  props.deleteProperty('missingLinks');
  props.deleteProperty('checkboxConflicts');
  props.deleteProperty('checkboxConflictIndex');
  props.deleteProperty('checkboxAutoFilledCount');
  props.deleteProperty('checkboxSkippedTypedColumns');
  
  // Pokaż końcowe podsumowanie
  showFinalSummary(
    totalElements,
    autoFilledCount,
    fixedConflicts,
    0,
    autoFilledCheckboxCount,
    fixedCheckboxConflicts,
    0,
    skippedTypedColumns,
    updatedModules,
    missingLinks
  );
}

/**
 * Pokazuje HTML dialog z następnym konfliktem checkboxów.
 * Jeśli nie ma więcej konfliktów, aktualizuje moduły i pokazuje podsumowanie.
 */
function showNextCheckboxConflict() {
  const props = PropertiesService.getScriptProperties();
  const checkboxData = JSON.parse(props.getProperty('checkboxConflicts'));
  const currentIndex = parseInt(props.getProperty('checkboxConflictIndex') || '0');
  
  if (currentIndex >= checkboxData.conflicts.length) {
    // Wszystkie konflikty checkboxów rozwiązane - finalizuj
    finalizeCheckboxConflicts();
    return;
  }
  
  const conflict = checkboxData.conflicts[currentIndex];
  const COL_UPTODATE_LETTER = checkboxData.COL_UPTODATE_LETTER;
  
  // Przygotuj dane dla HTML
  const locations = conflict.entries.map(e => ({
    sheet: e.sheet,
    row: e.row,
    column: COL_UPTODATE_LETTER,
    checked: e.checked
  }));
  
  // Stwórz HTML z template
  const template = HtmlService.createTemplateFromFile('checkboxConflictResolver');
  template.elementName = conflict.name;
  template.locations = locations;
  template.currentConflict = currentIndex + 1;
  template.totalConflicts = checkboxData.conflicts.length;
  
  const html = template.evaluate()
    .setWidth(500)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(html, `Konflikt checkboxów (${currentIndex + 1}/${checkboxData.conflicts.length})`);
}

/**
 * Funkcja wywoływana z HTML dialog po wyborze stanu checkboxa.
 * Aplikuje wybrany stan i pokazuje następny konflikt.
 */
function resolveCheckboxConflict(elementName, selectedState) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const checkboxData = JSON.parse(props.getProperty('checkboxConflicts'));
  const currentIndex = parseInt(props.getProperty('checkboxConflictIndex') || '0');
  
  const sheetsCache = {
    'Moduły CNC': ss.getSheetByName('Moduły CNC'),
    'Zestawy CNC': ss.getSheetByName('Zestawy CNC')
  };
  
  const elementCheckboxes = checkboxData.elementCheckboxes;
  const allEntries = elementCheckboxes[elementName] || [];
  
  let skippedTypedColumns = parseInt(props.getProperty('checkboxSkippedTypedColumns') || '0');
  let successCount = 0;
  
  // Aplikuj wybrany stan do wszystkich wystąpień
  for (const e of allEntries) {
    const sheet = sheetsCache[e.sheet];
    try {
      const cell = sheet.getRange(e.row, e.col);
      cell.setValue(Boolean(selectedState));
      successCount++;
    } catch (error) {
      if (String(error.message).includes('typed column') || 
          String(error.message).includes('not allowed')) {
        skippedTypedColumns++;
      } else {
        try {
          const dataValidation = cell.getDataValidation();
          if (!dataValidation) {
            const rule = SpreadsheetApp.newDataValidation()
              .requireCheckbox()
              .setAllowInvalid(false)
              .build();
            cell.setDataValidation(rule);
          }
          cell.setValue(Boolean(selectedState));
          successCount++;
        } catch (e2) {
          skippedTypedColumns++;
        }
      }
    }
  }
  
  // Zaktualizuj licznik pominiętych typed columns
  props.setProperty('checkboxSkippedTypedColumns', String(skippedTypedColumns));
  
  Logger.log(`Rozwiązano konflikt checkboxa ${currentIndex + 1}: ${elementName} -> ${selectedState ? 'TRUE' : 'FALSE'} (${successCount} aktualizacji)`);
  
  // Przejdź do następnego konfliktu
  props.setProperty('checkboxConflictIndex', String(currentIndex + 1));
  showNextCheckboxConflict();
}

/**
 * Finalizuje przetwarzanie po rozwiązaniu wszystkich konfliktów checkboxów.
 */
function finalizeCheckboxConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  const sheetModuly = ss.getSheetByName('Moduły CNC');
  const sheetZestawy = ss.getSheetByName('Zestawy CNC');
  
  const COL_ELEMENT = 1;
  const COL_UPTODATE = 6;
  
  const sheetsCache = {
    'Moduły CNC': sheetModuly,
    'Zestawy CNC': sheetZestawy
  };
  
  function isModule(name) {
    if (!name) return false;
    return /^[MX]/i.test(String(name).trim());
  }
  
  // Pobierz dane
  const modData = sheetModuly.getDataRange().getValues();
  const zestData = sheetZestawy.getDataRange().getValues();
  const allModCheckboxValues = modData.length > 1 
    ? sheetModuly.getRange(2, COL_UPTODATE + 1, modData.length - 1, 1).getValues()
    : [];
  
  // Aktualizuj moduły
  const updatedModules = updateModuleCheckboxesInZestawy(
    modData,
    allModCheckboxValues,
    zestData,
    sheetsCache,
    COL_ELEMENT,
    COL_UPTODATE,
    isModule
  );
  
  // Pobierz wszystkie statystyki
  const checkboxData = JSON.parse(props.getProperty('checkboxConflicts'));
  const fixedCheckboxConflicts = checkboxData.conflicts.length;
  const autoFilledCheckboxCount = parseInt(props.getProperty('checkboxAutoFilledCount') || '0');
  const skippedTypedColumns = parseInt(props.getProperty('checkboxSkippedTypedColumns') || '0');
  
  const autoFilledCount = parseInt(props.getProperty('autoFilledCount') || '0');
  const totalElements = parseInt(props.getProperty('totalElements') || '0');
  const missingLinksJson = props.getProperty('missingLinks') || '[]';
  const missingLinks = JSON.parse(missingLinksJson);
  
  // Pobierz liczbę naprawionych konfliktów UCANCAM
  const ucancamData = JSON.parse(props.getProperty('ucancamConflicts') || '{"conflicts":[]}');
  const fixedConflicts = ucancamData.conflicts.length;
  
  // Wyczyść cache
  props.deleteProperty('ucancamConflicts');
  props.deleteProperty('ucancamConflictIndex');
  props.deleteProperty('autoFilledCount');
  props.deleteProperty('totalElements');
  props.deleteProperty('missingLinks');
  props.deleteProperty('checkboxConflicts');
  props.deleteProperty('checkboxConflictIndex');
  props.deleteProperty('checkboxAutoFilledCount');
  props.deleteProperty('checkboxSkippedTypedColumns');
  
  // Pokaż końcowe podsumowanie
  showFinalSummary(
    totalElements,
    autoFilledCount,
    fixedConflicts,
    0,
    autoFilledCheckboxCount,
    fixedCheckboxConflicts,
    0,
    skippedTypedColumns,
    updatedModules,
    missingLinks
  );
}

/**
 * Pokazuje końcowe podsumowanie sprawdzenia UCANCAM i UP-TO-DATE.
 */
function showFinalSummary(
  totalElements,
  autoFilledCount,
  fixedConflicts,
  unresolvedConflicts,
  autoFilledCheckboxCount,
  fixedCheckboxConflicts,
  unresolvedCheckboxConflicts,
  skippedTypedColumns,
  updatedModules,
  missingLinks
) {
  const ui = SpreadsheetApp.getUi();
  const summaryLines = [];
  
  summaryLines.push(`📊 Sprawdzono ${totalElements} elementów.`);
  summaryLines.push('');
  
  // Podsumowanie UCANCAM
  summaryLines.push(`📦 LINKI UCANCAM (kolumna F):`);
  if (autoFilledCount > 0) {
    summaryLines.push(`  ✅ Automatycznie uzupełniono ${autoFilledCount} pustych komórek.`);
  }
  if (fixedConflicts > 0) {
    summaryLines.push(`  🔧 Naprawiono konflikty dla ${fixedConflicts} elementów.`);
  }
  if (autoFilledCount === 0 && fixedConflicts === 0 && unresolvedConflicts === 0) {
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
  if (autoFilledCheckboxCount === 0 && fixedCheckboxConflicts === 0 && unresolvedCheckboxConflicts === 0 && skippedTypedColumns === 0) {
    summaryLines.push(`  ✅ Wszystkie checkboxy są spójne.`);
  }
  
  if (updatedModules > 0) {
    summaryLines.push('');
    summaryLines.push(`🔄 MODUŁY W ZESTAWACH CNC:`);
    summaryLines.push(`  ✅ Zaktualizowano checkboxy dla ${updatedModules} modułów.`);
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

/**
 * Aktualizuje checkboxy modułów w arkuszu "Zestawy CNC" na podstawie statusu ich elementów.
 * Moduł ma checkbox TRUE tylko wtedy, gdy WSZYSTKIE jego elementy w "Moduły CNC" mają TRUE.
 */
function updateModuleCheckboxesInZestawy(modData, allModCheckboxValues, zestData, sheetsCache, COL_ELEMENT, COL_UPTODATE, isModule) {
  // 1. Zbuduj mapę: moduł -> status (czy wszystkie elementy są TRUE)
  const moduleStatus = {}; // { "M1594": { allTrue: true, elementCount: 3 }, ... }
  
  // Przejdź przez wszystkie wiersze w "Moduły CNC"
  for (let i = 1; i < modData.length; i++) {
    const moduleName = String(modData[i][0]).trim(); // Kolumna A - Nr modułu
    const elementName = String(modData[i][COL_ELEMENT]).trim(); // Kolumna B - element
    
    // Pomijamy puste wiersze i sprawdzamy tylko moduły
    if (!moduleName || !isModule(moduleName)) continue;
    if (!elementName || isModule(elementName)) continue; // Pomijamy inne moduły w kolumnie B
    
    // Pobierz wartość checkboxa elementu (już mamy w allModCheckboxValues)
    const value = allModCheckboxValues[i - 1][0];
    let checked = false;
    
    if (typeof value === 'boolean') {
      checked = value;
    } else if (typeof value === 'string') {
      const upperValue = value.toUpperCase().trim();
      checked = (upperValue === 'TRUE');
    }
    
    // Inicjalizuj status modułu jeśli jeszcze nie istnieje
    if (!moduleStatus[moduleName]) {
      moduleStatus[moduleName] = { allTrue: true, elementCount: 0 };
    }
    
    moduleStatus[moduleName].elementCount++;
    
    // Jeśli choć jeden element jest FALSE, cały moduł = FALSE
    if (!checked) {
      moduleStatus[moduleName].allTrue = false;
    }
  }
  
  // 2. Zaktualizuj checkboxy modułów w arkuszu "Zestawy CNC"
  const zestSheet = sheetsCache['Zestawy CNC'];
  let updatedCount = 0;
  
  for (let i = 1; i < zestData.length; i++) {
    const cellValue = String(zestData[i][COL_ELEMENT]).trim(); // Kolumna B
    
    // Sprawdź czy to moduł i czy mamy dla niego status
    if (isModule(cellValue) && moduleStatus[cellValue]) {
      const shouldBeTrue = moduleStatus[cellValue].allTrue;
      const currentValue = zestData[i][COL_UPTODATE]; // Kolumna G (0-based = 6)
      
      // Konwersja currentValue na boolean
      let currentBool = false;
      if (typeof currentValue === 'boolean') {
        currentBool = currentValue;
      } else if (typeof currentValue === 'string') {
        currentBool = (currentValue.toUpperCase().trim() === 'TRUE');
      }
      
      // Zaktualizuj tylko jeśli wartość się zmienia
      if (currentBool !== shouldBeTrue) {
        try {
          zestSheet.getRange(i + 1, COL_UPTODATE + 1).setValue(shouldBeTrue);
          updatedCount++;
        } catch (e) {
          Logger.log(`Nie można zaktualizować modułu ${cellValue} w wierszu ${i + 1}: ${e.message}`);
        }
      }
    }
  }
  
  return updatedCount;
}
