/**
 * Lokalne zmiany witam witam pozdrawiam serdecznie
 * Rejestr Plików CNC - pobieranie DXF
 * Wklej ten kod do Extensions -> Apps Script.
 *
 * Struktura arkuszy:
 *  - "Dane" (nieużywany bezpośrednio tutaj, pozostawiony)
 *  - "Zestawy CNC"  kol A = Nr zestawu, kol B = Nr kat. elementu (może być moduł lub element z hiperlinkiem)
 *  - "Moduły CNC"   kol A = Nr Moduły, kol B = Nr kat. elementu (elementy tworzące moduł; komórki B mogą zawierać hiperłącza)
 *
 * Zasada:
 *  - moduł: nazwa zaczyna się od "M" lub "X" (wielkie litery)
 *  - element: wszystko inne → musi mieć hiperłącze w komórce
 *
 * Pliki zapisywane są do folderu na Twoim Dysku Google.
 */

/* ========== KONFIGURACJA (opcjonalna) ========== */
const SHEET_ZESTAWY = 'Zestawy CNC';
const SHEET_MODULE = 'Moduły CNC';
/* =============================================== */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CNC')
    //.addItem('Pobierz pliki dla zestawu...', 'promptAndDownload')
    .addItem('Pobierz pliki dla zestawu (z kolorami)...', 'promptAndDownloadWithColors') // 🆕 nowa opcja
    .addItem('Pobierz pliki dla modułu...', 'promptAndDownloadModule')
    .addToUi();

  ui.createMenu('Sync')
    .addItem('Ustaw / edytuj link dla elementu...', 'promptAndSyncLink')
    .addItem('Porównaj linki (SyncLinks)', 'promptAndCompareLinks')
     .addItem('Masowe sprawdzenie linków', 'massCheckAndFixLinks')
    .addToUi();
}

function promptAndCompareLinks() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const elementNameResponse = ui.prompt('Porównaj linki', 'Podaj nazwę elementu:', ui.ButtonSet.OK_CANCEL);
  if (elementNameResponse.getSelectedButton() !== ui.Button.OK) return;

  const elementName = elementNameResponse.getResponseText().trim();
  if (!elementName) return ui.alert('Nie podano nazwy elementu.');

  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC', 'Elementy CNC'];
  const foundLinks = []; // {sheet, row, link, ok, status}

  function checkDriveLinkStatus(link) {
    if (!link || link === '(brak linku)') return { ok: false, code: 0, status: 'Brak linku' };
    const match = link.match(/[-\w]{25,}/);
    if (!match) return { ok: false, code: 0, status: 'Nieprawidłowy format linku' };

    const fileId = match[0];
    try {
      const file = DriveApp.getFileById(fileId);
      file.getName();
      return { ok: true, code: 200, status: 'OK' };
    } catch (e) {
      const msg = String(e);
      if (msg.includes('File not found')) return { ok: false, code: 404, status: 'Nie znaleziono pliku' };
      if (msg.includes('User does not have permission')) return { ok: false, code: 403, status: 'Brak dostępu' };
      return { ok: false, code: 500, status: 'Błąd: ' + e.message };
    }
  }

  // Zbierz wszystkie linki
  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const range = sheet.getDataRange();
    const values = range.getValues();
    const richValues = range.getRichTextValues();

    for (let r = 0; r < values.length; r++) {
      const name = String(values[r][1]).trim(); // kolumna B
      if (name === elementName) {
        let link = null;
        try {
          link = richValues[r][1]?.getLinkUrl() || '(brak linku)';
        } catch (e) {
          link = '(brak linku)';
        }

        const status = checkDriveLinkStatus(link);
        foundLinks.push({ sheet: sheetName, row: r + 1, link, ...status });
      }
    }
  }

  if (foundLinks.length === 0) return ui.alert(`Nie znaleziono elementu "${elementName}".`);

    // --- 1️⃣ Sprawdź błędne linki ---
  const validLinks = foundLinks.filter(f => f.ok);
  const invalidLinks = foundLinks.filter(f => !f.ok);

  // Jeśli są tylko "braki linków", ale przynajmniej jeden link poprawny – traktuj jak różne linki
  const hasValid = validLinks.length > 0;
  const onlyMissing = invalidLinks.every(f => f.status === 'Brak linku');

  if (invalidLinks.length > 0 && !(hasValid && onlyMissing)) {
    // czyli: występują błędy inne niż "brak linku" (np. 404/403)
    const msg = invalidLinks.map(f => `• ${f.sheet}!B${f.row} → ${f.status}`).join('\n');
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

  // --- 2️⃣ Wszystkie linki poprawne LUB brak linku w części z nich ---
  const allLinks = foundLinks.map(f => f.link).filter(l => l && l !== '(brak linku)');
  const uniqueLinks = [...new Set(allLinks)];

  if (uniqueLinks.length === 0) {
    ui.alert(`❌ Brak jakichkolwiek linków dla "${elementName}".`);
    return;
  }

  if (uniqueLinks.length > 1 || (hasValid && onlyMissing)) {
    // Występują różne linki lub część komórek nie ma linku – zapytaj użytkownika, który zachować
    let msg = `Znaleziono ${uniqueLinks.length} różne linki (lub brak w niektórych miejscach) dla "${elementName}":\n\n`;
    uniqueLinks.forEach((l, i) => {
      msg += `${i + 1}. ${l}\n`;
    });
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

// --- Pomocnicza funkcja do aktualizacji linków ---
function updateLinks(ss, foundLinks, elementName, newLink, ui) {
  let updatedCount = 0;

  for (const f of foundLinks) {
    const sheet = ss.getSheetByName(f.sheet);
    if (!sheet) continue;
    const cell = sheet.getRange(f.row, 2);
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



function promptAndDownloadWithColors() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz pliki DXF z kolorami', 'Podaj numer zestawu (np. P1608):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const setId = resp.getResponseText().trim();
  if (!setId) {
    ui.alert('Nie podano numeru zestawu.');
    return;
  }

  try {
    downloadSetFilesWithColors(setId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function promptAndDownload() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz pliki DXF', 'Podaj numer zestawu (np. P1608):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const setId = resp.getResponseText().trim();
  if (!setId) {
    ui.alert('Nie podano numeru zestawu.');
    return;
  }
  try {
    downloadSetFiles(setId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

function promptAndDownloadModule() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Pobierz pliki DXF', 'Podaj numer modułu (np. M1594):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const modId = resp.getResponseText().trim();
  if (!modId) {
    ui.alert('Nie podano numeru modułu.');
    return;
  }
  try {
    downloadModuleFiles(modId);
  } catch (e) {
    ui.alert('Błąd', 'Wystąpił błąd: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Pobiera pliki dla zestawu z podziałem na kolory (nazwy folderów)
 */
function downloadSetFilesWithColors(setId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetZest = ss.getSheetByName(SHEET_ZESTAWY);
  const sheetMod = ss.getSheetByName(SHEET_MODULE);
  const ui = SpreadsheetApp.getUi();

  if (!sheetZest || !sheetMod) {
    ui.alert('Błąd', `Brakuje arkuszy "${SHEET_ZESTAWY}" lub "${SHEET_MODULE}".`);
    return;
  }

  const zestValues = sheetZest.getDataRange().getValues();
  const zestRich = sheetZest.getDataRange().getRichTextValues();
  const modValues = sheetMod.getDataRange().getValues();
  const modRich = sheetMod.getDataRange().getRichTextValues();

  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

  const startElements = zestawyMap[setId];
  if (!startElements || startElements.length === 0) {
    ui.alert('Nie znaleziono zestawu', `Brak wierszy o Nr zestawu = "${setId}".`);
    return;
  }

  // 🔍 Zbierz listę wszystkich unikalnych elementów (rekurencyjnie)
  const elements = [];
  const collect = (list) => {
    for (const e of list) {
      const n = e.text.trim();
      if (!isModuleName(n)) {
        if (!elements.find(el => el.text === n)) elements.push(e);
      } else {
        const sub = modulesMap[n];
        if (sub && sub.length) collect(sub);
      }
    }
  };
  collect(startElements);

  if (elements.length === 0) {
    ui.alert('Brak elementów do pobrania.');
    return;
  }

  // 🧩 Przygotuj dane do HTML
  const htmlTemplate = HtmlService.createTemplateFromFile('colorSelector');
  htmlTemplate.data = elements;
  htmlTemplate.setId = setId;

  const htmlOutput = htmlTemplate.evaluate()
    .setTitle(`Kolory dla ${setId}`)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Kolory dla ${setId}`);
}


function downloadSetFiles(setId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Pobierz dane z arkuszy (values i rich text dla kolumn z linkami)
  const sheetZest = ss.getSheetByName(SHEET_ZESTAWY);
  const sheetMod = ss.getSheetByName(SHEET_MODULE);
  if (!sheetZest || !sheetMod) {
    ui.alert('Błąd', `Brakuje wymaganych arkuszy: "${SHEET_ZESTAWY}" lub "${SHEET_MODULE}".`, ui.ButtonSet.OK);
    return;
  }

  const zestValues = sheetZest.getDataRange().getValues(); // pełne wiersze
  const zestRich = sheetZest.getDataRange().getRichTextValues();
  const modValues = sheetMod.getDataRange().getValues();
  const modRich = sheetMod.getDataRange().getRichTextValues();

  // Tworzymy mapy (bez sprawdzania pustych wierszy)
const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

// Ostrzeżenia o pustych komórkach tylko dla bieżącego zestawu
const dataWarnings = [];
for (let r = 0; r < zestValues.length; r++) {
  const rowSet = String(zestValues[r][0]).trim();
  const rowElement = String(zestValues[r][1]).trim();
  if (rowSet === setId && !rowElement) {
    const rowNumber = r + 1;
    const colLetter = 'B';
    dataWarnings.push(`Brak numeru elementu w arkuszu "${SHEET_ZESTAWY}" dla zestawu "${setId}" w komórce ${colLetter}${rowNumber}`);
  }
}

  // Sprawdź czy zestaw istnieje
  const startElements = zestawyMap[setId];
  if (!startElements || startElements.length === 0) {
    ui.alert('Nie znaleziono zestawu', `Nie znaleziono wierszy o Nr zestawu = "${setId}" w arkuszu "${SHEET_ZESTAWY}".`, ui.ButtonSet.OK);
    return;
  }

  // Utwórz folder docelowy na Dysku
  const folderName = `Rejestr Plików CNC - Pobrania ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl(); // <-- LINK DO FOLDERU

  // Rekurencyjne przetwarzanie
  const visited = {};
  const missingLinks = [];
  const downloaded = [];
  const errors = [];

  for (let e of startElements) {
    processElementRecursive(e.text, e.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors);
  }

  // Podsumowanie
  const summaryLines = [];
  summaryLines.push(`📁 Utworzony folder:`);
  summaryLines.push(folderUrl);
  summaryLines.push('');
  summaryLines.push(`Pobrano plików: ${downloaded.length}`);
  if (downloaded.length) {
    downloaded.slice(0, 20).forEach(d => {
  const surfaceStr = d.surface ? ` (${d.surface.toFixed(3)} m²)` : '';
  const pretty = d.prettyName ? ` – ${d.prettyName}` : '';
  summaryLines.push(`• ${d.name}${pretty}${surfaceStr}`);
});
    if (downloaded.length > 20) summaryLines.push(`... + ${downloaded.length - 20} innych`);
  }
  if (missingLinks.length) {
    summaryLines.push('');
    summaryLines.push(`Elementy bez hiperłącza (${missingLinks.length}):`);
    missingLinks.forEach(m => summaryLines.push(`• ${m}`));
  }
  if (errors.length) {
    summaryLines.push('');
    summaryLines.push(`Błędy (${errors.length}):`);
    errors.forEach(err => summaryLines.push(`• ${err}`));
  }
  if (dataWarnings.length) {
  summaryLines.push('');
  summaryLines.push(`⚠️ Ostrzeżenia dotyczące danych (${dataWarnings.length}):`);
  dataWarnings.forEach(w => summaryLines.push(`• ${w}`));
}

  ui.alert('Pobieranie zakończone', summaryLines.join('\n'), ui.ButtonSet.OK);
}

function downloadModuleFiles(modId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheetZest = ss.getSheetByName(SHEET_ZESTAWY);
  const sheetMod = ss.getSheetByName(SHEET_MODULE);
  if (!sheetZest || !sheetMod) {
    ui.alert('Błąd', `Brakuje wymaganych arkuszy: "${SHEET_ZESTAWY}" lub "${SHEET_MODULE}".`, ui.ButtonSet.OK);
    return;
  }

  const zestValues = sheetZest.getDataRange().getValues();
  const zestRich = sheetZest.getDataRange().getRichTextValues();
  const modValues = sheetMod.getDataRange().getValues();
  const modRich = sheetMod.getDataRange().getRichTextValues();

  // mapy
  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

  // ostrzeżenia dla pustych komórek tylko dla tego modułu
  const dataWarnings = [];
  for (let r = 0; r < modValues.length; r++) {
    const rowMod = String(modValues[r][0]).trim();
    const rowElement = String(modValues[r][1]).trim();
    if (rowMod === modId && !rowElement) {
      const rowNumber = r + 1;
      const colLetter = 'B';
      dataWarnings.push(`Brak numeru elementu w arkuszu "${SHEET_MODULE}" dla modułu "${modId}" w komórce ${colLetter}${rowNumber}`);
    }
  }

  // sprawdź czy moduł istnieje
  const startElements = modulesMap[modId];
  if (!startElements || startElements.length === 0) {
    ui.alert('Nie znaleziono modułu', `Nie znaleziono wierszy o Nr modułu = "${modId}" w arkuszu "${SHEET_MODULE}".`, ui.ButtonSet.OK);
    return;
  }

  // folder docelowy
  const folderName = `Rejestr Plików CNC - Moduł ${modId} - ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl();

  // proces rekurencyjny
  const visited = {};
  const missingLinks = [];
  const downloaded = [];
  const errors = [];

  for (let e of startElements) {
    processElementRecursive(e.text, e.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors);
  }

  // podsumowanie
  const summaryLines = [];
  summaryLines.push(`📁 Utworzony folder:`);
  summaryLines.push(folderUrl);
  summaryLines.push('');
  summaryLines.push(`Pobrano plików: ${downloaded.length}`);
  if (downloaded.length) {
    downloaded.slice(0, 20).forEach(d => {
      const surfaceStr = d.surface ? ` (${d.surface.toFixed(3)} m²)` : '';
      const pretty = d.prettyName ? ` – ${d.prettyName}` : '';
      summaryLines.push(`• ${d.name}${pretty}${surfaceStr}`);
    });
    if (downloaded.length > 20) summaryLines.push(`... + ${downloaded.length - 20} innych`);
  }
  if (missingLinks.length) {
    summaryLines.push('');
    summaryLines.push(`Elementy bez hiperłącza (${missingLinks.length}):`);
    missingLinks.forEach(m => summaryLines.push(`• ${m}`));
  }
  if (errors.length) {
    summaryLines.push('');
    summaryLines.push(`Błędy (${errors.length}):`);
    errors.forEach(err => summaryLines.push(`• ${err}`));
  }
  if (dataWarnings.length) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Ostrzeżenia dotyczące danych (${dataWarnings.length}):`);
    dataWarnings.forEach(w => summaryLines.push(`• ${w}`));
  }

  ui.alert('Pobieranie zakończone', summaryLines.join('\n'), ui.ButtonSet.OK);
}

function promptAndSyncLink() {
  const ui = SpreadsheetApp.getUi();

  // 1. Zapytaj o numer elementu
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

  // 2. Zapytaj o link
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

  // 3. Uruchom synchronizację
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
    const richValues = dataRange.getRichTextValues();

    // Zakładamy, że kolumna B (index 1) to Nr kat. elementu
    for (let r = 0; r < values.length; r++) {
      const cellValue = String(values[r][1]).trim();
      if (cellValue === elementName) {
        // Zbuduj nowy RichTextValue z linkiem
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(cellValue)
          .setLinkUrl(link)
          .build();

        sheet.getRange(r + 1, 2).setRichTextValue(richText);
        totalUpdated++;
      }
    }
  }

  return totalUpdated;
}

/* ====== helper functions ====== */

/**
 * Tworzy mapę: map[key] -> array obiektów:
 * {
 *   text: Nr kat. elementu,
 *   richLink: hyperlink (jeśli istnieje),
 *   surface: kol. E (m^2),
 *   name: kol. G (nazwa elementu)
 * }
 * Nie sprawdza pustych wierszy — walidacja odbywa się tylko w downloadSetFiles() dla bieżącego zestawu.
 */
function buildMapForSheet(values, richValues, idxKeyCol, idxDataCol, sheetName) {
  const map = {};

  for (let r = 0; r < values.length; r++) {
    const key = String(values[r][idxKeyCol]).trim();
    const dataText = String(values[r][idxDataCol]).trim();

    if (!key || !dataText) continue;

    let richLink = null;
    try {
      const richCell = richValues[r][idxDataCol];
      if (richCell && typeof richCell.getLinkUrl === 'function') {
        richLink = richCell.getLinkUrl();
      }
    } catch (e) {
      richLink = null;
    }

    const surface = values[r][4] ? Number(values[r][4]) : null;  // kol. E
    const name = values[r][6] ? String(values[r][6]).trim() : ''; // kol. G

    // **Nowe**: pobieramy ilość z kolumny C (index 2). Jeśli pusta/nieprawidłowa -> 1
    let count = 1;
    try {
      const raw = values[r][2];
      if (raw !== '' && raw !== null && raw !== undefined) {
        const num = Number(raw);
        if (!isNaN(num) && num > 0) count = num;
      }
    } catch (e) {
      count = 1;
    }

    if (!map[key]) map[key] = [];
    map[key].push({ text: dataText, richLink: richLink, surface: surface, name: name, count: count });
  }

  return { map };
}


/** 
 * Rozpoznaje, czy nazwa to moduł: zaczyna się od M lub X (wielkie lub małe rozważymy)
 */
function isModuleName(name) {
  if (!name) return false;
  name = String(name).trim();
  return /^[MX]/i.test(name);
}

/**
 * Jak processElementRecursive, ale z obsługą kolorów.
 */
function processElementRecursiveWithColor(name, providedRichLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap, multiplier = 1) {
  name = String(name).trim();
  if (!name) return;

  // blokujemy tylko moduły - elementy mogą wystąpić wielokrotnie i muszą być zliczone
  if (isModuleName(name)) {
    if (visited[name]) return;
    visited[name] = true;

    const children = modulesMap[name];
    if (!children || children.length === 0) {
      missingLinks.push(`Moduł ${name} - brak wpisów w "${SHEET_MODULE}"`);
      return;
    }

    for (let ch of children) {
      // ch.count pochodzi z kolumny C w arkuszu modułów
      const childMultiplier = multiplier * (ch.count || 1);
      processElementRecursiveWithColor(ch.text, ch.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap, childMultiplier);
    }
    return;
  }

  // --- element końcowy ---
  const link = providedRichLink || findLinkForElement(name, modulesMap, zestawyMap);
  if (!link) {
    missingLinks.push(name);
    return;
  }

  try {
    // pobierz dane elementu (nazwa, surface, etc.)
    const elementData = findElementData(name, modulesMap, zestawyMap);
    const elementCountFromSheet = (function() {
      // Spróbuj znaleźć count dla tego elementu w modulesMap/zestawyMap
      for (let key in modulesMap) {
        for (const e of modulesMap[key]) {
          if (e.text === name) return (e.count || 1);
        }
      }
      for (let key in zestawyMap) {
        for (const e of zestawyMap[key]) {
          if (e.text === name) return (e.count || 1);
        }
      }
      return 1;
    })();

    const effectiveCount = multiplier * elementCountFromSheet;

    // przygotuj link do pobrania
    const fileIdMatch = link.match(/[-\w]{25,}/);
    const directLink = fileIdMatch ? `https://drive.google.com/uc?export=download&id=${fileIdMatch[0]}` : link;
    const resp = UrlFetchApp.fetch(directLink, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      errors.push(`${name}: błąd HTTP ${code} przy pobieraniu ${directLink}`);
      return;
    }

    const blob = resp.getBlob();
    const fileName = sanitizeFileName(name) + '.dxf';
    blob.setName(fileName);

    // 🟢 Kolor (folder docelowy) - tworzymy plik tylko raz na dany element
    const color = colorMap[name] || 'Bez koloru';
    const colorFolder = getOrCreateSubfolder(folder, color);

    // Sprawdź czy już mamy wpis w downloaded dla tego elementu+koloru
    const existing = downloaded.find(d => d.name === name && d.color === color);

    if (!existing) {
      // utwórz plik na dysku (raz)
      colorFolder.createFile(blob);
      downloaded.push({
        name: name,
        color: color,
        prettyName: elementData?.name || '',
        count: effectiveCount
      });
    } else {
      // nie tworzymy ponownie pliku — tylko zwiększamy liczbę
      existing.count = (existing.count || 0) + effectiveCount;
    }

  } catch (e) {
    errors.push(`${name}: ${e.message}`);
  }
}

/**
 * Nowa wersja startDownloadWithColors:
 * - najpierw zbiera totals mapę element->ilość (uwzględniając kol. C),
 * - potem pobiera pliki raz na element+kolor i zapisuje downloaded z .count = totals[element]
 */
function startDownloadWithColors(colorMap, setId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetZest = ss.getSheetByName(SHEET_ZESTAWY);
  const sheetMod = ss.getSheetByName(SHEET_MODULE);

  const zestValues = sheetZest.getDataRange().getValues();
  const zestRich = sheetZest.getDataRange().getRichTextValues();
  const modValues = sheetMod.getDataRange().getValues();
  const modRich = sheetMod.getDataRange().getRichTextValues();

  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

  const startElements = zestawyMap[setId];
  if (!startElements) {
    ui.alert('Nie znaleziono zestawu ' + setId);
    return;
  }

  // 1) Zbierz totals: element -> totalCount
  const totals = {}; // { elementName: totalCount }
  const pathVisited = {}; // będzie używane w rekurencji (dla bieżącej ścieżki)

  for (let e of startElements) {
    const topCount = (typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    collectElementTotals(e.text, topCount, modulesMap, zestawyMap, totals, pathVisited);
  }

  // jeśli nie ma nic do pobrania
  const elementNames = Object.keys(totals);
  if (elementNames.length === 0) {
    ui.alert('Brak elementów do pobrania.');
    return;
  }

  // 2) Utwórz folder docelowy
  const folderName = `Rejestr Plików CNC - Pobrania ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl();

  const missingLinks = [];
  const downloaded = []; // { name, color, prettyName, count }
  const errors = [];

  // 3) Dla każdego unikalnego elementu pobierz plik raz (dla przypisanego koloru) i przypisz count z totals
  for (let name of elementNames) {
    const count = totals[name] || 0;
    const link = findLinkForElement(name, modulesMap, zestawyMap);
    if (!link) {
      missingLinks.push(name);
      continue;
    }

    try {
      const fileIdMatch = link.match(/[-\w]{25,}/);
      const directLink = fileIdMatch ? `https://drive.google.com/uc?export=download&id=${fileIdMatch[0]}` : link;
      const resp = UrlFetchApp.fetch(directLink, { muteHttpExceptions: true });
      const code = resp.getResponseCode();
      if (code < 200 || code >= 300) {
        errors.push(`${name}: błąd HTTP ${code} przy pobieraniu ${directLink}`);
        continue;
      }

      const blob = resp.getBlob();
      const fileName = sanitizeFileName(name) + '.dxf';
      blob.setName(fileName);

      // kolor z mapy z dialogu (jeśli brak -> "Bez koloru")
      const color = colorMap && colorMap[name] ? colorMap[name] : 'Bez koloru';
      const colorFolder = getOrCreateSubfolder(folder, color);
      colorFolder.createFile(blob);

      const elementData = findElementData(name, modulesMap, zestawyMap);

      downloaded.push({
        name: name,
        color: color,
        prettyName: elementData?.name || '',
        count: count
      });
    } catch (e) {
      errors.push(`${name}: ${e.message}`);
    }
  }

  // 4) Podsumowanie i plik TXT
  const summary = [];
  summary.push(`📁 Folder: ${folderUrl}`);
  summary.push(`Pobrano ${downloaded.length} plików.`);
  if (missingLinks.length) {
    summary.push('');
    summary.push('❌ Brak hiperłączy dla:');
    missingLinks.forEach(m => summary.push('• ' + m));
  }
  if (errors.length) {
    summary.push('');
    summary.push('⚠️ Błędy:');
    errors.forEach(e => summary.push('• ' + e));
  }

  const summaryFileUrl = createSummaryTxtFile(downloaded, folder);
  if (summaryFileUrl) {
    summary.push('');
    summary.push('📄 Utworzono plik podsumowania:');
    summary.push(summaryFileUrl);
  }

  ui.alert('Zakończono', summary.join('\n'), ui.ButtonSet.OK);
}


/**
 * Rekurencyjnie przetwarza element/moduł.
 * - jeśli element (nie moduł): szuka hiperłącza (jeśli przekazano richLink - używa) i pobiera plik
 * - jeśli moduł: rozkłada przez modulesMap (jeśli brak -> traktuj jako brak definicji)
 */
function processElementRecursive(name, providedRichLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, multiplier = 1) {
  name = String(name).trim();
  if (!name) return;

  if (isModuleName(name)) {
    if (visited[name]) return;
    visited[name] = true;

    const children = modulesMap[name];
    if (!children || children.length === 0) {
      missingLinks.push(`Moduł ${name} - brak wpisów w "${SHEET_MODULE}"`);
      return;
    }

    for (let ch of children) {
      const childMultiplier = multiplier * (ch.count || 1);
      processElementRecursive(ch.text, ch.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, childMultiplier);
    }
    return;
  }

  // element końcowy
  let link = providedRichLink || findLinkForElement(name, modulesMap, zestawyMap);
  if (!link) {
    missingLinks.push(name);
    return;
  }

  // znajdź dane o elemencie
  const elementData = findElementData(name, modulesMap, zestawyMap);
  const elementCountFromSheet = (function() {
    for (let key in modulesMap) {
      for (const e of modulesMap[key]) {
        if (e.text === name) return (e.count || 1);
      }
    }
    for (let key in zestawyMap) {
      for (const e of zestawyMap[key]) {
        if (e.text === name) return (e.count || 1);
      }
    }
    return 1;
  })();

  const effectiveCount = multiplier * elementCountFromSheet;

  try {
    let fileIdMatch = link.match(/[-\w]{25,}/);
    let directLink = link;
    if (fileIdMatch) {
      directLink = `https://drive.google.com/uc?export=download&id=${fileIdMatch[0]}`;
    }

    const resp = UrlFetchApp.fetch(directLink, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) {
      errors.push(`${name}: błąd HTTP ${code} przy pobieraniu ${directLink}`);
      return;
    }

    let blob = resp.getBlob();
    const fileName = sanitizeFileName(name) + '.dxf';
    blob.setName(fileName);

    // Tworzymy plik tylko raz; agregujemy count w downloaded
    const existing = downloaded.find(d => d.name === name);
    if (!existing) {
      folder.createFile(blob);
      downloaded.push({
        name: name,
        url: directLink,
        fileId: fileIdMatch ? fileIdMatch[0] : 'unknown',
        prettyName: elementData?.name || '',
        surface: elementData?.surface || null,
        count: effectiveCount
      });
    } else {
      existing.count = (existing.count || 0) + effectiveCount;
    }

  } catch (e) {
    errors.push(`${name}: ${e.message}`);
  }
}


/**
 * Próbuje znaleźć link dla elementu:
 * - przeszukuje moduły: wiersze w których kolumna B == name (i wykorzysta link z tej komórki)
 * - przeszukuje zestawy: wiersze w których kolumna B == name (i wykorzysta link z tej komórki)
 * Zwraca pierwszy znaleziony link albo null.
 */
const linkCache = {};
function findLinkForElement(name, modulesMap, zestawyMap) {
  if (linkCache[name]) return linkCache[name];
  for (let key in modulesMap) {
    for (let e of modulesMap[key]) {
      if (e.text === name && e.richLink) return (linkCache[name] = e.richLink);
    }
  }
  for (let key in zestawyMap) {
    for (let e of zestawyMap[key]) {
      if (e.text === name && e.richLink) return (linkCache[name] = e.richLink);
    }
  }
  return (linkCache[name] = null);
}


/**
 * Zwraca obiekt {name, surface} dla danego elementu
 */
function findElementData(elName, modulesMap, zestawyMap) {
  const searchIn = [modulesMap, zestawyMap];
  for (const map of searchIn) {
    for (const key in map) {
      const arr = map[key];
      for (const entry of arr) {
        if (entry.text === elName) {
          return { name: entry.name || '', surface: entry.surface || null };
        }
      }
    }
  }
  return null;
}

/** usuwa niebezpieczne znaki z nazwy pliku */
function sanitizeFileName(name) {
  return name.replace(/[\/\\\?\%\*\:\|\"<>\.]/g, '_').substring(0, 240);
}

/** prosty timestamp do nazwy folderu */
function timestampForName() {
  const d = new Date();
  const pad = (n) => (n<10?'0':'')+n;
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

/**
 * Pomocnicza: tworzy lub zwraca istniejący podfolder
 */
function getOrCreateSubfolder(parentFolder, subfolderName) {
  const existing = parentFolder.getFoldersByName(subfolderName);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(subfolderName);
}

/**
 * Tworzy plik TXT z podsumowaniem elementów.
 * @param {Array} downloaded - lista obiektów {name, color, namePretty, count}
 * @param {GoogleAppsScript.Drive.Folder} folder - folder, w którym zapisujemy
 */
/**
 * Szybsza wersja createSummaryTxtFile()
 * - używa StringBuildera (join tylko raz)
 * - unika pad() z powtarzaniem spacji dla każdej linii
 * - ogranicza sortowanie i formatowanie
 */
function createSummaryTxtFile(downloaded, folder) {
  if (!downloaded || downloaded.length === 0) return null;

  // Grupowanie po nazwie elementu + kolor
  const grouped = new Map();
  for (const item of downloaded) {
    const key = `${item.name}||${item.color || 'Bez koloru'}`;
    const g = grouped.get(key) || { count: 0, namePretty: item.prettyName || '', color: item.color || 'Bez koloru', name: item.name };
    g.count += (item.count || 1);
    grouped.set(key, g);
  }

  // Konwersja do tablicy i sortowanie alfabetyczne
  const sorted = Array.from(grouped.values()).sort((a, b) => a.name.localeCompare(b.name));

  // Użycie tablicy jako "string builder"
  const lines = [];
  lines.push('Nr kat. elementu       | Ilość | Nazwa elementu                       | Kolor');
  lines.push('------------------------+--------+-------------------------------------+------------');

  // Wydajne wyrównywanie za pomocą prostego paddingu (bez slice/repeat)
  const padRight = (txt, len) => (txt.length >= len ? txt.substring(0, len) : txt + ' '.repeat(len - txt.length));
  const padLeft = (txt, len) => (txt.length >= len ? txt.substring(0, len) : ' '.repeat(len - txt.length) + txt);

  for (const el of sorted) {
    lines.push(
      `${padRight(el.name, 23)}| ${padLeft(String(el.count), 6)}| ${padRight(el.namePretty, 37)}| ${el.color}`
    );
  }

  // Łączenie wszystkiego w jeden string JEDNYM joinem
  const content = lines.join('\n');

  // Tworzenie pliku tylko raz
  const blob = Utilities.newBlob(content, 'text/plain', 'Podsumowanie_elementów.txt');
  const file = folder.createFile(blob);

  return file.getUrl();
}




/**
 * Rekurencyjnie zbiera łączną ilość każdego elementu dla danej listy startowej.
 * - pathVisited blokuje moduł tylko na bieżącej ścieżce (zapobiega nieskończonej pętli)
 * - modulesMap: mapa moduł -> array entries {text, richLink, count, ...}
 */
function collectElementTotals(name, multiplier, modulesMap, zestawyMap, totals, pathVisited) {
  name = String(name).trim();
  if (!name) return;

  // jeśli to moduł -> rozpakuj jego dzieci
  if (isModuleName(name)) {
    // jeśli już na bieżącej ścieżce był ten moduł -> przerywamy (cykl)
    if (pathVisited[name]) {
      return;
    }
    pathVisited[name] = true;

    const children = modulesMap[name];
    if (!children || children.length === 0) {
      // brak wpisów modułu — nic do sumowania (można logować, ale pomijamy)
      pathVisited[name] = false;
      return;
    }

    for (let ch of children) {
      const childCount = (typeof ch.count === 'number' && !isNaN(ch.count) && ch.count > 0) ? ch.count : 1;
      collectElementTotals(ch.text, multiplier * childCount, modulesMap, zestawyMap, totals, pathVisited);
    }

    // odblokuj moduł dla innych ścieżek
    pathVisited[name] = false;
    return;
  }

  // to element końcowy — sumujemy
  const add = Number(multiplier) || 1;
  totals[name] = (totals[name] || 0) + add;
}

function checkDriveLinkStatus(link) {
  if (!link || link === '(brak linku)') return { ok: false, code: 0, status: 'Brak linku' };

  const fileIdMatch = link.match(/[-\w]{25,}/);
  if (!fileIdMatch) return { ok: false, code: 0, status: 'Nieprawidłowy format linku' };

  const fileId = fileIdMatch[0];
  try {
    const file = DriveApp.getFileById(fileId);
    if (!file) return { ok: false, code: 404, status: 'Nie znaleziono pliku' };
    // Jeśli udało się pobrać nazwę, to plik istnieje
    const name = file.getName();
    return { ok: true, code: 200, status: 'OK', name };
  } catch (e) {
    // Tu złapiemy przypadki 403, 404 itd.
    if (String(e).includes('File not found')) return { ok: false, code: 404, status: 'Nie znaleziono pliku' };
    if (String(e).includes('User does not have permission')) return { ok: false, code: 403, status: 'Brak dostępu' };
    return { ok: false, code: 500, status: 'Błąd: ' + e.message };
  }
}

/**
 * Masowe sprawdzanie wszystkich elementów w arkuszach
 * Działa jak Porównaj linki (SyncLinks), ale dla wszystkich elementów
 * Pomija 1. wiersz (nagłówki)
 */
function massCheckAndFixLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC', 'Elementy CNC'];
  const ui = SpreadsheetApp.getUi();

  // 1️⃣ Zbierz wszystkie wystąpienia elementów i ich linki
  const elementLinks = {}; 
  // Struktura: { "H-M1354-01": [ { sheet, row, link }, ... ] }

  sheetsToCheck.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const richTexts = sheet.getDataRange().getRichTextValues();

    for (let i = 1; i < data.length; i++) {
      const name = data[i][1];
      if (!name) continue;
      const linkObj = richTexts[i][1];
      const link = linkObj && linkObj.getLinkUrl ? linkObj.getLinkUrl() : null;

      if (!elementLinks[name]) elementLinks[name] = [];
      elementLinks[name].push({ sheet: sheetName, row: i + 1, link });
    }
  });

  const inconsistencies = []; // różne linki dla tego samego elementu
  const missingLinks = [];    // elementy bez linku, ale można je uzupełnić

  // 2️⃣ Sprawdź dla każdego elementu, co się dzieje
  for (const [name, entries] of Object.entries(elementLinks)) {
    const uniqueLinks = [...new Set(entries.map(e => e.link).filter(l => !!l))];

    if (uniqueLinks.length === 0) {
      // brak linku nigdzie — do raportu
      missingLinks.push({ name, entries });
    } else if (uniqueLinks.length === 1) {
      // jeden link — można uzupełnić brakujące komórki
      const validLink = uniqueLinks[0];
      entries.forEach(e => {
        if (!e.link) {
          const sheet = ss.getSheetByName(e.sheet);
          const cell = sheet.getRange(e.row, 2);
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(name)
            .setLinkUrl(validLink)
            .build();
          cell.setRichTextValue(richText);
        }
      });
    } else {
      // więcej niż jeden unikalny link → konflikt
      inconsistencies.push({ name, uniqueLinks, entries });
    }
  }

   // 3️⃣ Jeśli są konflikty — zapytaj użytkownika o każdy z nich
  if (inconsistencies.length > 0) {
    for (const conflict of inconsistencies) {
      const { name, uniqueLinks, entries } = conflict;
      let msg = `Element "${name}" ma różne linki:\n\n`;
      uniqueLinks.forEach((l, i) => {
        msg += `${i + 1}. ${l}\n`;
      });
      msg += `\nWpisz numer (1-${uniqueLinks.length}), który link ma być używany we wszystkich wystąpieniach.`;

      const button = ui.prompt(msg, ui.ButtonSet.OK_CANCEL);
      if (button.getSelectedButton() === ui.Button.OK) {
        const response = button.getResponseText().trim();
        const index = parseInt(response, 10);

        if (!isNaN(index) && index >= 1 && index <= uniqueLinks.length) {
          const chosen = uniqueLinks[index - 1]; // wybierz link wg numeru
          entries.forEach(e => {
            const sheet = ss.getSheetByName(e.sheet);
            const cell = sheet.getRange(e.row, 2);
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

  // 4️⃣ Raport o brakach
  if (missingLinks.length > 0) {
    let msg = "⚠️ Elementy, które nigdzie nie mają przypisanego linku:\n\n";
    missingLinks.forEach(e => {
      msg += `• ${e.name} (wystąpień: ${e.entries.length})\n`;
    });
    ui.alert(msg);
  } else if (inconsistencies.length === 0) {
    ui.alert("✅ Wszystkie elementy mają spójne linki i zostały uzupełnione, jeśli brakowało.");
  }
}



/**
 * Pomocnicza wersja promptAndCompareLinks() działająca dla jednego elementu
 * Zwraca: 'fixed' | 'skipped' | 'ok'
 */
function compareAndFixElementLinks(elementName, ss, ui) {
  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC', 'Elementy CNC'];
  const foundLinks = [];

  function checkDriveLinkStatus(link) {
    if (!link || link === '(brak linku)') return { ok: false, status: 'Brak linku' };
    const match = link.match(/[-\w]{25,}/);
    if (!match) return { ok: false, status: 'Nieprawidłowy format linku' };
    const fileId = match[0];
    try {
      DriveApp.getFileById(fileId).getName();
      return { ok: true, status: 'OK' };
    } catch (e) {
      const msg = String(e);
      if (msg.includes('File not found')) return { ok: false, status: 'Nie znaleziono pliku' };
      if (msg.includes('User does not have permission')) return { ok: false, status: 'Brak dostępu' };
      return { ok: false, status: 'Błąd: ' + e.message };
    }
  }

  // 🔍 Zbierz linki z arkuszy
  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const range = sheet.getDataRange();
    const values = range.getValues();
    const rich = range.getRichTextValues();

    for (let r = 1; r < values.length; r++) { // pomijamy nagłówek
      const name = String(values[r][1]).trim();
      if (name === elementName) {
        let link = rich[r][1]?.getLinkUrl() || '(brak linku)';
        const status = checkDriveLinkStatus(link);
        foundLinks.push({ sheet: sheetName, row: r + 1, link, ...status });
      }
    }
  }

  if (foundLinks.length === 0) return 'skipped';

  const invalid = foundLinks.filter(f => !f.ok);
  const uniqueLinks = [...new Set(foundLinks.map(f => f.link))];

  // ❌ Błędne linki
  if (invalid.length > 0) {
    const msg = invalid.map(f => `• ${f.sheet}!B${f.row} → ${f.status}`).join('\n');
    const resp = ui.prompt(`Błędne linki dla "${elementName}"`, `${msg}\n\nPodaj nowy link:`, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return 'skipped';
    const newLink = resp.getResponseText().trim();
    if (!newLink) return 'skipped';
    updateLinks(ss, foundLinks, elementName, newLink, ui);
    return 'fixed';
  }

  // ⚠️ Różne linki
  if (uniqueLinks.length > 1) {
    let msg = `Znaleziono ${uniqueLinks.length} różne linki dla "${elementName}":\n\n`;
    uniqueLinks.forEach((l, i) => (msg += `${i + 1}. ${l}\n`));
    msg += `\nWpisz numer linku, który chcesz zachować (lub wklej nowy):`;
    const resp = ui.prompt('Różne linki', msg, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return 'skipped';
    const inp = resp.getResponseText().trim();
    const idx = /^\d+$/.test(inp) ? parseInt(inp, 10) - 1 : null;
    const selectedLink = idx != null ? uniqueLinks[idx] : inp;
    if (selectedLink) {
      updateLinks(ss, foundLinks, elementName, selectedLink, ui);
      return 'fixed';
    }
    return 'skipped';
  }

  return 'ok';
}


// --- Obsługa pojedynczego elementu ---
function handleSingleElement(el, ui, ss) {
  if (el.invalidLinks.length > 0) {
    const msg = el.invalidLinks.map(f => `• ${f.sheet}!B${f.row} → ${f.status}`).join('\n');
    const response = ui.prompt(
      'Znaleziono błędne linki',
      `Dla elementu "${el.name}" wykryto błędne linki:\n\n${msg}\n\nPodaj nowy, poprawny link:`,
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() === ui.Button.OK) {
      const newLink = response.getResponseText().trim();
      if (newLink) updateLinks(ss, el.linksArr, el.name, newLink, ui);
    }
  } else if (el.uniqueLinks.length > 1) {
    let msg = `Znaleziono ${el.uniqueLinks.length} różne linki dla elementu "${el.name}":\n\n`;
    el.uniqueLinks.forEach((l, i) => {
      msg += `${i + 1}. ${l}\n`;
    });
    msg += `\nWpisz numer linku, który chcesz zachować (lub wklej nowy link):`;
    const response = ui.prompt('Różne linki wykryte', msg, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      const userInput = response.getResponseText().trim();
      let selectedLink = null;
      if (/^\d+$/.test(userInput)) {
        const idx = parseInt(userInput, 10) - 1;
        selectedLink = el.uniqueLinks[idx];
      } else {
        selectedLink = userInput;
      }
      if (selectedLink) updateLinks(ss, el.linksArr, el.name, selectedLink, ui);
    }
  }
}

// --- Pomocnicza funkcja do aktualizacji linków ---
function updateLinks(ss, foundLinks, elementName, newLink, ui) {
  let updatedCount = 0;

  for (const f of foundLinks) {
    const sheet = ss.getSheetByName(f.sheet);
    if (!sheet) continue;
    const cell = sheet.getRange(f.row, 2);
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


/**
 * Próbuje odgadnąć rozszerzenie pliku:
 *  - najpierw z końcówki URL
 *  - potem z content-type
 * Zwraca np. ".dxf" lub "dxf" albo null
 */
function guessExtension(url, blob) {
  try {
    // 1) z URL
    const m = url.match(/(\.[a-z0-9]{1,6})(?:[\?#]|$)/i);
    if (m && m[1]) {
      return m[1];
    }
    // 2) z blob contentType
    const ct = blob.getContentType();
    if (ct) {
      if (ct.indexOf('dxf') !== -1) return '.dxf';
      // inne mapowania możliwe:
      if (ct.indexOf('octet-stream') !== -1) return '';
      // spróbuj wyciąć typ/subtype
      const parts = ct.split('/');
      if (parts.length > 1) {
        const subtype = parts[1].split('+')[0];
        return '.' + subtype;
      }
    }
  } catch (e) {
    // ignore
  }
  return '';
}
