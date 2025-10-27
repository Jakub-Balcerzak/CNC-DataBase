/**
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
    .addItem('Pobierz pliki dla zestawu...', 'promptAndDownload')
    .addItem('Pobierz pliki dla zestawu (z kolorami)...', 'promptAndDownloadWithColors') // 🆕 nowa opcja
    .addItem('Pobierz pliki dla modułu...', 'promptAndDownloadModule')
    .addToUi();

  ui.createMenu('Sync')
    .addItem('Ustaw / edytuj link dla elementu...', 'promptAndSyncLink')
    .addItem('Porównaj linki (SyncLinks)', 'promptAndCompareLinks')
    .addToUi();
}

function promptAndCompareLinks() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = ['Zestawy CNC', 'Moduły CNC'];

  // 1️⃣ Pytanie o numer katalogowy elementu
  const resp = ui.prompt('Porównaj linki', 'Podaj numer katalogowy elementu (np. H_P3300_14):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const elementName = resp.getResponseText().trim();
  if (!elementName) {
    ui.alert('Nie podano numeru katalogowego elementu.');
    return;
  }

  // 2️⃣ Zbierz wszystkie linki dla tego elementu ze wszystkich arkuszy
  const foundLinks = []; // [{sheet, row, link}]
  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const range = sheet.getDataRange();
    const values = range.getValues();
    const richValues = range.getRichTextValues();

    for (let r = 0; r < values.length; r++) {
      const cellVal = String(values[r][1]).trim();
      if (cellVal === elementName) {
        const rich = richValues[r][1];
        let link = null;
        try {
          link = rich.getLinkUrl();
        } catch (e) {
          link = null;
        }
        foundLinks.push({
          sheet: sheetName,
          row: r + 1,
          link: link || '(brak linku)'
        });
      }
    }
  }

  // 3️⃣ Walidacja — brak powtórzeń
  if (foundLinks.length === 0) {
    ui.alert('Nie znaleziono', `Nie znaleziono elementu "${elementName}" w arkuszach.`, ui.ButtonSet.OK);
    return;
  }

  // 4️⃣ Sprawdzenie czy wszystkie linki są takie same
  const uniqueLinks = [...new Set(foundLinks.map(f => f.link))];

  if (uniqueLinks.length === 1) {
    ui.alert('Synchronizacja OK ✅', `Wszystkie wystąpienia elementu "${elementName}" mają ten sam link:\n\n${uniqueLinks[0]}`, ui.ButtonSet.OK);
    return;
  }

  // 5️⃣ Występują różne linki → pokaż listę i zapytaj, który ma być prawidłowy
  let msg = `Znaleziono różne linki dla elementu "${elementName}":\n\n`;
  foundLinks.forEach(f => {
    msg += `📄 ${f.sheet}!B${f.row}\n→ ${f.link}\n\n`;
  });
  msg += `Wpisz dokładnie numer opcji (1–${uniqueLinks.length}) z poniższej listy, który ma być ustawiony jako prawidłowy:\n\n`;
  uniqueLinks.forEach((l, i) => {
    msg += `${i + 1}. ${l}\n`;
  });

  const resp2 = ui.prompt('Wybierz link do synchronizacji', msg, ui.ButtonSet.OK_CANCEL);
  if (resp2.getSelectedButton() !== ui.Button.OK) return;
  const chosenIdx = parseInt(resp2.getResponseText().trim());
  if (isNaN(chosenIdx) || chosenIdx < 1 || chosenIdx > uniqueLinks.length) {
    ui.alert('Nieprawidłowy wybór.', ui.ButtonSet.OK);
    return;
  }

  const correctLink = uniqueLinks[chosenIdx - 1];

  // 6️⃣ Podmień wszystkie linki na wybrany
  let updated = 0;
  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const range = sheet.getDataRange();
    const values = range.getValues();

    for (let r = 0; r < values.length; r++) {
      const cellVal = String(values[r][1]).trim();
      if (cellVal === elementName) {
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(cellVal)
          .setLinkUrl(correctLink)
          .build();
        sheet.getRange(r + 1, 2).setRichTextValue(richText);
        updated++;
      }
    }
  }

  ui.alert('Synchronizacja zakończona 🔁', `Ujednolicono ${updated} komórek dla elementu "${elementName}".\nUstawiony link:\n${correctLink}`, ui.ButtonSet.OK);
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

    if (!map[key]) map[key] = [];
    map[key].push({ text: dataText, richLink: richLink, surface: surface, name: name });
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
function processElementRecursiveWithColor(name, providedRichLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap) {
  name = String(name).trim();
  if (!name) return;
  if (visited[name]) return;
  visited[name] = true;

  if (isModuleName(name)) {
    const children = modulesMap[name];
    if (!children || children.length === 0) {
      missingLinks.push(`Moduł ${name} - brak wpisów w "${SHEET_MODULE}"`);
      return;
    }
    for (let ch of children) {
      processElementRecursiveWithColor(ch.text, ch.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap);
    }
  } else {
    const link = providedRichLink || findLinkForElement(name, modulesMap, zestawyMap);
    if (!link) {
      missingLinks.push(name);
      return;
    }

    try {
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

      // 🟢 Kolor (folder docelowy)
      const color = colorMap[name] || 'Bez koloru';
      const colorFolder = getOrCreateSubfolder(folder, color);
      const file = colorFolder.createFile(blob);

      downloaded.push({ name, color });
    } catch (e) {
      errors.push(`${name}: ${e.message}`);
    }
  }
}

/**
 * Rekurencyjnie przetwarza element/moduł.
 * - jeśli element (nie moduł): szuka hiperłącza (jeśli przekazano richLink - używa) i pobiera plik
 * - jeśli moduł: rozkłada przez modulesMap (jeśli brak -> traktuj jako brak definicji)
 */
function processElementRecursive(name, providedRichLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors) {
  name = String(name).trim();
  if (!name) return;
  if (visited[name]) return;
  visited[name] = true;

  if (isModuleName(name)) {
    const children = modulesMap[name];
    if (!children || children.length === 0) {
      missingLinks.push(`Moduł ${name} - brak wpisów w "${SHEET_MODULE}"`);
      return;
    }
    for (let ch of children) {
      processElementRecursive(ch.text, ch.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors);
    }
  } else {
    let link = providedRichLink || findLinkForElement(name, modulesMap, zestawyMap);
    if (!link) {
      missingLinks.push(name);
      return;
    }

    // znajdź dane o elemencie
    const elementData = findElementData(name, modulesMap, zestawyMap);

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
      const file = folder.createFile(blob);

      downloaded.push({
        name: name,
        url: directLink,
        fileId: fileIdMatch ? fileIdMatch[0] : 'unknown',
        prettyName: elementData?.name || '',
        surface: elementData?.surface || null
      });
    } catch (e) {
      errors.push(`${name}: ${e.message}`);
    }
  }
}

/**
 * Próbuje znaleźć link dla elementu:
 * - przeszukuje moduły: wiersze w których kolumna B == name (i wykorzysta link z tej komórki)
 * - przeszukuje zestawy: wiersze w których kolumna B == name (i wykorzysta link z tej komórki)
 * Zwraca pierwszy znaleziony link albo null.
 */
function findLinkForElement(name, modulesMap, zestawyMap) {
  // przeszukujemy modulesMap - to tablica obiektów {text, richLink} dla każdego modułu klucza
  // modulesMap jest mapą moduł -> array; ale chcemy przeszukać wszystkie wartości arrays
  for (let key in modulesMap) {
    const arr = modulesMap[key];
    for (let entry of arr) {
      if (entry.text === name && entry.richLink) return entry.richLink;
    }
  }
  // przeszukaj zestawy (mogą zawierać elementy z linkiem)
  for (let key in zestawyMap) {
    const arr = zestawyMap[key];
    for (let entry of arr) {
      if (entry.text === name && entry.richLink) return entry.richLink;
    }
  }
  return null;
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

  const folderName = `Rejestr Plików CNC - Pobrania ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl();

  const visited = {};
  const missingLinks = [];
  const downloaded = [];
  const errors = [];

  for (let e of startElements) {
    processElementRecursiveWithColor(e.text, e.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap);
  }

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

  ui.alert('Zakończono', summary.join('\n'), ui.ButtonSet.OK);
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
