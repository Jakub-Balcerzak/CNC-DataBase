/** Downloading-related functions (DXF download, recursion, colors) */

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

  // Zbierz listę unikalnych elementów (rozwiń moduły rekurencyjnie)
  const elements = [];
  const seen = new Set();
  const collect = (list) => {
    for (const e of list) {
      if (!e || !e.text) continue;
      const name = String(e.text).trim();
      if (!name) continue;

      if (isModuleName(name)) {
        const children = modulesMap[name] || [];
        collect(children);
      } else {
        if (!seen.has(name)) {
          seen.add(name);
          const data = findElementData(name, modulesMap, zestawyMap) || {};
          const rich = e.richLink || findLinkForElement(name, modulesMap, zestawyMap);
          elements.push({ text: name, name: data.name || '', richLink: rich });
        }
      }
    }
  };

  collect(startElements);

  if (elements.length === 0) {
    ui.alert('Brak elementów do pobrania.');
    return;
  }

  const htmlTemplate = HtmlService.createTemplateFromFile('colorSelector');
  htmlTemplate.data = elements;
  htmlTemplate.setId = setId;

  const htmlOutput = htmlTemplate.evaluate()
    .setTitle(`Kolory dla ${setId}`)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Kolory dla ${setId}`);
}

/**
 * Tworzy plik TXT z listą elementów (ilości) dla podanego zestawu i zapisuje go w nowym folderze na Dysku.
 */
function createElementListForSet(setId) {
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

  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

  const startElements = zestawyMap[setId];
  if (!startElements || startElements.length === 0) {
    ui.alert('Nie znaleziono zestawu', `Nie znaleziono wierszy o Nr zestawu = "${setId}" w arkuszu "${SHEET_ZESTAWY}".`, ui.ButtonSet.OK);
    return;
  }

  const totals = {};
  const pathVisited = {};
  for (const e of startElements) {
    const count = (e && typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    collectElementTotals(e.text, count, modulesMap, zestawyMap, totals, pathVisited);
  }

  const downloaded = [];
  for (const name of Object.keys(totals).sort()) {
    const data = findElementData(name, modulesMap, zestawyMap) || {};
    downloaded.push({ name: name, count: totals[name], prettyName: data.name || '', color: 'Bez koloru' });
  }

  if (downloaded.length === 0) {
    ui.alert('Brak elementów', `Zestaw ${setId} nie zawiera elementów do listy.`);
    return;
  }

  const folderName = `Rejestr Plików CNC - Lista ${setId} - ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);

  // createSummaryTxtFile (in helpers.js) expects (downloaded, folder, filename?)
  const filename = `Lista_${setId}_${timestampForName()}.txt`;
  const fileUrl = createSummaryTxtFile(downloaded, folder, filename);
  if (fileUrl) {
    ui.alert('Gotowe', `Utworzono listę elementów dla zestawu ${setId}.
Plik: ${fileUrl}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Błąd', 'Nie udało się utworzyć pliku z listą elementów.', ui.ButtonSet.OK);
  }
}

/**
 * Tworzy plik TXT z listą elementów dla modułu (używa arkusza 'Moduły CNC').
 */
function createElementListForModule(modId) {
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

  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

  const startElements = modulesMap[modId];
  if (!startElements || startElements.length === 0) {
    ui.alert('Nie znaleziono modułu', `Nie znaleziono wierszy o Nr modułu = "${modId}" w arkuszu "${SHEET_MODULE}".`, ui.ButtonSet.OK);
    return;
  }

  const totals = {};
  const pathVisited = {};
  // startElements are children of the module; use each child's count
  for (const e of startElements) {
    const count = (e && typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    collectElementTotals(e.text, count, modulesMap, zestawyMap, totals, pathVisited);
  }

  const downloaded = [];
  for (const name of Object.keys(totals).sort()) {
    const data = findElementData(name, modulesMap, zestawyMap) || {};
    downloaded.push({ name: name, count: totals[name], prettyName: data.name || '', color: 'Bez koloru' });
  }

  if (downloaded.length === 0) {
    ui.alert('Brak elementów', `Moduł ${modId} nie zawiera elementów do listy.`);
    return;
  }

  const folderName = `Rejestr Plików CNC - Lista moduł ${modId} - ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const filename = `Lista_modul_${modId}_${timestampForName()}.txt`;
  const fileUrl = createSummaryTxtFile(downloaded, folder, filename);
  if (fileUrl) {
    ui.alert('Gotowe', `Utworzono listę elementów dla modułu ${modId}.
Plik: ${fileUrl}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Błąd', 'Nie udało się utworzyć pliku z listą elementów.', ui.ButtonSet.OK);
  }
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

  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

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

  const startElements = modulesMap[modId];
  if (!startElements || startElements.length === 0) {
    ui.alert('Nie znaleziono modułu', `Nie znaleziono wierszy o Nr modułu = "${modId}" w arkuszu "${SHEET_MODULE}".`, ui.ButtonSet.OK);
    return;
  }

  const folderName = `Rejestr Plików CNC - Moduł ${modId} - ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl();

  const visited = {};
  const missingLinks = [];
  const downloaded = [];
  const errors = [];

  for (let e of startElements) {
    processElementRecursive(e.text, e.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors);
  }

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

  let link = providedRichLink || findLinkForElement(name, modulesMap, zestawyMap);
  if (!link) {
    missingLinks.push(name);
    return;
  }

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

function processElementRecursiveWithColor(name, providedRichLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap, multiplier = 1) {
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
      processElementRecursiveWithColor(ch.text, ch.richLink, modulesMap, zestawyMap, visited, folder, downloaded, missingLinks, errors, colorMap, childMultiplier);
    }
    return;
  }

  const link = providedRichLink || findLinkForElement(name, modulesMap, zestawyMap);
  if (!link) {
    missingLinks.push(name);
    return;
  }

  try {
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

    const color = colorMap[name] || 'Bez koloru';
    const colorFolder = getOrCreateSubfolder(folder, color);

    const existing = downloaded.find(d => d.name === name && d.color === color);
    if (!existing) {
      colorFolder.createFile(blob);
      downloaded.push({ name: name, color: color, prettyName: elementData?.name || '', count: effectiveCount });
    } else {
      existing.count = (existing.count || 0) + effectiveCount;
    }

  } catch (e) {
    errors.push(`${name}: ${e.message}`);
  }
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

  const totals = {};
  const pathVisited = {};

  for (let e of startElements) {
    const topCount = (typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    collectElementTotals(e.text, topCount, modulesMap, zestawyMap, totals, pathVisited);
  }

  const elementNames = Object.keys(totals);
  if (elementNames.length === 0) {
    ui.alert('Brak elementów do pobrania.');
    return;
  }

  const folderName = `Rejestr Plików CNC - Pobrania ${setId} - ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl();

  const missingLinks = [];
  const downloaded = [];
  const errors = [];

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

      const color = colorMap && colorMap[name] ? colorMap[name] : 'Bez koloru';
      const colorFolder = getOrCreateSubfolder(folder, color);
      colorFolder.createFile(blob);

      const elementData = findElementData(name, modulesMap, zestawyMap);

      downloaded.push({ name: name, color: color, prettyName: elementData?.name || '', count: count });
    } catch (e) {
      errors.push(`${name}: ${e.message}`);
    }
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
  // Note: Summary TXT creation is handled by a separate action/menu.
  ui.alert('Zakończono', summary.join('\n'), ui.ButtonSet.OK);
}
