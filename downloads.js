/** Downloading-related functions (DXF download, recursion, colors) */

function downloadSetFilesWithColors(setId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetZest = ss.getSheetByName(SHEET_ZESTAWY);
  const sheetMod = ss.getSheetByName(SHEET_MODULE);
  const ui = SpreadsheetApp.getUi();

  if (!sheetZest || !sheetMod) {
    ui.alert('Błąd', `Brakuje arkuszy "${SHEET_ZESTAWY}" lub "${SHEET_MODULE}".`, ui.ButtonSet.OK);
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
    ui.alert('Nie znaleziono zestawu', `Brak wierszy o Nr zestawu = "${setId}".`, ui.ButtonSet.OK);
    return;
  }

  // Zbierz listę unikalnych elementów (rozwiń moduły rekurencyjnie)
  const elements = [];
  const seen = new Set();
  const collect = (list, parentModule) => {
    for (const e of list) {
      if (!e || !e.text) continue;
      const name = String(e.text).trim();
      if (!name) continue;

      if (isModuleName(name)) {
        const children = modulesMap[name] || [];
        collect(children, name);
      } else {
        if (!seen.has(name)) {
          seen.add(name);
          const data = findElementData(name, modulesMap, zestawyMap) || {};
          const rich = e.richLink || findLinkForElement(name, modulesMap, zestawyMap);
          elements.push({ text: name, name: data.name || '', richLink: rich, module: parentModule || '' });
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
  const moduleMap = {};
  for (const e of startElements) {
    const count = (e && typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    collectElementTotals(e.text, count, modulesMap, zestawyMap, totals, pathVisited, moduleMap);
  }

  const downloaded = [];
  for (const name of Object.keys(totals).sort()) {
    const data = findElementData(name, modulesMap, zestawyMap) || {};
    downloaded.push({ name: name, count: totals[name], prettyName: data.name || '', color: 'Bez koloru', module: moduleMap[name] || '' });
  }

  if (downloaded.length === 0) {
    ui.alert('Brak elementów', `Zestaw ${setId} nie zawiera elementów do listy.`, ui.ButtonSet.OK);
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
  const moduleMap = {};
  for (const e of startElements) {
    const count = (e && typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    // pass parentModule = modId so collected elements get this module noted
    collectElementTotals(e.text, count, modulesMap, zestawyMap, totals, pathVisited, moduleMap, modId);
  }

  const downloaded = [];
  for (const name of Object.keys(totals).sort()) {
    const data = findElementData(name, modulesMap, zestawyMap) || {};
    downloaded.push({ name: name, count: totals[name], prettyName: data.name || '', color: 'Bez koloru', module: moduleMap[name] || '' });
  }

  if (downloaded.length === 0) {
    ui.alert('Brak elementów', `Moduł ${modId} nie zawiera elementów do listy.`, ui.ButtonSet.OK);
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

  // ═══════════════════════════════════════════════════════════════════════════
  // POBIERANIE PLIKÓW UCANCAM DLA ELEMENTÓW
  // ═══════════════════════════════════════════════════════════════════════════
  
  const COL_ELEMENT = 1; // kolumna B (0-based)
  const COL_UCANCAM = 5; // kolumna F (0-based)
  
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

  const ucancamDownloaded = [];
  const ucancamMissing = [];
  const ucancamErrors = [];
  let ucancamFolder = null;
  const processedElements = new Set(); // unikaj duplikatów

  // Szukaj wszystkich elementów w module i pobierz ich pliki UCANCAM
  for (let i = 1; i < modValues.length; i++) {
    const rowModName = String(modValues[i][0]).trim();
    const elementName = String(modValues[i][COL_ELEMENT]).trim();
    
    // Sprawdź czy to wiersz z naszego modułu i czy jest element (nie moduł)
    if (rowModName === modId && elementName && !isModuleName(elementName)) {
      // Sprawdź czy już przetworzono ten element
      if (processedElements.has(elementName)) continue;
      processedElements.add(elementName);
      
      const ucancamLink = getRichLink(modRich[i][COL_UCANCAM]);
      
      if (ucancamLink) {
        try {
          const fileIdMatch = ucancamLink.match(/[-\w]{25,}/);
          if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const sourceFile = DriveApp.getFileById(fileId);
            const originalFileName = sourceFile.getName();
            
            // Utwórz podfolder UCANCAM jeśli nie istnieje
            if (!ucancamFolder) {
              ucancamFolder = folder.createFolder('UCANCAM');
            }
            
            sourceFile.makeCopy(originalFileName, ucancamFolder);
            ucancamDownloaded.push({ name: elementName, fileName: originalFileName });
          } else {
            ucancamErrors.push(`${elementName}: nieprawidłowy format linku`);
          }
        } catch (e) {
          ucancamErrors.push(`${elementName}: ${e.message}`);
        }
      } else {
        ucancamMissing.push(elementName);
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // GENEROWANIE PLIKU RAPORTU TXT
  // ═══════════════════════════════════════════════════════════════════════════
  
  const reportLines = [];
  reportLines.push(`RAPORT POBIERANIA - Moduł ${modId}`);
  reportLines.push(`Data: ${new Date().toLocaleString('pl-PL')}`);
  reportLines.push('═'.repeat(60));
  reportLines.push('');
  reportLines.push(`📁 Folder: ${folderUrl}`);
  reportLines.push('');
  
  // Pobrane pliki DXF
  reportLines.push(`📄 POBRANE PLIKI DXF: ${downloaded.length}`);
  reportLines.push('-'.repeat(40));
  if (downloaded.length > 0) {
    downloaded.forEach(d => {
      const countStr = d.count > 1 ? ` x${d.count}` : '';
      const prettyStr = d.prettyName ? ` - ${d.prettyName}` : '';
      const surfaceStr = d.surface ? ` (${d.surface.toFixed(3)} m²)` : '';
      reportLines.push(`  • ${d.name}${prettyStr}${countStr}${surfaceStr}`);
    });
  }
  reportLines.push('');
  
  // Pobrane pliki UCANCAM
  reportLines.push(`📦 POBRANE PLIKI UCANCAM: ${ucancamDownloaded.length}`);
  reportLines.push('-'.repeat(40));
  if (ucancamDownloaded.length > 0) {
    ucancamDownloaded.forEach(u => {
      reportLines.push(`  ✅ ${u.name} → ${u.fileName}`);
    });
  }
  reportLines.push('');
  
  // Brakujące linki DXF
  if (missingLinks.length > 0) {
    reportLines.push(`❌ BRAK HIPERŁĄCZY DXF: ${missingLinks.length}`);
    reportLines.push('-'.repeat(40));
    missingLinks.forEach(m => {
      reportLines.push(`  • ${m}`);
    });
    reportLines.push('');
  }
  
  // Brakujące linki UCANCAM
  if (ucancamMissing.length > 0) {
    reportLines.push(`⚠️ BRAK LINKÓW UCANCAM: ${ucancamMissing.length}`);
    reportLines.push('-'.repeat(40));
    ucancamMissing.forEach(m => {
      reportLines.push(`  • ${m}`);
    });
    reportLines.push('');
  }
  
  // Błędy
  if (errors.length > 0 || ucancamErrors.length > 0) {
    reportLines.push(`⚠️ BŁĘDY: ${errors.length + ucancamErrors.length}`);
    reportLines.push('-'.repeat(40));
    errors.forEach(e => {
      reportLines.push(`  • ${e}`);
    });
    ucancamErrors.forEach(e => {
      reportLines.push(`  • UCANCAM: ${e}`);
    });
    reportLines.push('');
  }
  
  // Ostrzeżenia dotyczące danych
  if (dataWarnings.length > 0) {
    reportLines.push(`⚠️ OSTRZEŻENIA DOTYCZĄCE DANYCH: ${dataWarnings.length}`);
    reportLines.push('-'.repeat(40));
    dataWarnings.forEach(w => {
      reportLines.push(`  • ${w}`);
    });
    reportLines.push('');
  }
  
  reportLines.push('═'.repeat(60));
  reportLines.push(`Wygenerowano automatycznie przez Rejestr Plików CNC`);
  
  // Zapisz plik raportu
  const reportFilename = `Raport_${modId}_${timestampForName()}.txt`;
  const reportContent = reportLines.join('\n');
  folder.createFile(reportFilename, reportContent, 'text/plain');

  // ═══════════════════════════════════════════════════════════════════════════
  // PODSUMOWANIE (dialog)
  // ═══════════════════════════════════════════════════════════════════════════

  const summaryLines = [];
  summaryLines.push(`📁 Utworzony folder:`);
  summaryLines.push(folderUrl);
  summaryLines.push('');
  summaryLines.push(`📄 Pobrano plików DXF: ${downloaded.length}`);
  if (downloaded.length) {
    downloaded.slice(0, 15).forEach(d => {
      const surfaceStr = d.surface ? ` (${d.surface.toFixed(3)} m²)` : '';
      const pretty = d.prettyName ? ` – ${d.prettyName}` : '';
      summaryLines.push(`   • ${d.name}${pretty}${surfaceStr}`);
    });
    if (downloaded.length > 15) summaryLines.push(`   ... + ${downloaded.length - 15} innych`);
  }
  
  // Info o UCANCAM
  summaryLines.push('');
  if (ucancamDownloaded.length > 0) {
    summaryLines.push(`📦 Pobrano plików UCANCAM: ${ucancamDownloaded.length}`);
    ucancamDownloaded.slice(0, 10).forEach(u => {
      summaryLines.push(`   ✅ ${u.name} → ${u.fileName}`);
    });
    if (ucancamDownloaded.length > 10) {
      summaryLines.push(`   ... + ${ucancamDownloaded.length - 10} innych`);
    }
  }
  
  if (ucancamMissing.length > 0) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Elementy bez linku UCANCAM (${ucancamMissing.length}):`);
    ucancamMissing.slice(0, 10).forEach(m => summaryLines.push(`   • ${m}`));
    if (ucancamMissing.length > 10) {
      summaryLines.push(`   ... + ${ucancamMissing.length - 10} innych`);
    }
  }
  
  if (missingLinks.length) {
    summaryLines.push('');
    summaryLines.push(`❌ Elementy bez hiperłącza DXF (${missingLinks.length}):`);
    missingLinks.slice(0, 10).forEach(m => summaryLines.push(`   • ${m}`));
    if (missingLinks.length > 10) summaryLines.push(`   ... + ${missingLinks.length - 10} innych`);
  }
  if (errors.length || ucancamErrors.length) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Błędy (${errors.length + ucancamErrors.length}):`);
    errors.slice(0, 5).forEach(err => summaryLines.push(`   • ${err}`));
    ucancamErrors.slice(0, 5).forEach(err => summaryLines.push(`   • UCANCAM: ${err}`));
    const totalErrors = errors.length + ucancamErrors.length;
    if (totalErrors > 10) summaryLines.push(`   ... + ${totalErrors - 10} innych`);
  }
  if (dataWarnings.length) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Ostrzeżenia dotyczące danych (${dataWarnings.length}):`);
    dataWarnings.slice(0, 5).forEach(w => summaryLines.push(`   • ${w}`));
    if (dataWarnings.length > 5) summaryLines.push(`   ... + ${dataWarnings.length - 5} innych`);
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

  const COL_UCANCAM = 5; // kolumna F (0-based index)
  const COL_ELEMENT = 1; // kolumna B (0-based index)

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

  // Optimized download:
  // - precompute maps
  // - use DriveApp.copy (makeCopy) for Drive fileIds when possible
  // - use UrlFetchApp.fetchAll in batches for the rest
  const zestValues = sheetZest.getDataRange().getValues();
  const zestRich = sheetZest.getDataRange().getRichTextValues();
  const modValues = sheetMod.getDataRange().getValues();
  const modRich = sheetMod.getDataRange().getRichTextValues();

  const zestawyMap = buildMapForSheet(zestValues, zestRich, 0, 1, SHEET_ZESTAWY).map;
  const modulesMap = buildMapForSheet(modValues, modRich, 0, 1, SHEET_MODULE).map;

  const startElements = zestawyMap[setId];
  if (!startElements || startElements.length === 0) {
    if (ui) ui.alert('Nie znaleziono zestawu', `Brak wierszy o Nr zestawu = "${setId}".`);
    return;
  }

  // 1) totals + zbieranie elementów i modułów użytych w zestawie
  const totals = {};
  const pathVisited = {};
  const moduleMap = {};
  const usedModules = new Set(); // moduły użyte w zestawie
  const directElements = new Set(); // elementy bezpośrednio w zestawie (nie w modułach)

  // Rozszerzona wersja collectElementTotals która zbiera moduły i śledzi pochodzenie elementów
  function collectWithModules(name, multiplier, parentModule, isDirectInSet) {
    name = String(name).trim();
    if (!name) return;

    if (isModuleName(name)) {
      usedModules.add(name); // dodaj moduł do listy
      
      if (pathVisited[name]) return;
      pathVisited[name] = true;

      const children = modulesMap[name];
      if (children && children.length > 0) {
        for (let ch of children) {
          const childCount = (typeof ch.count === 'number' && !isNaN(ch.count) && ch.count > 0) ? ch.count : 1;
          collectWithModules(ch.text, multiplier * childCount, name, false);
        }
      }
      pathVisited[name] = false;
      return;
    }

    // Element końcowy
    const add = Number(multiplier) || 1;
    totals[name] = (totals[name] || 0) + add;
    if (parentModule && !moduleMap[name]) {
      moduleMap[name] = parentModule;
    }
    
    // Zapamiętaj elementy bezpośrednie w zestawie
    if (isDirectInSet) {
      directElements.add(name);
    }
  }

  for (let e of startElements) {
    const topCount = (typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
    collectWithModules(e.text, topCount, null, true);
  }

  const elementNames = Object.keys(totals);
  if (elementNames.length === 0) {
    if (ui) ui.alert('Brak elementów do pobrania.');
    return;
  }

  // Precompute link and elementData maps to avoid repeated scanning
  const linkMap = {};
  const dataMap = {};
  for (const name of elementNames) {
    const link = findLinkForElement(name, modulesMap, zestawyMap);
    linkMap[name] = link;
    const d = findElementData(name, modulesMap, zestawyMap) || {};
    dataMap[name] = d;
  }

  // 2) create folder
  const folderName = `Rejestr Plików CNC - Pobrania ${setId} - ${timestampForName()}`;
  const folder = DriveApp.createFolder(folderName);
  const folderUrl = folder.getUrl();

  const missingLinks = [];
  const downloaded = [];
  const errors = [];

  // Partition into Drive-copyable and fetch-needed
  const copyJobs = [];
  const fetchJobs = [];

  for (const name of elementNames) {
    const link = linkMap[name];
    const count = totals[name] || 0;
    const color = (colorMap && colorMap[name]) ? colorMap[name] : 'Bez koloru';
    if (!link) {
      missingLinks.push(name);
      continue;
    }
    const fileIdMatch = link.match(/[-\w]{25,}/);
    if (fileIdMatch) {
      copyJobs.push({ name, fileId: fileIdMatch[0], color, count });
    } else {
      const direct = link;
      fetchJobs.push({ name, url: direct, color, count });
    }
  }

  // 3) Perform Drive copies (fast)
  for (const job of copyJobs) {
    try {
      const src = DriveApp.getFileById(job.fileId);
      const fileName = sanitizeFileName(job.name) + '.dxf';
      const colorFolder = getOrCreateSubfolder(folder, job.color);
      src.makeCopy(fileName, colorFolder);
      downloaded.push({ name: job.name, color: job.color, prettyName: dataMap[job.name]?.name || '', count: job.count, module: moduleMap[job.name] || '' });
    } catch (e) {
      errors.push(`${job.name}: (copy) ${e.message}`);
      try {
        const dl = `https://drive.google.com/uc?export=download&id=${job.fileId}`;
        fetchJobs.push({ name: job.name, url: dl, color: job.color, count: job.count });
      } catch (e2) {
        Logger.log('Fallback enqueue failed for ' + job.name + ': ' + e2.message);
      }
    }
  }

  // 4) Fetch remaining via UrlFetchApp.fetchAll in batches
  const BATCH = 20;
  for (let i = 0; i < fetchJobs.length; i += BATCH) {
    const slice = fetchJobs.slice(i, i + BATCH);
    const requests = slice.map(j => ({ url: j.url, muteHttpExceptions: true, followRedirects: true }));
    let responses = [];
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      Logger.log('fetchAll failed: ' + e.message + ' — falling back to sequential fetch');
      for (let k = 0; k < slice.length; k++) {
        try {
          const r = UrlFetchApp.fetch(slice[k].url, { muteHttpExceptions: true });
          responses.push(r);
        } catch (e2) {
          responses.push(null);
        }
      }
    }

    for (let k = 0; k < slice.length; k++) {
      const job = slice[k];
      const resp = responses[k];
      if (!resp) {
        errors.push(`${job.name}: fetch failed (no response)`);
        continue;
      }
      try {
        const code = resp.getResponseCode();
        if (code < 200 || code >= 300) {
          errors.push(`${job.name}: HTTP ${code} for ${job.url}`);
          continue;
        }
        const blob = resp.getBlob();
        const fileName = sanitizeFileName(job.name) + '.dxf';
        blob.setName(fileName);
        const colorFolder = getOrCreateSubfolder(folder, job.color);
        colorFolder.createFile(blob);
        downloaded.push({ name: job.name, color: job.color, prettyName: dataMap[job.name]?.name || '', count: job.count, module: moduleMap[job.name] || '' });
      } catch (e) {
        errors.push(`${job.name}: ${e.message}`);
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // POBIERANIE PLIKÓW UCANCAM
  // ═══════════════════════════════════════════════════════════════════════════
  
  const ucancamDownloaded = [];
  const ucancamMissing = [];
  const ucancamErrors = [];
  let ucancamFolder = null;
  const processedUcancam = new Set(); // unikaj duplikatów

  // 5a) Pobierz UCANCAM dla ELEMENTÓW bezpośrednio w zestawie (z arkusza Zestawy CNC)
  for (let i = 1; i < zestValues.length; i++) {
    const rowSetId = String(zestValues[i][0]).trim();
    const elementName = String(zestValues[i][COL_ELEMENT]).trim();
    
    // Sprawdź czy to wiersz z naszego zestawu i czy to element (nie moduł)
    if (rowSetId === setId && elementName && !isModuleName(elementName)) {
      // Sprawdź czy już przetworzono ten element
      if (processedUcancam.has(elementName)) continue;
      processedUcancam.add(elementName);
      
      const ucancamLink = getRichLink(zestRich[i][COL_UCANCAM]);
      
      if (ucancamLink) {
        try {
          const fileIdMatch = ucancamLink.match(/[-\w]{25,}/);
          if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const sourceFile = DriveApp.getFileById(fileId);
            const originalFileName = sourceFile.getName();
            
            if (!ucancamFolder) {
              ucancamFolder = folder.createFolder('UCANCAM');
            }
            
            sourceFile.makeCopy(originalFileName, ucancamFolder);
            ucancamDownloaded.push({ name: elementName, fileName: originalFileName, type: 'element', source: 'zestaw' });
          } else {
            ucancamErrors.push(`${elementName}: nieprawidłowy format linku`);
          }
        } catch (e) {
          ucancamErrors.push(`${elementName}: ${e.message}`);
        }
      } else {
        ucancamMissing.push({ name: elementName, type: 'element', source: 'zestaw' });
      }
    }
  }

  // 5b) Pobierz UCANCAM dla wszystkich ELEMENTÓW z MODUŁÓW użytych w zestawie
  for (const modName of usedModules) {
    // Szukaj wszystkich elementów w module
    for (let i = 1; i < modValues.length; i++) {
      const rowModName = String(modValues[i][0]).trim();
      const elementName = String(modValues[i][COL_ELEMENT]).trim();
      
      // Sprawdź czy to wiersz z tego modułu i czy jest element (nie moduł)
      if (rowModName === modName && elementName && !isModuleName(elementName)) {
        // Sprawdź czy już przetworzono ten element
        if (processedUcancam.has(elementName)) continue;
        processedUcancam.add(elementName);
        
        const ucancamLink = getRichLink(modRich[i][COL_UCANCAM]);
        
        if (ucancamLink) {
          try {
            const fileIdMatch = ucancamLink.match(/[-\w]{25,}/);
            if (fileIdMatch) {
              const fileId = fileIdMatch[0];
              const sourceFile = DriveApp.getFileById(fileId);
              const originalFileName = sourceFile.getName();
              
              if (!ucancamFolder) {
                ucancamFolder = folder.createFolder('UCANCAM');
              }
              
              sourceFile.makeCopy(originalFileName, ucancamFolder);
              ucancamDownloaded.push({ name: elementName, fileName: originalFileName, type: 'element', source: modName });
            } else {
              ucancamErrors.push(`${elementName} (z ${modName}): nieprawidłowy format linku`);
            }
          } catch (e) {
            ucancamErrors.push(`${elementName} (z ${modName}): ${e.message}`);
          }
        } else {
          ucancamMissing.push({ name: elementName, type: 'element', source: modName });
        }
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // GENEROWANIE PLIKU RAPORTU TXT
  // ═══════════════════════════════════════════════════════════════════════════
  
  const reportLines = [];
  reportLines.push(`RAPORT POBIERANIA - ${setId}`);
  reportLines.push(`Data: ${new Date().toLocaleString('pl-PL')}`);
  reportLines.push('═'.repeat(60));
  reportLines.push('');
  reportLines.push(`📁 Folder: ${folderUrl}`);
  reportLines.push('');
  
  // Pobrane pliki DXF
  reportLines.push(`📄 POBRANE PLIKI DXF: ${downloaded.length}`);
  reportLines.push('-'.repeat(40));
  if (downloaded.length > 0) {
    // Grupuj wg kolorów
    const byColor = {};
    for (const d of downloaded) {
      const c = d.color || 'Bez koloru';
      if (!byColor[c]) byColor[c] = [];
      byColor[c].push(d);
    }
    for (const color of Object.keys(byColor).sort()) {
      reportLines.push(`  [${color}]`);
      byColor[color].forEach(d => {
        const countStr = d.count > 1 ? ` x${d.count}` : '';
        const prettyStr = d.prettyName ? ` - ${d.prettyName}` : '';
        const moduleStr = d.module ? ` (${d.module})` : '';
        reportLines.push(`    • ${d.name}${prettyStr}${countStr}${moduleStr}`);
      });
    }
  }
  reportLines.push('');
  
  // Pobrane pliki UCANCAM
  reportLines.push(`📦 POBRANE PLIKI UCANCAM: ${ucancamDownloaded.length}`);
  reportLines.push('-'.repeat(40));
  if (ucancamDownloaded.length > 0) {
    // Grupuj według źródła
    const bySource = {};
    for (const u of ucancamDownloaded) {
      const src = u.source || 'nieznane';
      if (!bySource[src]) bySource[src] = [];
      bySource[src].push(u);
    }
    
    // Najpierw elementy z zestawu
    if (bySource['zestaw']) {
      reportLines.push(`  [Z zestawu]`);
      bySource['zestaw'].forEach(u => {
        reportLines.push(`    ✅ ${u.name} → ${u.fileName}`);
      });
      delete bySource['zestaw'];
    }
    
    // Potem z modułów
    for (const src of Object.keys(bySource).sort()) {
      reportLines.push(`  [Z modułu: ${src}]`);
      bySource[src].forEach(u => {
        reportLines.push(`    ✅ ${u.name} → ${u.fileName}`);
      });
    }
  }
  reportLines.push('');
  
  // Brakujące linki DXF
  if (missingLinks.length > 0) {
    reportLines.push(`❌ BRAK HIPERŁĄCZY DXF: ${missingLinks.length}`);
    reportLines.push('-'.repeat(40));
    missingLinks.forEach(m => {
      reportLines.push(`  • ${m}`);
    });
    reportLines.push('');
  }
  
  // Brakujące linki UCANCAM
  if (ucancamMissing.length > 0) {
    reportLines.push(`⚠️ BRAK LINKÓW UCANCAM: ${ucancamMissing.length}`);
    reportLines.push('-'.repeat(40));
    
    // Grupuj według źródła
    const missingBySource = {};
    for (const u of ucancamMissing) {
      const src = u.source || 'nieznane';
      if (!missingBySource[src]) missingBySource[src] = [];
      missingBySource[src].push(u.name);
    }
    
    if (missingBySource['zestaw']) {
      reportLines.push(`  [Z zestawu]`);
      missingBySource['zestaw'].forEach(n => {
        reportLines.push(`    • ${n}`);
      });
      delete missingBySource['zestaw'];
    }
    
    for (const src of Object.keys(missingBySource).sort()) {
      reportLines.push(`  [Z modułu: ${src}]`);
      missingBySource[src].forEach(n => {
        reportLines.push(`    • ${n}`);
      });
    }
    reportLines.push('');
  }
  
  // Błędy
  if (errors.length > 0 || ucancamErrors.length > 0) {
    reportLines.push(`⚠️ BŁĘDY: ${errors.length + ucancamErrors.length}`);
    reportLines.push('-'.repeat(40));
    errors.forEach(e => {
      reportLines.push(`  • ${e}`);
    });
    ucancamErrors.forEach(e => {
      reportLines.push(`  • UCANCAM: ${e}`);
    });
    reportLines.push('');
  }
  
  reportLines.push('═'.repeat(60));
  reportLines.push(`Wygenerowano automatycznie przez Rejestr Plików CNC`);
  
  // Zapisz plik raportu
  const reportFilename = `Raport_${setId}_${timestampForName()}.txt`;
  const reportContent = reportLines.join('\n');
  folder.createFile(reportFilename, reportContent, 'text/plain');

  // ═══════════════════════════════════════════════════════════════════════════
  // PODSUMOWANIE (dialog)
  // ═══════════════════════════════════════════════════════════════════════════
  
  const summary = [];
  summary.push(`📁 Folder: ${folderUrl}`);
  summary.push('');
  summary.push(`📄 Pobrano plików DXF: ${downloaded.length}`);
  
  // Info o UCANCAM
  if (ucancamDownloaded.length > 0) {
    summary.push('');
    summary.push(`📦 Pobrano plików UCANCAM: ${ucancamDownloaded.length}`);
    
    // Pokaż pierwsze 10, grupując według źródła
    const bySource = {};
    for (const u of ucancamDownloaded) {
      const src = u.source || 'nieznane';
      if (!bySource[src]) bySource[src] = [];
      bySource[src].push(u);
    }
    
    let shown = 0;
    const maxShow = 10;
    
    if (bySource['zestaw']) {
      summary.push('   [Z zestawu]');
      for (const u of bySource['zestaw']) {
        if (shown >= maxShow) break;
        summary.push(`     ✅ ${u.name} → ${u.fileName}`);
        shown++;
      }
      delete bySource['zestaw'];
    }
    
    for (const src of Object.keys(bySource).sort()) {
      if (shown >= maxShow) break;
      summary.push(`   [Z modułu: ${src}]`);
      for (const u of bySource[src]) {
        if (shown >= maxShow) break;
        summary.push(`     ✅ ${u.name} → ${u.fileName}`);
        shown++;
      }
    }
    
    if (ucancamDownloaded.length > maxShow) {
      summary.push(`   ... + ${ucancamDownloaded.length - maxShow} innych`);
    }
  }
  
  if (missingLinks.length) {
    summary.push('');
    summary.push(`❌ Brak hiperłączy DXF (${missingLinks.length}):`);
    missingLinks.slice(0, 10).forEach(m => summary.push(`   • ${m}`));
    if (missingLinks.length > 10) summary.push(`   ... + ${missingLinks.length - 10} innych`);
  }
  
  if (ucancamMissing.length > 0) {
    summary.push('');
    summary.push(`⚠️ Brak linków UCANCAM (${ucancamMissing.length}):`);
    ucancamMissing.slice(0, 10).forEach(u => {
      const sourceStr = u.source === 'zestaw' ? 'z zestawu' : `z ${u.source}`;
      summary.push(`   • ${u.name} (${sourceStr})`);
    });
    if (ucancamMissing.length > 10) {
      summary.push(`   ... + ${ucancamMissing.length - 10} innych`);
    }
  }
  
  if (errors.length || ucancamErrors.length) {
    summary.push('');
    summary.push(`⚠️ Błędy (${errors.length + ucancamErrors.length}):`);
    errors.slice(0, 5).forEach(e => summary.push(`   • ${e}`));
    ucancamErrors.slice(0, 5).forEach(e => summary.push(`   • UCANCAM: ${e}`));
    const totalErrors = errors.length + ucancamErrors.length;
    if (totalErrors > 10) summary.push(`   ... + ${totalErrors - 10} innych`);
  }

  if (ui) ui.alert('Zakończono', summary.join('\n'), ui.ButtonSet.OK);
  else Logger.log(summary.join('\n'));

  return { folderUrl, downloaded, missingLinks, errors, ucancamDownloaded, ucancamMissing, ucancamErrors };
}
