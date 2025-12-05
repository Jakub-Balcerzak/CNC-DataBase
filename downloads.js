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
  // POBIERANIE PLIKU UCANCAM
  // ═══════════════════════════════════════════════════════════════════════════
  
  const COL_UCANCAM = 5; // kolumna F (0-based index)
  let ucancamDownloaded = null;
  let ucancamError = null;
  let ucancamMissing = false;

  // Szukaj linku UCANCAM dla modułu w arkuszu 'Moduły CNC'
  // Bierzemy pierwszy znaleziony wiersz z tym modułem
  let ucancamLink = null;
  for (let i = 1; i < modValues.length; i++) {
    const rowModName = String(modValues[i][0]).trim();
    if (rowModName === modId) {
      // Sprawdź czy jest link w kolumnie F
      try {
        const richCell = modRich[i][COL_UCANCAM];
        if (richCell && typeof richCell.getLinkUrl === 'function') {
          ucancamLink = richCell.getLinkUrl() || null;
        }
      } catch (e) {
        ucancamLink = null;
      }
      if (ucancamLink) break; // znaleziono link, przerywamy
    }
  }

  if (ucancamLink) {
    try {
      const fileIdMatch = ucancamLink.match(/[-\w]{25,}/);
      if (fileIdMatch) {
        const fileId = fileIdMatch[0];
        const sourceFile = DriveApp.getFileById(fileId);
        const originalFileName = sourceFile.getName();
        
        // Utwórz podfolder UCANCAM
        const ucancamFolder = folder.createFolder('UCANCAM');
        
        // Skopiuj plik z oryginalną nazwą
        sourceFile.makeCopy(originalFileName, ucancamFolder);
        ucancamDownloaded = originalFileName;
      } else {
        ucancamError = 'Nieprawidłowy format linku UCANCAM';
      }
    } catch (e) {
      ucancamError = e.message;
    }
  } else {
    ucancamMissing = true;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // PODSUMOWANIE
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
  if (ucancamDownloaded) {
    summaryLines.push(`📦 UCANCAM: ✅ ${ucancamDownloaded}`);
  } else if (ucancamError) {
    summaryLines.push(`📦 UCANCAM: ❌ Błąd - ${ucancamError}`);
  } else if (ucancamMissing) {
    summaryLines.push(`📦 UCANCAM: ⚠️ Brak linku dla modułu ${modId}`);
  }
  
  if (missingLinks.length) {
    summaryLines.push('');
    summaryLines.push(`❌ Elementy bez hiperłącza DXF (${missingLinks.length}):`);
    missingLinks.forEach(m => summaryLines.push(`   • ${m}`));
  }
  if (errors.length) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Błędy (${errors.length}):`);
    errors.forEach(err => summaryLines.push(`   • ${err}`));
  }
  if (dataWarnings.length) {
    summaryLines.push('');
    summaryLines.push(`⚠️ Ostrzeżenia dotyczące danych (${dataWarnings.length}):`);
    dataWarnings.forEach(w => summaryLines.push(`   • ${w}`));
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

    // 1) totals
    const totals = {};
    const pathVisited = {};
    const moduleMap = {};
    for (let e of startElements) {
      const topCount = (typeof e.count === 'number' && !isNaN(e.count) && e.count > 0) ? e.count : 1;
      collectElementTotals(e.text, topCount, modulesMap, zestawyMap, totals, pathVisited, moduleMap);
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
    const copyJobs = []; // { name, fileId, color, count }
    const fetchJobs = []; // { name, url, color, count }

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
        // try to use link as direct download
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
        // makeCopy into target folder
        src.makeCopy(fileName, colorFolder);
  downloaded.push({ name: job.name, color: job.color, prettyName: dataMap[job.name]?.name || '', count: job.count, module: moduleMap[job.name] || '' });
      } catch (e) {
        errors.push(`${job.name}: (copy) ${e.message}`);
        // fallback: try to add to fetchJobs using direct download URL constructed from id
        try {
          const dl = `https://drive.google.com/uc?export=download&id=${job.fileId}`;
          fetchJobs.push({ name: job.name, url: dl, color: job.color, count: job.count });
        } catch (e2) {
          Logger.log('Fallback enqueue failed for ' + job.name + ': ' + e2.message);
        }
      }
    }

    // 4) Fetch remaining via UrlFetchApp.fetchAll in batches
    const BATCH = 20; // reasonable batch size
    for (let i = 0; i < fetchJobs.length; i += BATCH) {
      const slice = fetchJobs.slice(i, i + BATCH);
      const requests = slice.map(j => ({ url: j.url, muteHttpExceptions: true, followRedirects: true }));
      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(requests);
      } catch (e) {
        // in case fetchAll fails entirely, fallback to sequential fetch
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

    // 5) Summary
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

    if (ui) ui.alert('Zakończono', summary.join('\n'), ui.ButtonSet.OK);
    else Logger.log(summary.join('\n'));

    return { folderUrl, downloaded, missingLinks, errors };
}
