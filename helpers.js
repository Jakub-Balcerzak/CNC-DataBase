// Helpers and shared constants
const SHEET_ZESTAWY = 'Zestawy CNC';
const SHEET_MODULE = 'Moduły CNC';

const linkCache = {};

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
    const name = values[r][7] ? String(values[r][7]).trim() : ''; // kol. H (była G)

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

function isModuleName(name) {
  if (!name) return false;
  name = String(name).trim();
  return /^[MX]/i.test(name);
}

function isSetName(name) {
  if (!name) return false;
  name = String(name).trim();
  return /^P/i.test(name);
}

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

function sanitizeFileName(name) {
  return name.replace(/[\/\\\?\%\*\:\|\"<>\.]/g, '_').substring(0, 240);
}

function timestampForName() {
  const d = new Date();
  const pad = (n) => (n<10?'0':'')+n;
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function getOrCreateSubfolder(parentFolder, subfolderName) {
  const existing = parentFolder.getFoldersByName(subfolderName);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(subfolderName);
}

function createSummaryTxtFile(downloaded, folder, filename) {
  if (!downloaded || downloaded.length === 0) return null;

  // Group by element name but preserve first-seen module (if present)
  const grouped = new Map();
  for (const item of downloaded) {
    const key = `${item.name}`; // group by element only
    const existing = grouped.get(key);
    if (!existing) {
      grouped.set(key, {
        count: item.count || 1,
        namePretty: item.prettyName || '',
        name: item.name,
        module: item.module || ''
      });
    } else {
      existing.count += (item.count || 1);
      // keep existing.module (first seen)
    }
  }

  // Sort by module (if present) then by element name. Elements without module go after those with a module.
  const sorted = Array.from(grouped.values()).sort((a, b) => {
    const ma = (a.module || '').toString().trim();
    const mb = (b.module || '').toString().trim();
    const aHas = ma !== '';
    const bHas = mb !== '';
    if (aHas && !bHas) return -1;
    if (!aHas && bHas) return 1;
    if (aHas && bHas) {
      const cmp = ma.localeCompare(mb);
      if (cmp !== 0) return cmp;
    }
    return a.name.localeCompare(b.name);
  });
  const hasModule = sorted.some(s => s.module && String(s.module).trim() !== '');

  const lines = [];
  if (hasModule) {
    lines.push('Nr modułu | Nr kat. elementu       | Ilość | Nazwa elementu');
    lines.push('----------+------------------------+-------+-------------------------------------');
  } else {
    lines.push('Nr kat. elementu       | Ilość | Nazwa elementu');
    lines.push('------------------------+-------+-------------------------------------');
  }

  const padRight = (txt, len) => (txt.length >= len ? txt.substring(0, len) : txt + ' '.repeat(len - txt.length));
  const padLeft = (txt, len) => (txt.length >= len ? txt.substring(0, len) : ' '.repeat(len - txt.length) + txt);

  for (const el of sorted) {
    if (hasModule) {
      const mod = padRight(String(el.module || ''), 8);
      lines.push(`${mod} | ${padRight(el.name, 23)}| ${padLeft(String(el.count), 5)} | ${padRight(el.namePretty, 37)}`);
    } else {
      lines.push(`${padRight(el.name, 23)}| ${padLeft(String(el.count), 5)} | ${padRight(el.namePretty, 37)}`);
    }
  }

  const content = lines.join('\n');
  const fileName = filename || 'Podsumowanie_elementów.txt';
  const blob = Utilities.newBlob(content, 'text/plain', fileName);
  const file = folder.createFile(blob);
  return file.getUrl();
}

function collectElementTotals(name, multiplier, modulesMap, zestawyMap, totals, pathVisited, moduleMap, parentModule) {
  name = String(name).trim();
  if (!name) return;
  
  // If this is a module, descend into its children and mark current module as parent
  if (isModuleName(name)) {
    if (pathVisited[name]) {
      return;
    }
    pathVisited[name] = true;

    const children = modulesMap[name];
    if (!children || children.length === 0) {
      pathVisited[name] = false;
      return;
    }

    for (let ch of children) {
      const childCount = (typeof ch.count === 'number' && !isNaN(ch.count) && ch.count > 0) ? ch.count : 1;
      // pass current module name as parentModule for children
      collectElementTotals(ch.text, multiplier * childCount, modulesMap, zestawyMap, totals, pathVisited, moduleMap, name);
    }

    pathVisited[name] = false;
    return;
  }
  
  // If this is a nested set (starts with P), descend into its children
  if (isSetName(name)) {
    if (pathVisited[name]) {
      return;
    }
    pathVisited[name] = true;

    const children = zestawyMap[name];
    if (!children || children.length === 0) {
      pathVisited[name] = false;
      return;
    }

    for (let ch of children) {
      const childCount = (typeof ch.count === 'number' && !isNaN(ch.count) && ch.count > 0) ? ch.count : 1;
      // pass parentModule through for nested set children
      collectElementTotals(ch.text, multiplier * childCount, modulesMap, zestawyMap, totals, pathVisited, moduleMap, parentModule);
    }

    pathVisited[name] = false;
    return;
  }

  const add = Number(multiplier) || 1;
  totals[name] = (totals[name] || 0) + add;

  // Record module for element if provided and not already set
  if (moduleMap && parentModule) {
    if (!moduleMap[name]) moduleMap[name] = parentModule;
  }
}

function guessExtension(url, blob) {
  try {
    const m = url.match(/(\.[a-z0-9]{1,6})(?:[\?#]|$)/i);
    if (m && m[1]) {
      return m[1];
    }
    const ct = blob.getContentType();
    if (ct) {
      if (ct.indexOf('dxf') !== -1) return '.dxf';
      if (ct.indexOf('octet-stream') !== -1) return '';
      const parts = ct.split('/');
      if (parts.length > 1) {
        const subtype = parts[1].split('+')[0];
        return '.' + subtype;
      }
    }
  } catch (e) {}
  return '';
}
