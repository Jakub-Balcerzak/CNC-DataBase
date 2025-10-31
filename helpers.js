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
    const name = values[r][6] ? String(values[r][6]).trim() : ''; // kol. G

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

function createSummaryTxtFile(downloaded, folder) {
  if (!downloaded || downloaded.length === 0) return null;

  const grouped = new Map();
  for (const item of downloaded) {
    const key = `${item.name}||${item.color || 'Bez koloru'}`;
    const g = grouped.get(key) || { count: 0, namePretty: item.prettyName || '', color: item.color || 'Bez koloru', name: item.name };
    g.count += (item.count || 1);
    grouped.set(key, g);
  }

  const sorted = Array.from(grouped.values()).sort((a, b) => a.name.localeCompare(b.name));
  const lines = [];
  lines.push('Nr kat. elementu       | Ilość | Nazwa elementu                       | Kolor');
  lines.push('------------------------+--------+-------------------------------------+------------');

  const padRight = (txt, len) => (txt.length >= len ? txt.substring(0, len) : txt + ' '.repeat(len - txt.length));
  const padLeft = (txt, len) => (txt.length >= len ? txt.substring(0, len) : ' '.repeat(len - txt.length) + txt);

  for (const el of sorted) {
    lines.push(
      `${padRight(el.name, 23)}| ${padLeft(String(el.count), 6)}| ${padRight(el.namePretty, 37)}| ${el.color}`
    );
  }

  const content = lines.join('\n');
  const blob = Utilities.newBlob(content, 'text/plain', 'Podsumowanie_elementów.txt');
  const file = folder.createFile(blob);
  return file.getUrl();
}

function collectElementTotals(name, multiplier, modulesMap, zestawyMap, totals, pathVisited) {
  name = String(name).trim();
  if (!name) return;

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
      collectElementTotals(ch.text, multiplier * childCount, modulesMap, zestawyMap, totals, pathVisited);
    }

    pathVisited[name] = false;
    return;
  }

  const add = Number(multiplier) || 1;
  totals[name] = (totals[name] || 0) + add;
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
