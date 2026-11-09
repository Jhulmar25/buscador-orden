const get = id => document.getElementById(id);
const input = get('document');
const suggestions = get('suggestions');
const statusMessage = get('status');
const result = get('result');
const title = get('result-title');

const excelFiles = [
  'BRIGADA_CANINA.xlsx',
  'CCO.xlsx',
  'CONDUCTORES_DE_BUS.xlsx',
  'DRONES.xlsx',
  'GAES.xlsx',
  'GODOFREDO.xlsx',
  'PERSONAL.xlsx',
  'SECTOR_1.xlsx',
  'SECTOR_2.xlsx',
  'SECTOR_3.xlsx',
  'SECTOR_4.xlsx',
  'SECTOR_5.xlsx'
];

const byDocument = new Map();
let documents = [];

function normalizeName(value) {
  return value.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase();
}

function names(rows) {
  return [...new Set(rows.map(row => row.name))].join(' / ');
}

function showResult(documentNumber) {
  const rows = byDocument.get(documentNumber);
  input.value = documentNumber;
  suggestions.replaceChildren();
  statusMessage.textContent = '';
  title.textContent = names(rows);

  const fileCount = new Set(rows.map(row => row.file)).size;
  get('result-summary').textContent =
    `DNI/CE ${documentNumber} · ${rows.length} registro(s) en ${fileCount} archivo(s)`;

  const body = get('matches');
  body.replaceChildren();

  for (const { file, name, number } of rows) {
    const row = body.insertRow();
    row.insertCell().textContent = file;
    row.insertCell().textContent = name;
    row.insertCell().textContent = number || 'Sin correlativo';
  }

  result.hidden = false;
  title.focus();
}

function search() {
  result.hidden = true;
  suggestions.replaceChildren();
  const query = input.value.trim();

  if (!query) {
    statusMessage.textContent = '';
    return;
  }

  const isDocument = /^[0-9]+$/.test(query);
  const terms = normalizeName(query).split(/\s+/);

  const matches = documents.filter(number => {
    if (isDocument) return number.startsWith(query);

    return byDocument.get(number).some(({ name }) => {
      const normalized = normalizeName(name);
      return terms.every(term => normalized.includes(term));
    });
  });

  if (matches.length === 0) {
    statusMessage.textContent = 'No se encontraron coincidencias.';
    return;
  }

  statusMessage.textContent = `${matches.length} coincidencia(s).`;
  if (matches.length > 12) {
    statusMessage.textContent +=
      ' Se muestran las primeras 12. Escribe más para acotar la búsqueda.';
  }

  for (const number of matches.slice(0, 12)) {
    const rows = byDocument.get(number);
    const item = document.createElement('li');
    const button = document.createElement('button');

    button.type = 'button';
    button.textContent = `${number} · ${names(rows)} · ${rows.length} registro(s)`;
    button.addEventListener('click', () => showResult(number));

    item.append(button);
    suggestions.append(item);
  }
}

input.addEventListener('input', search);

get('search').addEventListener('submit', event => {
  event.preventDefault();
  const query = input.value.trim();

  if (byDocument.has(query)) {
    showResult(query);
  } else {
    suggestions.querySelector('button')?.focus();
  }
});

get('search').addEventListener('reset', event => {
  event.preventDefault();
  input.value = '';
  search();
  input.focus();
});

// Los libros XLSX contienen las hojas como archivos XML dentro de un ZIP.
function openWorkbook(buffer) {
  const data = new DataView(buffer);
  const decoder = new TextDecoder();
  const entries = new Map();
  const minimum = Math.max(0, data.byteLength - 65557);
  let end = data.byteLength - 22;

  while (
    end >= minimum &&
    !(data.getUint32(end, true) === 0x06054b50 &&
      end + 22 + data.getUint16(end + 20, true) === data.byteLength)
  ) {
    end--;
  }

  if (end < minimum) throw new Error('El archivo no es un Excel válido.');

  let offset = data.getUint32(end + 16, true);
  const entryCount = data.getUint16(end + 10, true);

  for (let i = 0; i < entryCount; i++) {
    if (data.getUint32(offset, true) !== 0x02014b50) {
      throw new Error('Índice ZIP inválido.');
    }

    const nameLength = data.getUint16(offset + 28, true);
    const name = decoder.decode(
      new Uint8Array(buffer, offset + 46, nameLength)
    );

    entries.set(name, {
      method: data.getUint16(offset + 10, true),
      length: data.getUint32(offset + 20, true),
      offset: data.getUint32(offset + 42, true)
    });

    offset += 46 + nameLength
      + data.getUint16(offset + 30, true)
      + data.getUint16(offset + 32, true);
  }

  return async path => {
    const entry = entries.get(path);
    if (!entry) throw new Error(`Falta ${path} en el libro.`);

    const start = entry.offset + 30
      + data.getUint16(entry.offset + 26, true)
      + data.getUint16(entry.offset + 28, true);
    const bytes = new Uint8Array(buffer, start, entry.length);
    let xml;

    if (entry.method === 8) {
      const stream = new Blob([bytes]).stream()
        .pipeThrough(new DecompressionStream('deflate-raw'));
      xml = await new Response(stream).text();
    } else if (entry.method === 0) {
      xml = decoder.decode(bytes);
    } else {
      throw new Error('Compresión del Excel no compatible.');
    }

    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    if (parsed.querySelector('parsererror')) {
      throw new Error('Contenido XML inválido.');
    }

    return parsed;
  };
}

const elements = (node, tag) => [...node.getElementsByTagNameNS('*', tag)];

function textRuns(node) {
  return elements(node, 't').map(element => element.textContent).join('');
}

function headerKey(value) {
  return value.toUpperCase().replace(/[^A-Z0-9]/g, '');
}

async function readExcel(file, buffer) {
  const xml = openWorkbook(buffer);
  const workbook = await xml('xl/workbook.xml');
  const rels = await xml('xl/_rels/workbook.xml.rels');
  const relations = elements(rels, 'Relationship');
  const byId = new Map(relations.map(rel => [rel.getAttribute('Id'), rel]));

  function targetPath(rel) {
    return new URL(
      rel.getAttribute('Target'),
      'https://local.invalid/xl/workbook.xml'
    ).pathname.slice(1);
  }

  const stringsRel = relations.find(rel =>
    rel.getAttribute('Type').endsWith('/sharedStrings')
  );
  const strings = stringsRel
    ? elements(await xml(targetPath(stringsRel)), 'si').map(textRuns)
    : [];

  const stylesRel = relations.find(rel =>
    rel.getAttribute('Type').endsWith('/styles')
  );
  const styles = stylesRel ? await xml(targetPath(stylesRel)) : null;
  const formats = new Map(
    (styles ? elements(styles, 'numFmt') : []).map(node => [
      node.getAttribute('numFmtId'),
      node.getAttribute('formatCode')
    ])
  );
  const cellStyles = [
    ...(styles?.getElementsByTagNameNS('*', 'cellXfs')[0]?.children ?? [])
  ];

  function valueOf(cell) {
    const type = cell.getAttribute('t');
    const value = elements(cell, 'v')[0]?.textContent ?? '';

    if (type === 'inlineStr') return textRuns(cell);
    if (type === 's') return strings[Number(value)] ?? '';

    const style = cellStyles[Number(cell.getAttribute('s'))];
    const format = formats.get(style?.getAttribute('numFmtId'));

    if ((type === null || type === 'n') && /^\d+$/.test(value) &&
        /^0+$/.test(format ?? '')) {
      return value.padStart(format.length, '0');
    }

    return value;
  }

  const records = [];
  let recognized = false;

  for (const sheet of elements(workbook, 'sheet')) {
    const id = [...sheet.attributes].find(attr => attr.localName === 'id')?.value;
    const relation = byId.get(id);
    if (!relation?.getAttribute('Type').endsWith('/worksheet')) continue;

    const page = await xml(targetPath(relation));
    let columns;

    for (const row of elements(page, 'row')) {
      const cells = new Map(elements(row, 'c').map(cell => [
        cell.getAttribute('r').replace(/\d+$/, ''),
        valueOf(cell).trim()
      ]));

      const headers = new Map(
        [...cells].map(([column, value]) => [headerKey(value), column])
      );
      const docColumn = headers.get('DNICE') ?? headers.get('DNI');
      const nameColumn = headers.get('APELLIDOSYNOMBRES');
      const numberColumn = headers.get('N');

      if (docColumn && nameColumn && numberColumn) {
        columns = [docColumn, nameColumn, numberColumn];
        recognized = true;
        continue;
      }

      if (!columns) continue;

      const [doc, name, number] = columns.map(column => cells.get(column) ?? '');
      if (/^[0-9]+$/.test(doc)) {
        records.push({ doc, name, file, number });
      }
    }
  }

  if (!recognized) {
    throw new Error('No se encontraron las columnas DNI/CE, apellidos y nombres y N°.');
  }

  return records;
}

async function loadExcelFiles() {
  get('summary').textContent = 'Leyendo archivos Excel…';
  get('load-errors').textContent = '';

  const results = await Promise.allSettled(excelFiles.map(async file => {
    const response = await fetch(encodeURIComponent(file), { cache: 'no-store' });

    if (!response.ok) {
      throw new Error(`No se pudo abrir el archivo (HTTP ${response.status}).`);
    }

    return readExcel(file, await response.arrayBuffer());
  }));

  let loaded = 0;
  let count = 0;
  const errors = [];

  results.forEach((outcome, index) => {
    if (outcome.status === 'rejected') {
      errors.push(`${excelFiles[index]}: ${outcome.reason.message}`);
      return;
    }

    loaded++;
    count += outcome.value.length;

    for (const record of outcome.value) {
      if (!byDocument.has(record.doc)) byDocument.set(record.doc, []);
      byDocument.get(record.doc).push(record);
    }
  });

  documents = [...byDocument.keys()].sort();
  get('summary').textContent =
    `${loaded} de ${excelFiles.length} archivos Excel · ${count} registros`;
  get('load-errors').textContent = errors.join(' · ');
  input.disabled = count === 0;
  if (count) input.focus();
}

if (location.protocol === 'http:' || location.protocol === 'https:') {
  loadExcelFiles();
} else {
  get('summary').textContent =
    'Abre index.html con Live Server para cargar automáticamente los Excel.';
}