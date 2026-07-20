const CSV_PATH = 'residuos.csv';

/* Copia de respaldo de residuos.csv: se usa solo si el archivo se abre
   directamente (doble clic) o se comparte sin la carpeta completa, casos
   en los que el navegador bloquea la carga del CSV externo. Cuando el
   sitio esté publicado (http/https), siempre se usa el residuos.csv real,
   así que sigue editando ese archivo en Excel como fuente principal. */
const CSV_FALLBACK = `nombre,tacho,sinonimos,nota
Papel impreso o de oficina,reciclable,papel|hoja|hoja impresa|papel de oficina|hojas|papel bond,Debe estar limpio y seco
Papel kraft,reciclable,kraft|papel kraft|bolsa kraft,Limpio y seco
Revista o periodico,reciclable,revista|periodico|diario,Limpio y seco
Caja de carton,reciclable,caja|carton|caja de carton|caja de courier|caja de delivery|caja de zapatos|cuaderno|libreta,Desarmala o doblala antes de botarla debe estar limpia y seca
Botella de plastico,reciclable,botella|botella de agua|botella de gaseosa|botella pet|agua|gaseosa|coca cola|inca kola|cielo|san luis|pet,Vaciala enjuagala y secala antes de botarla
Envase plastico de comida o taper,reciclable,tupper|taper|tapper|fiambrera|envase plastico|envase de comida|vaso de yogurt,Enjuagalo y sacalo debe quedar sin restos de comida
Plastico film o bolsa plastica,reciclable,film plastico|bolsa plastica|bolsa de plastico|wrap|plastico film|envoltura plastica transparente|bolsa del super,Limpio y seco
Plastico burbuja,reciclable,plastico burbuja|papel burbuja|burbujas|embalaje con burbujas,Limpio y seco
Cubiertos de plastico flexibles que no se rompen al doblar,reciclable,cubiertos flexibles|cubiertos que no se rompen|cubiertos doblables|cubiertos blandos,Hazles la prueba doblalos si no se rompen van aqui
Lata de bebida,reciclable,lata|lata de gaseosa|lata de cerveza|aluminio|lata de energizante|red bull,Enjuagala antes de botarla
Lata de conserva o atun,reciclable,lata de atun|lata de conserva|lata de sardina|conserva,Enjuagala antes de botarla
Envase o botella de vidrio entera,reciclable,frasco|frasco de vidrio|botella de vidrio|envase de vidrio|vidrio|copa,Debe estar entera no rota enjuagala
Tetra Pak,reciclable,caja de leche|tetrapak|tetra pak|envase de jugo|caja de jugo|leche|caja de leche|rehidratante,Enjuagalo y aplastalo
Restos de comida cruda o cocida,organico,comida|sobras|restos de almuerzo|restos de comida|comida sobrante|plato de comida,Al balde de organicos sin bolsa ni empaque
Cascara de fruta,organico,cascara|cascara de platano|cascara de naranja|cascara de manzana|piel de fruta|fruta,
Cascara de verdura,organico,cascara de papa|cascara de zanahoria|verdura|vegetal,
Cascara de huevo,organico,cascara de huevo|huevo|cascaron,
Restos carnicos,organico,hueso|huesos|carne|pollo|restos de carne|espina,
Borra de cafe,organico,borra de cafe|residuo de cafe|cafe molido usado|posos,Solo el molido de cafe sin la bolsa filtrante ni el hilo
Cubiertos o vajilla de bambu o caña,organico,cubiertos de bambu|cubiertos de caña|vajilla de bambu|palitos de bambu,
Vaso de cafe descartable,general,vaso de carton|vaso descartable|vaso starbucks|vaso cafe|cafe|starbucks|vaso de coffee|vaso to go|polipapel,Parece carton pero tiene una capa interna plastificada aunque diga biodegradable va a general
Empaque compostable o biodegradable,general,empaque biodegradable|empaque compostable|envase compostable|bioplastico,Si el interior se siente suave o plastificado al tacto no se recicla aunque diga compostable
Cubiertos de plastico descartable,general,cubiertos descartables|tenedor plastico|cuchara plastica|cuchillo plastico|cubiertos|cubiertos rigidos,Hazles la prueba doblalos si se rompen van a general si no se rompen van a reciclable
Bolsita de te usada,general,bolsa de te|te usado|saquito de te|filtrante,La bolsita y el hilo van a general si separas solo las hojas sueltas esas si van al balde de organicos
Servilleta usada,general,servilleta|papel servilleta|servilletas sucias,
Papel toalla usado,general,papel toalla|toalla de papel,
Papel higienico usado,general,papel higienico|papel baño,
Envoltura de snack sachet o salsa,general,envoltura|bolsa de snack|paquete de papas|sachet|sobre de salsa|envoltorio metalizado|piqueo|bolsa de piqueo,
Plato o vaso de tecnopor,general,tecnopor|plato tecnopor|icopor|vaso tecnopor|poliestireno|plastico 6|plastico numero 6,
Envase acrilico o plastico rigido tipo 7,general,acrilico|plastico rigido transparente|plastico numero 7|plastico 7,
Vaso plato o vidrio roto,general,vaso roto|plato roto|vidrio roto|copa rota|taza rota|ceramica rota,Por seguridad el vidrio o ceramica rota va a general nunca a reciclables envuelvelo antes de botarlo
Lapicero o plumon,general,lapicero|lapiceros|plumon|plumones|boligrafo,
Capsula de cafe,general,capsula de cafe|capsula nespresso|capsula dolce gusto,
Voucher o comprobante termico,general,voucher|boleta termica|ticket|comprobante,
Globo,general,globo|globos,
Mascarilla o guantes descartables,general,mascarilla|tapabocas|guantes|guantes de latex,Por higiene siempre van a general
Colillas de cigarro,general,colilla|cigarro|ceniza,
Taper de delivery con comida y envoltorios,reciclable,taper de delivery|delivery|caja de delivery con comida,Separalo en tres la comida al balde de organicos las servilletas cubiertos rigidos y sachets a general y el taper plastico enjuagado y seco a reciclable
Pilas y baterias,especial,pila|bateria|pilas alcalinas|bateria de laptop,No las boto en los tachos de Sinba llevalas a un punto de acopio de RAEE
Aparatos electronicos,especial,electronico|cargador|cable|audifonos|mouse|teclado|celular malogrado|laptop malograda|raee,Los residuos electronicos necesitan un punto de acopio especial no van con la recoleccion de Sinba
Toner o cartucho de impresora,especial,toner|cartucho|tinta de impresora,Muchas marcas tienen programa de retorno de cartuchos
Focos y lamparas,especial,foco|fluorescente|lampara|led fundido,Contienen materiales que requieren manejo especial
Medicamentos vencidos,especial,medicina|pastillas|medicamento vencido,Llevalos a un punto de acopio de farmacia`;

const BIN_LABELS = {
  organico: 'Orgánicos',
  general: 'Generales',
  reciclable: 'Reciclables',
  especial: 'Punto especial',
};

let ITEMS = [];

const searchInput = document.getElementById('search-input');
const searchForm = document.getElementById('search-form');
const resultArea = document.getElementById('result-area');
const idleHint = document.getElementById('idle-hint');

init();

async function init() {
  try {
    if (location.protocol === 'file:') throw new Error('file-protocol');
    const res = await fetch(CSV_PATH);
    if (!res.ok) throw new Error('fetch-failed');
    const text = await res.text();
    ITEMS = parseCSV(text);
  } catch (err) {
    ITEMS = parseCSV(CSV_FALLBACK);
  }

  searchInput.addEventListener('input', () => {
    const value = searchInput.value.trim();
    if (value.length === 0) {
      renderIdle();
      return;
    }
    runSearch(value);
  });

  searchForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const value = searchInput.value.trim();
    if (value.length > 0) runSearch(value);
  });

  document.querySelectorAll('.chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      searchInput.value = chip.dataset.example;
      runSearch(chip.dataset.example);
      searchInput.focus();
    });
  });
}

/* ---------- CSV parsing (handles simple quoted fields) ---------- */
function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
  const header = splitCSVLine(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCSVLine(lines[i]);
    if (cols.length < header.length) continue;
    const row = {};
    header.forEach((h, idx) => { row[h.trim()] = (cols[idx] || '').trim(); });
    rows.push({
      nombre: row.nombre,
      tacho: row.tacho,
      nota: row.nota || '',
      sinonimos: (row.sinonimos || '')
        .split('|')
        .map((s) => s.trim())
        .filter(Boolean),
    });
  }
  return rows;
}

function splitCSVLine(line) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      inQuotes = !inQuotes;
    } else if (c === ',' && !inQuotes) {
      result.push(cur);
      cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur);
  return result;
}

/* ---------- Text normalization ---------- */
function normalize(str) {
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* ---------- Levenshtein distance ---------- */
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return dp[m][n];
}

/* ---------- Scoring a query against one candidate phrase ---------- */
function scorePhrase(query, phrase) {
  const nq = normalize(query);
  const np = normalize(phrase);
  if (!nq || !np) return 0;

  if (nq === np) return 1000;
  if (np.includes(nq) || nq.includes(np)) return 800;

  const qWords = nq.split(' ');
  const pWords = np.split(' ');
  let bestWordScore = 0;

  for (const qw of qWords) {
    for (const pw of pWords) {
      if (qw === pw) { bestWordScore = Math.max(bestWordScore, 700); continue; }
      if (qw.length < 3 || pw.length < 3) continue;
      const dist = levenshtein(qw, pw);
      const maxLen = Math.max(qw.length, pw.length);
      const allowedDist = maxLen <= 4 ? 1 : 2;
      if (dist <= allowedDist) {
        const score = 600 - dist * 80 - Math.abs(qw.length - pw.length) * 5;
        bestWordScore = Math.max(bestWordScore, score);
      }
    }
  }
  return bestWordScore;
}

function scoreItem(query, item) {
  let best = scorePhrase(query, item.nombre);
  for (const syn of item.sinonimos) {
    best = Math.max(best, scorePhrase(query, syn));
  }
  return best;
}

/* ---------- Search + render ---------- */
function runSearch(query) {
  const scored = ITEMS
    .map((item) => ({ item, score: scoreItem(query, item) }))
    .filter((s) => s.score > 0)
    .sort((a, b) => b.score - a.score);

  if (scored.length === 0) {
    renderNoResult(query);
    return;
  }

  const top = scored[0];
  const alternatives = [];
  const seenNames = new Set([top.item.nombre]);
  for (let i = 1; i < scored.length && alternatives.length < 3; i++) {
    const cand = scored[i];
    if (seenNames.has(cand.item.nombre)) continue;
    if (top.score - cand.score > 300) break;
    seenNames.add(cand.item.nombre);
    alternatives.push(cand.item);
  }

  renderResult(top.item, alternatives);
}

function renderIdle() {
  resultArea.innerHTML = '';
  idleHint.classList.remove('hidden');
}

function renderResult(item, alternatives) {
  idleHint.classList.add('hidden');
  const label = BIN_LABELS[item.tacho] || item.tacho;

  let html = `
    <div class="result-card result-card--${item.tacho}">
      <span class="result-card__badge"><span class="result-card__dot"></span>${label}</span>
      <p class="result-card__title">${escapeHTML(item.nombre)}</p>
      ${item.nota ? `<p class="result-card__note">${escapeHTML(item.nota)}</p>` : ''}
    </div>
  `;

  if (alternatives.length > 0) {
    html += `
      <div class="suggestions">
        <p class="suggestions__label">¿O buscabas...?</p>
        <div class="suggestions__list">
          ${alternatives.map((alt) => `<button type="button" class="suggestion" data-name="${escapeHTML(alt.nombre)}">${escapeHTML(alt.nombre)}</button>`).join('')}
        </div>
      </div>
    `;
  }

  resultArea.innerHTML = html;

  resultArea.querySelectorAll('.suggestion').forEach((btn) => {
    btn.addEventListener('click', () => {
      searchInput.value = btn.dataset.name;
      runSearch(btn.dataset.name);
    });
  });
}

function renderNoResult(query) {
  idleHint.classList.add('hidden');
  resultArea.innerHTML = `
    <div class="no-result">
      <p class="no-result__title">No encontramos "${escapeHTML(query)}" todavía</p>
      <ul>
        <li><strong>Orgánicos:</strong> restos de comida, cáscaras, posos de café.</li>
        <li><strong>Generales:</strong> todo lo sucio, mixto o descartable (vasos de café, tecnopor, mascarillas).</li>
        <li><strong>Reciclables:</strong> botellas, latas, papel y cartón limpios y secos.</li>
      </ul>
      <a href="mailto:micaela.paez.a@astara.com?cc=alejandra.arroyo.d@astara.com&subject=Residuo%20no%20encontrado&body=No%20encontr%C3%A9%3A%20${encodeURIComponent(query)}">Escríbenos sobre "${escapeHTML(query)}"</a>
    </div>
  `;
}

function escapeHTML(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
