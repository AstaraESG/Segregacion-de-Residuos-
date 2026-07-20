const CSV_PATH = 'residuos.csv';

/* Copia de respaldo de residuos.csv: se usa solo si el archivo se abre
   directamente (doble clic) o se comparte sin la carpeta completa, casos
   en los que el navegador bloquea la carga del CSV externo. Cuando el
   sitio esté publicado (http/https), siempre se usa el residuos.csv real,
   así que sigue editando ese archivo en Excel como fuente principal. */
const CSV_FALLBACK = `nombre,tacho,sinonimos,nota
Papel impreso o de oficina,reciclable,papel|hoja|hoja impresa|papel de oficina|hojas|papel bond|papel de colores|agenda|apuntes|documento|impresion|hoja membretada|hoja rayada|hoja cuadriculada|hoja reciclada|instructivo impreso|guia de remision|factura|libro viejo,Debe estar limpio y seco.
Papel kraft,reciclable,kraft|papel kraft|bolsa kraft,Debe estar limpio y seco.
Revista o periódico,reciclable,revista|periodico|diario|folleto|revistero,Debe estar limpio y seco.
Caja de cartón,reciclable,caja|carton|caja de carton|caja de courier|caja de delivery|caja de zapatos|cuaderno|libreta|caja de pastel|caja de torta|bandeja de carton|casillero de carton|esquinero de carton|mini caja|tubo de carton|rollo de carton|cartulina,Desármala o dóblala para ahorrar espacio. Debe estar limpia y seca.
Sobre o carpeta manila,reciclable,sobre manila|carpeta manila|folder manila|sobre,Debe estar limpio y seco. Retira el plástico de la ventanilla si lo tiene.
Nota adhesiva o post-it,reciclable,post it|post-it|nota adhesiva|notas adhesivas,"El papel se recicla. Antes de botarlo, arranca la pequeña tira con pegamento: esa parte va a generales."
Documento anillado o empastado,reciclable,anillado|documento anillado|empastado|bloc espiral,"Retira el espiral o gancho metálico y ponlo en generales. Las hojas de papel van a reciclables."
"Calendario, agenda vencida o bloc de notas",reciclable,calendario|agenda vencida|bloc de notas,"Si tiene espiral metálico o tapa plástica, retíralo antes: ese va a generales."
Tetra Pak,reciclable,caja de leche|tetrapak|tetra pak|envase de jugo|caja de jugo|leche|caja de leche|rehidratante|empaque tetra pak|empaque de jugo,"Ábrelo, enjuágalo y aplástalo antes de botarlo."
Envase o botella de vidrio entera,reciclable,frasco|frasco de vidrio|botella de vidrio|envase de vidrio|vidrio|copa|frasco vacio|frasco de perfume|frasco de crema,"Debe estar entera, no rota. Enjuágala. Por seguridad, el vidrio roto nunca se recicla."
Lata de bebida,reciclable,lata|lata de gaseosa|lata de cerveza|aluminio|lata de energizante|red bull|chapa|corona metalica,"Enjuágala bien, sin restos de grasa ni aceite."
Lata de conserva o atún,reciclable,lata de atun|lata de conserva|lata de sardina|conserva,"Enjuágala bien, sin restos de grasa ni aceite."
Cable o perfil de plástico,reciclable,cable|pvc|plastico 3|perfil plastico,"Este tipo de plástico (código 3) se recicla, para cables y perfiles específicos."
Plástico film o bolsa plástica limpia,reciclable,film plastico|bolsa plastica|bolsa de plastico|wrap|plastico film|envoltura plastica transparente|bolsa del super|plastico 4|ldpe|bolsa ziploc|bolsa de hielo|lamina plastica|forro plastico|resma vacia|celofan,Debe estar limpio y seco.
Plástico de burbujas,reciclable,plastico burbuja|papel burbuja|burbujas|embalaje con burbujas,Debe estar limpio y seco.
Botella de plástico,reciclable,botella|botella de agua|botella de gaseosa|botella grande|botella pet|agua|agua mineral|gaseosa|coca cola|inca kola|cielo|san luis|pet|plastico 1|bebida energetica|bebida isotonica|bebida proteica|refresco,"Vacíala y enjuágala antes de botarla."
Envase plástico de limpieza o shampoo,reciclable,envase de shampoo|envase de detergente|envase de limpieza|shampoo|detergente|hdpe|plastico 2,Enjuágalo si tuvo contacto con comida.
Envase plástico de comida o táper,reciclable,tupper|taper|tapper|fiambrera|envase plastico|envase de comida|vaso de yogurt|envase de yogurt|envase de helado|envase de crema|bandeja de plastico|comida para llevar|vaso de plastico|pp|plastico 5|envase,"Enjuágalo, sin restos de comida. Si tienes dudas, revisa el número dentro del triángulo en la base: 1, 2, 3, 4 o 5 se reciclan; 6 o 7 van a generales."
Tapa de botella,reciclable,tapa|tapa de botella|tapa plastica|chapita|corcholata,"Va por separado, no la dejes puesta en la botella."
Cubiertos de plástico flexibles que no se rompen al doblar,reciclable,cubiertos flexibles|cubiertos que no se rompen|cubiertos doblables|cubiertos blandos,"Hazles la prueba: dóblalos. Si se doblan sin romperse, van aquí."
Sorbete / cañita / pajita,reciclable,sorbete|sorbetes|cañita|cañitas|pajita|pajitas|popote|pitillo,"Es plástico tipo 5, igual que la tapa de botella. Debe estar limpio."
Restos de comida cruda o cocida,organico,comida|sobras|restos de almuerzo|restos de comida|comida sobrante|plato de comida|arroz|comida de casa|pan dulce|pan integral|pan frances|sandwich,"Va directo al balde de orgánicos, sin bolsa ni empaque."
Cáscara de fruta,organico,cascara|cascara de platano|cascara de naranja|cascara de manzana|cascara de mandarina|cascara de mango|cascara de piña|cascara de palta|cascara de aguacate|piel de fruta|fruta,
Cáscara de verdura,organico,cascara de papa|cascara de zanahoria|verdura|vegetal|semillas,
Cáscara de huevo,organico,cascara de huevo|huevo|cascaron,
Restos cárnicos y huesos,organico,hueso|huesos|carne|pollo|pescado|restos de carne|espina,Los huesos de cualquier tamaño van al balde de orgánicos.
Borra de café,organico,borra de cafe|residuo de cafe|cafe molido usado|posos|orujo|filtro de cafe,"Solo el molido o filtro de papel con café. La bolsita de té y su hilo van a generales."
Bolsita de té usada,organico,bolsa de te|te usado|saquito de te|filtrante,"Corta la bolsita: solo las hojas de té van al balde de orgánicos. La grapa, el hilo y el papel exterior van a generales."
Cubiertos o vajilla de bambú / madera / caña,organico,cubiertos de bambu|cubiertos de madera|cubiertos de caña|vajilla de bambu|palitos chinos|palitos de bambu|envase de fibra natural|palito de helado,
Removedor o palito para café,general,removedor|palito de cafe|agitador de cafe|palito removedor,"Si es de madera, va a orgánicos. Si es de plástico, va a generales."
Vaso de café descartable,general,vaso de carton|vaso descartable|vaso starbucks|vaso cafe|cafe|starbucks|vaso de coffee|vaso to go|polipapel|vaso de maquina|vaso de cafe de maquina|taza descartable,"Por fuera parece cartón, pero por dentro tiene una capa de plástico fundida para que no filtre líquido. Aunque diga biodegradable o compostable, esa mezcla no se puede separar. Siempre va a generales."
Empaque biodegradable o compostable,general,empaque biodegradable|empaque compostable|envase compostable|bioplastico|compostable|biodegradable,"Si el interior se siente suave o plastificado al tacto, no se recicla aunque el empaque lo indique."
Cubiertos de plástico descartable,general,cubiertos descartables|tenedor plastico|cuchara plastica|cuchillo plastico|cubiertos|cubiertos rigidos|cubiertos de cumpleaños|cubiertos de fiesta|cuchara|cuchillo|tenedor|cuchara descartable|cuchillo descartable,"Hazles la prueba: dóblalos. Si se rompen, van a generales. Si se doblan sin romperse, van a reciclables."
Servilleta usada,general,servilleta|papel servilleta|servilletas sucias,El papel con grasa o restos de comida no se recicla.
Papel toalla usado,general,papel toalla|toalla de papel,
Papel higiénico usado,general,papel higienico|papel baño,
Envoltura o sachet de snack y salsas,general,envoltura|empaque de snack|bolsa de snack|paquete de papas|paquete de galletas|sachet|sobre de salsa|mayonesa|ketchup|envoltorio metalizado|piqueo|bolsa de piqueo|snack|snacks|empaque de sandwich|empaque de chocolate|empaque de cafe|empaque de te|empaque de caramelos|empaque de alfajor|empaque de brownie|empaque de muffin|empaque de dona|empaque de croissant|empaque de galleton|empaque de frutos secos|empaque de cancha|empaque de gomitas|empaque de chicle|empaque de barra energetica|empaque de avena|empaque de cereal|empaque de granola|chocolate|caramelo|caramelo envuelto|galleta|dulce|postre|cupcake|malvavisco|wrapper|barra de cereal|paquete de azucar|paquete de sal|paquete de mostaza|paquete de aji|paquete de mermelada|paquete de mantequilla|bolsa de cafe|bolsa metalizada,Son laminados mixtos de plástico y aluminio que no se pueden separar.
Lata o tubo de snacks tipo Pringles,general,pringles|lata de pringles|tubo de papas,"Combina cartón, metal y plástico en una sola pieza, por eso no se puede separar."
Plato o vaso de tecnopor,general,tecnopor|plato tecnopor|icopor|vaso tecnopor|poliestireno|plastico 6|plastico numero 6|ps,Este tipo de plástico no se puede reciclar.
Plato descartable de cumpleaños o fiesta,general,plato de fiesta|plato de cumpleaños|plato plastico|plato descartable,
Envase acrílico o plástico rígido,general,acrilico|plastico rigido transparente|plastico numero 7|plastico 7|policarbonato|otros|visor plastico|blister,Este tipo de plástico es una mezcla que no se puede reciclar fácilmente.
Vouchers y boletas,general,voucher|boleta termica|ticket|comprobante|boleta|ticket de compra|recibo|ticket de estacionamiento|ticket de peaje|ticket de supermercado|tarjeta de embarque|comprobante pos,El papel térmico tiene químicos que contaminan el reciclaje de papel.
Lapicero o plumón,general,lapicero|lapiceros|plumon|plumones|boligrafo|corrector|corrector liquido|limpiatipo|lapiz|lapiz de color|marcador|marker|rotulador|resaltador|marcador permanente|marcador borrable|marcador fluorescente|portaminas|boli|capuchon de lapicero|capuchon de plumon|plumilla|borrador,
Globo,general,globo|globos,
Vajilla o vidrio roto,general,vaso roto|plato roto|vidrio roto|copa rota|taza rota|ceramica rota|vajilla dañada|pizarra rota,"Por seguridad, la cerámica y el vidrio roto van a generales, nunca a reciclables. Envuélvelo antes de botarlo."
Cápsula de café,general,capsula de cafe|capsula nespresso|capsula dolce gusto,"Sea de aluminio o de plástico, no se recicla como material simple."
Mascarilla o guantes descartables,general,mascarilla|tapabocas|cubrebocas|guantes|guantes de latex|guante descartable,"Por higiene, siempre van a generales."
Colillas de cigarro,general,colilla|cigarro|ceniza,
Goma de mascar o chicle,general,chicle|goma de mascar|chicle usado,
Mica o funda plástica,general,mica|micas|funda plastica|porta papeles plastico|folder plastico|celofan,
Grapa o gancho de grapadora,general,grapa|grapas|gancho de grapadora|grampa|corchete|grapas usadas,
Imperdible o alfiler,general,imperdible|imperdibles|alfiler|alfileres,
Imán,general,iman|imanes,
Tubo de pasta de dientes,general,pasta de dientes|pasta dental|tubo de pasta,Combina varias capas de material que no se pueden separar.
Cepillo de dientes,general,cepillo de dientes|cepillo dental,
Botón,general,boton|botones,
CD o DVD,general,cd|dvd|disco compacto|memoria sd|micro sd,Combina plástico y una capa metálica que no se puede separar.
"Envase con atomizador, gel o desinfectante",general,atomizador|rociador|difusor ambiental|dispensador vacio|desinfectante|gel antibacterial|desodorante,El mecanismo de bomba mezcla varios materiales que no se pueden separar.
Toallita húmeda o paño desechable,general,toalla humeda|toallita desinfectante|paño desechable|paño de microfibra,
"Hisopo, esponja o curita",general,hisopo|esponja|curita|curitas|banda adhesiva|algodon|hilo dental,
Confeti,general,confeti,
Vela,general,vela|velas,
Tarjeta o credencial plástica,general,tarjeta de presentacion|tarjeta de regalo|tarjeta magnetica|tarjeta vencida|identificador|escarapela|cordon porta carnet|lanyard|colgador de credencial|porta credencial|pulsera de evento|vincha|billetera de papel,
"Cúter, cutter o abrecartas",general,cutter|cuter|abrecartas|repuesto de cutter,Es un objeto cortante. Envuélvelo antes de botarlo para evitar accidentes.
"Candado, llave o llavero roto",general,candado|llave|llavero roto|correa de reloj|cordon de zapato|percha,
Clip o sujetapapeles,general,clip|clips|clip mariposa|sujetapapeles|sujetador metalico|pinza sujetapapeles|pinza|gancho binder|gancho negro|gancho mariposa|broche,
Cinta adhesiva o precinto,general,cinta adhesiva|cinta de embalaje|cinta doble cara|cinta masking|cinta transparente|precinto|brida plastica|liga|elastico|liga de pelo|liguita|banda elastica,
"Goma, silicona o pegamento",general,pegamento|goma en barra|goma liquida|silicona|silicona liquida|silicona en barra|sellador,
Carpeta o archivador de plástico,general,carpeta|folder|archivador|divisor de carpetas|dossier|folder colgante|porta documentos|anillo metalico|aro de carpeta|bandeja organizadora|porta lapiceros roto|porta notas|porta clips|faja archivadora,
"Etiqueta, sticker o pegatina",general,etiqueta|etiqueta adhesiva|etiqueta de envio|etiqueta de precio|etiqueta termica|sticker|pegatina|vinil adhesivo|papel de sticker|rollo de etiquetas|rollo de stickers,
Regla o percha,general,regla,
Papel de regalo,general,papel de regalo|mantel de papel,
Papel aluminio o papel plata,general,papel aluminio|papel plata|foil|aluminio de cocina,"Si tiene restos de comida va a generales. Si está limpio, puedes juntarlo con las latas."
Táper de delivery con comida y envoltorios,reciclable,taper de delivery|delivery|caja de delivery con comida|bolsa de delivery|caja de pizza|caja de hamburguesa|caja de sushi,"Sepáralo en 3 pasos: 1) las sobras de comida van al balde de orgánicos, 2) las servilletas, cubiertos rotos y sachets van a generales, 3) el táper se enjuaga y va a reciclables."
Pilas y baterías,especial,pila|bateria|pilas alcalinas|bateria de laptop|pilas aa|pilas aaa,"No van en estos tachos. Llévalas al punto de acopio de RAEE (Residuos de Aparatos Eléctricos y Electrónicos) del área de TIC."
Aparatos electrónicos,especial,electronico|cargador|cable electronico|audifonos|mouse|teclado|celular malogrado|laptop malograda|raee|cable usb|cable usb-c|cable lightning|cable hdmi|cable vga|cable ethernet|cable de poder|cable de red|adaptador usb|adaptador hdmi|conector usb|memoria usb dañada|bateria externa|carga portatil|mouse dañado|teclado dañado|audifonos dañados|audifonos descartables|microfono roto|calculadora dañada|temporizador|soporte para celular|soporte para laptop|protector de pantalla|protector de cables|limpiador de pantalla|puntero laser,"Son RAEE (Residuos de Aparatos Eléctricos y Electrónicos). Llévalos al punto de acopio del área de TIC, no van con la recolección normal."
Tóner o cartucho de impresora,especial,toner|cartucho|tinta de impresora|cartucho de tinta|recarga de tinta|jeringa de tinta|cartucho de toner,Muchas marcas tienen programa de retorno de cartuchos.
Focos y lámparas,especial,foco|fluorescente|lampara|led fundido|bombilla led,Contienen materiales que requieren manejo especial.
Medicamentos vencidos,especial,medicina|pastillas|medicamento vencido|blister de medicamentos,Llévalos a un punto de acopio de farmacia.`;

const BIN_LABELS = {
  organico: 'Orgánicos',
  general: 'Generales',
  reciclable: 'Reciclables',
  especial: 'Punto especial',
};

const TIPS = [
  'Reciclar es una decisión.',
  'Si no sabes dónde va, bótalo en generales.',
  'Una botella de plástico puede tardar hasta 450 años en degradarse.',
  'El aluminio se recicla infinitas veces sin perder calidad.',
  'El vidrio es 100% reciclable, sin límite de veces.',
  'El papel puede reciclarse hasta 7 veces antes de perder su fibra.',
  'Una pila mal botada puede contaminar miles de litros de agua.',
  'Separar bien es un gesto pequeño con impacto grande.',
  'Lo que separas bien hoy, mañana se convierte en algo nuevo.',
  'Tu esfuerzo de hoy es el planeta de mañana.',
];

function randomTip() {
  return TIPS[Math.floor(Math.random() * TIPS.length)];
}

let ITEMS = [];

const searchInput = document.getElementById('search-input');
const searchForm = document.getElementById('search-form');
const resultArea = document.getElementById('result-area');
const idleHint = document.getElementById('idle-hint');
const idleTip = document.getElementById('idle-tip');

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

  if (idleTip) idleTip.textContent = randomTip();

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

  html += `<p class="tip-banner">${escapeHTML(randomTip())}</p>`;

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
      <p class="no-result__fallback">Si no sabes dónde va, ponlo en generales.</p>
      <ul>
        <li><strong>Orgánicos:</strong> restos de comida, cáscaras, posos de café.</li>
        <li><strong>Generales:</strong> todo lo sucio, mixto o descartable (vasos de café, tecnopor, mascarillas).</li>
        <li><strong>Reciclables:</strong> botellas, latas, papel y cartón limpios y secos.</li>
      </ul>
      <a href="mailto:micaela.paez.a@astara.com?cc=alejandra.arroyo.d@astara.com&subject=Residuo%20no%20encontrado&body=No%20encontr%C3%A9%3A%20${encodeURIComponent(query)}">Escríbenos sobre "${escapeHTML(query)}"</a>
    </div>
    <p class="tip-banner">${escapeHTML(randomTip())}</p>
  `;
}

function escapeHTML(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
