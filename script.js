/* ---------------------------------------------------------------------
   MODO CLARO / ESCURO
--------------------------------------------------------------------- */
(function initTheme(){
  const root = document.documentElement;
  const btn = document.getElementById('theme-toggle');
  const icon = document.getElementById('theme-icon');
  const label = document.getElementById('theme-label');

  function apply(theme){
    root.setAttribute('data-theme', theme);
    icon.textContent = theme === 'dark' ? '☀' : '☾';
    label.textContent = theme === 'dark' ? 'Modo claro' : 'Modo escuro';
    try{ localStorage.setItem('composicao-preco-theme', theme); }catch(e){}
  }

  let saved = 'light';
  try{ saved = localStorage.getItem('composicao-preco-theme') || 'light'; }catch(e){}
  apply(saved);

  btn.addEventListener('click', () => {
    apply(root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
  });
})();

/* ---------------------------------------------------------------------
   HELPERS
--------------------------------------------------------------------- */
function fmtBRL(n){ return (n || 0).toLocaleString('pt-BR', {style:'currency', currency:'BRL'}); }
function fmtPct(n){ return (n || 0).toLocaleString('pt-BR', {maximumFractionDigits: 2}) + '%'; }
function escapeHtml(str){
  return String(str == null ? '' : str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

/* ---------------------------------------------------------------------
   LEITURA DA PLANILHA ("Ver. 3.0 - Planilha de Cálculo de Revenda ou
   Distribuição"). O layout é fixo; as colunas abaixo seguem esse modelo.
   O site NÃO recalcula: lê os valores já calculados (em cache) pelo Excel.
--------------------------------------------------------------------- */
const SHEET_NAME = 'Revenda';
const MAX_HEADER_SCAN = 60;  // quantas linhas do topo são examinadas para achar o cabeçalho
const MAX_ITEMS = 100;       // itens lidos abaixo do cabeçalho
const COL = {
  item: 'A',      // código do item
  desc: 'E',      // descrição
  ncm: 'G',       // NCM
  qtd: 'P',       // quantidade
  icms: 'R',      // ICMS de saída (%)
  pisCofins: 'S', // PIS/COFINS de saída (%)
  ipi: 'T',       // IPI de saída (%)
  unitDireta: 'W',  // preço unitário - revenda direta (receita bruta s/ IPI)
  totalDireta: 'X', // preço total - revenda direta
  unitPortal: 'Z',  // preço unitário - portal (líquido de ICMS e IPI)
  totalPortal: 'AA' // preço total - portal
};

let items = [];

/* células com erro (#REF!, #VALUE!...) são tratadas como vazias */
function cellAt(ws, addr){
  const c = ws[addr];
  return (c && c.t !== 'e') ? c : null;
}
function numAt(ws, addr){
  const c = cellAt(ws, addr);
  return (c && typeof c.v === 'number' && isFinite(c.v)) ? c.v : null;
}
function strAt(ws, addr){
  const c = cellAt(ws, addr);
  return (c && c.v != null) ? String(c.v).trim() : '';
}

/* procura uma coluna pelo nome do cabeçalho (nas duas linhas do cabeçalho) */
function findColumnByHeader(ws, headerRow, re){
  for(const row of [headerRow - 1, headerRow]){
    for(let c = 0; c < 60; c++){
      const addr = XLSX.utils.encode_cell({r: row - 1, c: c});
      if(re.test(normalizeName(strAt(ws, addr)))) return XLSX.utils.encode_col(c);
    }
  }
  return null;
}

/* acha a aba "Revenda" pelo nome (ignora maiúsculas, acentos e espaços nas pontas) */
function normalizeName(n){
  return String(n).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
}
function findSheetName(names){
  const alvo = normalizeName(SHEET_NAME);
  return names.find(n => normalizeName(n) === alvo) || null;
}

/* acha a linha do cabeçalho: coluna A = "ITEM" e coluna X = "P. TOTAL" */
function findHeaderRow(ws){
  const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
  const ultima = ref ? Math.min(ref.e.r + 1, MAX_HEADER_SCAN) : MAX_HEADER_SCAN;
  let soItem = null;
  for(let r = 1; r <= ultima; r++){
    if(normalizeName(strAt(ws, COL.item + r)) !== 'item') continue;
    if(/total/i.test(strAt(ws, COL.totalDireta + r))) return {row: r, exato: true};
    if(soItem === null) soItem = r;
  }
  return soItem !== null ? {row: soItem, exato: false} : null;
}

/* procura um rótulo acima do cabeçalho e devolve o valor da célula ao lado */
function findLabeled(ws, re, ateLinha, numerico){
  for(let r = 1; r <= ateLinha; r++){
    for(let c = 0; c < 14; c++){
      const t = strAt(ws, XLSX.utils.encode_cell({r: r - 1, c: c}));
      if(!t || !re.test(normalizeName(t))) continue;
      const alcance = numerico ? 4 : 1;
      for(let k = c + 1; k <= c + alcance; k++){
        const addr = XLSX.utils.encode_cell({r: r - 1, c: k});
        if(numerico){ const n = numAt(ws, addr); if(n !== null) return n; }
        else { const v = strAt(ws, addr); if(v) return v; }
      }
    }
  }
  return null;
}

function parseWorkbook(wb, sheetName){
  const warnings = [];
  const ws = wb.Sheets[sheetName];

  const cab = findHeaderRow(ws);
  if(!cab){
    return {list: [], warnings: [], proposta: '', cliente: '', headerRow: null};
  }
  const headerRow = cab.row;
  if(!cab.exato) warnings.push('Achei "ITEM" na linha ' + headerRow + ', mas a coluna ' + COL.totalDireta + ' não tem "P. TOTAL". Confira se as colunas leem os valores certos.');

  const firstRow = headerRow + 1;
  let lastRow = headerRow + MAX_ITEMS;
  if(ws['!ref']) lastRow = Math.min(lastRow, XLSX.utils.decode_range(ws['!ref']).e.r + 1);

  const descCol = findColumnByHeader(ws, headerRow, /^descricao$/);
  if(!descCol) warnings.push('Não encontrei a coluna "DESCRIÇÃO" no cabeçalho; usei a coluna ' + COL.desc + '.');
  const nomeCol = descCol || COL.desc;
  const list = [];
  const skipped = [];

  for(let r = firstRow; r <= lastRow; r++){
    const code = strAt(ws, COL.item + r);
    if(code === '') continue;

    const qtd = numAt(ws, COL.qtd + r);
    const prices = {
      unitDireta: numAt(ws, COL.unitDireta + r),
      totalDireta: numAt(ws, COL.totalDireta + r),
      unitPortal: numAt(ws, COL.unitPortal + r),
      totalPortal: numAt(ws, COL.totalPortal + r)
    };
    if(!(qtd > 0)){ skipped.push('Item ' + code + ' (linha ' + r + '): quantidade vazia.'); continue; }
    if(Object.values(prices).some(v => v === null)){
      skipped.push('Item ' + code + ' (linha ' + r + '): preço sem valor calculado (confira se a planilha tem erro, como #REF!).');
      continue;
    }

    list.push({
      code: code,
      desc: strAt(ws, nomeCol + r),
      ncm: strAt(ws, COL.ncm + r),
      qtd: qtd,
      icms: (numAt(ws, COL.icms + r) || 0) * 100,
      pisCofins: (numAt(ws, COL.pisCofins + r) || 0) * 100,
      ipi: (numAt(ws, COL.ipi + r) || 0) * 100,
      unitDireta: prices.unitDireta,
      totalDireta: prices.totalDireta,
      unitPortal: prices.unitPortal,
      totalPortal: prices.totalPortal
    });
  }

  skipped.forEach(s => warnings.push(s));

  /* conferência: soma lida x subtotal da própria planilha (primeiro número da coluna X acima do cabeçalho) */
  let subtotal = null;
  for(let r = headerRow - 1; r >= Math.max(1, headerRow - 4) && subtotal === null; r--){
    subtotal = numAt(ws, COL.totalDireta + r);
  }
  if(skipped.length === 0 && subtotal !== null){
    const soma = list.reduce((s, it) => s + it.totalDireta, 0);
    if(Math.abs(soma - subtotal) > 0.01) warnings.push('A soma dos itens (' + fmtBRL(soma) + ') difere do subtotal da planilha (' + fmtBRL(subtotal) + ').');
  }

  /* proposta e cliente: achados pelo rótulo, acima do cabeçalho */
  return {
    list: list,
    warnings: warnings,
    proposta: findLabeled(ws, /^proposta\b/, headerRow - 1, false) || '',
    cliente: findLabeled(ws, /^cliente\b/, headerRow - 1, false) || '',
    descCol: nomeCol,
    headerRow: headerRow
  };
}

document.getElementById('xlsx-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const statusEl = document.getElementById('upload-status');
  const warnEl = document.getElementById('upload-warnings');
  if(!file) return;
  statusEl.className = '';
  statusEl.textContent = 'Lendo planilha...';
  warnEl.innerHTML = '';

  if(typeof XLSX === 'undefined'){
    statusEl.className = 'err';
    statusEl.textContent = 'A biblioteca de leitura de Excel não carregou (verifique a conexão ou se algo bloqueia cdnjs.cloudflare.com).';
    return;
  }

  try{
    const buf = await file.arrayBuffer();

    /* 1ª leitura: só os nomes das abas (rápida, mesmo com abas grandes) */
    const nomes = XLSX.read(buf, {type: 'array', bookSheets: true}).SheetNames;
    const alvo = findSheetName(nomes);
    if(!alvo){
      items = [];
      statusEl.className = 'err';
      statusEl.textContent = 'Não encontrei a aba "' + SHEET_NAME + '" neste arquivo. Abas encontradas: ' + nomes.join(', ') + '.';
      renderOutput();
      return;
    }

    /* 2ª leitura: só a aba Revenda */
    const wb = XLSX.read(buf, {type: 'array', sheets: alvo});
    const res = parseWorkbook(wb, alvo);
    if(res.headerRow === null){
      items = [];
      statusEl.className = 'err';
      statusEl.textContent = 'Não encontrei o cabeçalho na aba "' + alvo + '" (procurei "ITEM" na coluna ' + COL.item + ' e "P. TOTAL" na coluna ' + COL.totalDireta + ' nas primeiras ' + MAX_HEADER_SCAN + ' linhas).';
      renderOutput();
      return;
    }
    items = res.list;

    const partes = ['aba "' + alvo + '"', 'cabeçalho na linha ' + res.headerRow, items.length + (items.length === 1 ? ' item lido' : ' itens lidos')];
    if(res.proposta) partes.push('proposta ' + res.proposta);
    if(res.cliente) partes.push('cliente ' + res.cliente);
    statusEl.className = items.length ? 'ok' : 'err';
    const resumo = partes.join(' · ');
    statusEl.textContent = items.length ? 'Planilha lida: ' + resumo + (resumo.endsWith('.') ? '' : '.') : 'Nenhum item válido encontrado na planilha.';
    warnEl.innerHTML = res.warnings.map(w => '<li>' + escapeHtml(w) + '</li>').join('');
    renderOutput();
  } catch(err){
    statusEl.className = 'err';
    statusEl.textContent = 'Erro ao ler a planilha: ' + err.message;
  }
});

/* ---------------------------------------------------------------------
   TABELA DE SAÍDA — Item | Modelo | Qtd | NCM | Preço Unitário | Preço Total
--------------------------------------------------------------------- */
function groupByRate(list, key){
  const map = new Map();
  list.forEach((it, idx) => {
    const v = Math.round(it[key] * 10000) / 10000;
    if(!map.has(v)) map.set(v, []);
    map.get(v).push(idx + 1);
  });
  return Array.from(map.entries()).map(([rate, nums]) => ({rate, nums}));
}
function describeRates(list, key){
  const groups = groupByRate(list, key);
  if(groups.length === 1) return fmtPct(groups[0].rate);
  return groups.map(g => fmtPct(g.rate) + ' (' + (g.nums.length > 1 ? 'itens ' : 'item ') + g.nums.join(', ') + ')').join(' e ');
}

function renderOutput(){
  const portal = document.getElementById('opt-preco').value === 'portal';
  const tbody = document.getElementById('output-tbody');
  let total = 0;

  const rowsHtml = items.map((it, idx) => {
    const unit = portal ? it.unitPortal : it.unitDireta;
    const tot = portal ? it.totalPortal : it.totalDireta;
    total += tot;
    const nome = it.desc || ('Item ' + it.code);
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td class="left"><strong>' + escapeHtml(nome) + '</strong></td>' +
      '<td>' + it.qtd + '</td>' +
      '<td>' + escapeHtml(it.ncm || '—') + '</td>' +
      '<td>' + fmtBRL(unit) + '</td>' +
      '<td>' + fmtBRL(tot) + '</td>' +
    '</tr>';
  }).join('');

  tbody.innerHTML = rowsHtml || '<tr><td colspan="6">Escolha a planilha para carregar os itens.</td></tr>';
  document.getElementById('output-total').textContent = fmtBRL(total);

  renderOutputNotes(portal);
}

/* bullets de impostos (alíquotas de saída lidas da planilha), estado e Zona Franca */
function renderOutputNotes(portal){
  const ncms = [];
  items.forEach(it => {
    const n = (it.ncm || '').trim();
    if(n && ncms.indexOf(n) === -1) ncms.push(n);
  });

  const ipiMode = portal ? 'excluso' : document.getElementById('opt-ipi').value;
  const taxLines = [];
  if(items.length){
    taxLines.push('PIS/COFINS: ' + describeRates(items, 'pisCofins') + ' incluso');
    taxLines.push('ICMS: ' + describeRates(items, 'icms') + (portal ? ' excluso (preços líquidos de ICMS)' : ' incluso'));
    taxLines.push('IPI: ' + describeRates(items, 'ipi') + ' ' + ipiMode);
  }
  const taxHtml = taxLines.map(l => '<li>' + escapeHtml(l) + '</li>').join('');

  document.getElementById('output-tax-list').innerHTML = ncms.length
    ? '<li>NCM: ' + escapeHtml(ncms.join(', ')) + '<ul>' + taxHtml + '</ul></li>'
    : taxHtml;

  const estado = document.getElementById('opt-estado').value.trim();
  const estadoItem = document.getElementById('output-estado-item');
  if(estado){
    estadoItem.textContent = 'Estamos considerando que o cliente é contribuinte de ICMS no estado do ' + estado + ';';
    estadoItem.classList.remove('hidden');
  } else {
    estadoItem.classList.add('hidden');
  }

  document.getElementById('output-zf').classList.toggle('hidden', !document.getElementById('opt-zf').checked);
}

['opt-preco', 'opt-estado', 'opt-ipi', 'opt-zf'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener('input', renderOutput);
  el.addEventListener('change', renderOutput);
});

/* ---------------------------------------------------------------------
   COPIAR — seleciona o bloco e copia mantendo a formatação (cola no Word)
--------------------------------------------------------------------- */
function copyBlock(wrapId, statusId){
  const statusEl = document.getElementById(statusId);
  const wrap = document.getElementById(wrapId);
  try{
    const range = document.createRange();
    range.selectNodeContents(wrap);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand('copy');
    sel.removeAllRanges();
    statusEl.textContent = ok ? 'Copiado! Cole com Ctrl+V no Word.' : 'Não copiou automaticamente — selecione o conteúdo manualmente e use Ctrl+C.';
  } catch(err){
    statusEl.textContent = 'Não copiou automaticamente — selecione o conteúdo manualmente e use Ctrl+C.';
  }
  setTimeout(() => { statusEl.textContent = ''; }, 4000);
}
document.getElementById('copy-table').addEventListener('click', () => copyBlock('output-wrap', 'copy-status'));
document.getElementById('copy-tech').addEventListener('click', () => copyBlock('tech-wrap', 'copy-tech-status'));

/* ---------------------------------------------------------------------
   PROPOSTA TÉCNICA — leitura do PDF da cotação do fornecedor
   Cada item começa em "P/N: <código> – (<descrição>)" e é seguido por
   "<qtd> <preço unitário> <valor>". O pdf.js só é baixado quando um PDF
   é escolhido (não pesa na abertura do site).
--------------------------------------------------------------------- */
const PDFJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
const PDFJS_WORKER = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
/* cabeçalhos e rodapés repetidos em cada página, removidos do texto lido */
const PDF_NOISE = [
  /Keeping the World Flowing for Future Generations\s*\d+\s*(?:of|de)\s*\d+/gi,
  /Rotork Quotation\s*-\s*\S+/gi
];

let techItems = [];

function loadPdfJs(){
  return new Promise((resolve, reject) => {
    const pronto = () => { pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; resolve(); };
    if(typeof pdfjsLib !== 'undefined'){ pronto(); return; }
    const sc = document.createElement('script');
    sc.src = PDFJS_URL;
    sc.onload = pronto;
    sc.onerror = () => reject(new Error('A biblioteca de leitura de PDF não carregou (verifique a conexão ou se algo bloqueia cdnjs.cloudflare.com).'));
    document.head.appendChild(sc);
  });
}

/* monta o texto da página: fim de linha vira espaço, exceto após hífen
   (ex.: "XW-00-" + "01580-..." continua sendo um único código) */
function pageText(items){
  let t = '';
  items.forEach(it => { t += it.str; if(it.hasEOL) t += '\n'; });
  t = t.replace(/-\n(?=\S)/g, '-').replace(/\s*\n\s*/g, ' ');
  PDF_NOISE.forEach(re => { t = t.replace(re, ' '); });
  return t.replace(/\s{2,}/g, ' ').trim();
}

/* "1.00" -> 1 ; "1,5" -> 1.5 ; "1.250,00" -> 1250 */
function parseQty(str){
  let v = String(str).trim();
  if(v.indexOf(',') !== -1 && v.indexOf('.') !== -1){
    v = v.lastIndexOf(',') > v.lastIndexOf('.') ? v.replace(/\./g, '').replace(',', '.') : v.replace(/,/g, '');
  } else if(v.indexOf(',') !== -1){
    v = /,\d{1,2}$/.test(v) ? v.replace(',', '.') : v.replace(/,/g, '');
  }
  const n = parseFloat(v);
  return isFinite(n) ? n : null;
}

/* separa a descrição em tópicos "Rótulo: valor" (vírgula seguida de um rótulo com dois-pontos) */
function splitSpecs(desc){
  if(!desc) return [];
  return desc
    .split(/,\s+(?=[A-Z][^,:;()]{0,50}:\s)/)
    .map(x => x.trim().replace(/[,;]+$/, '').trim())
    .filter(Boolean);
}

function parseTechItems(text){
  const re = /P\/N:\s*/g;
  const marks = [];
  let m;
  while((m = re.exec(text)) !== null) marks.push({start: m.index, after: m.index + m[0].length});

  const out = [];
  marks.forEach((mk, i) => {
    const seg = text.slice(mk.after, i + 1 < marks.length ? marks[i + 1].start : text.length);
    const pnMatch = seg.match(/^(\S+)/);
    if(!pnMatch) return;
    const pn = pnMatch[1].replace(/[–—-]+$/, '');
    const rest = seg.slice(pnMatch[0].length);

    let desc = '';
    let tail = rest;
    const open = rest.indexOf('(');
    if(open !== -1){
      let depth = 0, end = -1;
      for(let k = open; k < rest.length; k++){
        if(rest[k] === '(') depth++;
        else if(rest[k] === ')'){ depth--; if(depth === 0){ end = k; break; } }
      }
      if(end === -1) end = rest.length;
      desc = rest.slice(open + 1, end).trim();
      tail = rest.slice(end + 1);
    }

    const nums = tail.match(/(\d[\d.,]*)\s+(\d[\d.,]*)\s+(\d[\d.,]*)/);
    if(open === -1){
      const cut = nums ? rest.indexOf(nums[0]) : rest.length;
      desc = rest.slice(0, cut).replace(/^\s*[–—-]\s*/, '').trim();
    }

    out.push({pn: pn, desc: desc, bullets: splitSpecs(desc), qtd: nums ? parseQty(nums[1]) : null});
  });
  return out;
}

function renderTech(){
  const tbody = document.getElementById('tech-tbody');
  tbody.innerHTML = techItems.map((it, idx) => {
    const lis = it.bullets.map(b => '<li>' + escapeHtml(b) + '</li>').join('');
    const qtd = it.qtd === null ? '—' : it.qtd.toLocaleString('pt-BR', {maximumFractionDigits: 2});
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td class="left"><strong>' + escapeHtml(it.pn) + '</strong>' + (lis ? '<ul>' + lis + '</ul>' : '') + '</td>' +
      '<td class="right">' + qtd + '</td>' +
    '</tr>';
  }).join('') || '<tr><td colspan="3">Escolha o PDF da cotação para carregar os itens.</td></tr>';
}

document.getElementById('pdf-input').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const statusEl = document.getElementById('pdf-status');
  const warnEl = document.getElementById('pdf-warnings');
  if(!file) return;
  statusEl.className = '';
  statusEl.textContent = 'Lendo PDF...';
  warnEl.innerHTML = '';

  try{
    await loadPdfJs();
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: new Uint8Array(buf)}).promise;
    const textos = [];
    for(let i = 1; i <= pdf.numPages; i++){
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      textos.push(pageText(content.items));
    }
    await pdf.destroy();

    const texto = textos.join(' ');
    techItems = parseTechItems(texto);

    const warnings = [];
    techItems.forEach((it, idx) => {
      if(it.qtd === null) warnings.push('Item ' + (idx + 1) + ' (' + it.pn + '): quantidade não encontrada.');
      if(!it.desc) warnings.push('Item ' + (idx + 1) + ' (' + it.pn + '): descrição não encontrada.');
    });

    const ref = texto.match(/Nossa\s+refer[eê]ncia:\s*(\S+)/i);
    if(techItems.length){
      statusEl.className = 'ok';
      statusEl.textContent = 'PDF lido: ' + techItems.length + (techItems.length === 1 ? ' item' : ' itens') + (ref ? ' · cotação ' + ref[1] : '') + '.';
    } else {
      statusEl.className = 'err';
      statusEl.textContent = 'Não encontrei itens neste PDF (procurei linhas começando com "P/N:").';
    }
    warnEl.innerHTML = warnings.map(w => '<li>' + escapeHtml(w) + '</li>').join('');
    renderTech();
  } catch(err){
    statusEl.className = 'err';
    statusEl.textContent = 'Erro ao ler o PDF: ' + err.message;
  }
});

renderOutput();
renderTech();
