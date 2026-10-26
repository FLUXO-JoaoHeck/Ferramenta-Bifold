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
const HEADER_ROW = 13;
const FIRST_ROW = 14;
const LAST_ROW = 113;
const COL = {
  item: 'A',
  desc: 'E',
  ncm: 'G',
  qtd: 'P',
  icms: 'R',
  pisCofins: 'S',
  ipi: 'T',
  unitDireta: 'W',
  totalDireta: 'X',
  unitPortal: 'Z',
  totalPortal: 'AA'
};
const CELL_SERVICOS = 'E9';

let items = [];
let servicos = null;

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

/* procura uma coluna de PN nas linhas de cabeçalho (12 e 13) */
function findPnColumn(ws){
  const re = /^(pn|p\/n|part\s*number|partnumber|n[ºo°]?\s*da\s*pe[çc]a|n[úu]mero\s*da\s*pe[çc]a)$/i;
  for(const row of [HEADER_ROW - 1, HEADER_ROW]){
    for(let c = 0; c < 60; c++){
      const addr = XLSX.utils.encode_cell({r: row - 1, c: c});
      if(re.test(strAt(ws, addr))) return XLSX.utils.encode_col(c);
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

function parseWorkbook(wb, sheetName){
  const warnings = [];
  const ws = wb.Sheets[sheetName];

  const okLayout = strAt(ws, COL.item + HEADER_ROW).toUpperCase() === 'ITEM'
    && /TOTAL/i.test(strAt(ws, COL.totalDireta + HEADER_ROW));
  if(!okLayout) warnings.push('O cabeçalho da linha ' + HEADER_ROW + ' não parece o da planilha Ver. 3.0. Confira se as colunas leem os valores certos.');

  const pnCol = findPnColumn(ws);
  const list = [];
  const skipped = [];

  for(let r = FIRST_ROW; r <= LAST_ROW; r++){
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
      pn: pnCol ? strAt(ws, pnCol + r) : '',
      desc: strAt(ws, COL.desc + r),
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

  if(!pnCol) warnings.push('Não encontrei uma coluna "PN" no cabeçalho; a descrição será usada como identificação do item.');
  skipped.forEach(s => warnings.push(s));

  /* conferência: soma lida x subtotal da própria planilha (X11) */
  const subtotal = numAt(ws, COL.totalDireta + '11');
  if(skipped.length === 0 && subtotal !== null){
    const soma = list.reduce((s, it) => s + it.totalDireta, 0);
    if(Math.abs(soma - subtotal) > 0.01) warnings.push('A soma dos itens (' + fmtBRL(soma) + ') difere do subtotal da planilha (' + fmtBRL(subtotal) + ').');
  }

  return {
    list: list,
    warnings: warnings,
    servicos: numAt(ws, CELL_SERVICOS),
    proposta: strAt(ws, 'E2'),
    cliente: strAt(ws, 'E3'),
    pnCol: pnCol
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

    /* 1ª leitura: só os nomes das abas */
    const nomes = XLSX.read(buf, {type: 'array', bookSheets: true}).SheetNames;
    const alvo = findSheetName(nomes);
    if(!alvo){
      items = [];
      servicos = null;
      statusEl.className = 'err';
      statusEl.textContent = 'Não encontrei a aba "' + SHEET_NAME + '" neste arquivo. Abas encontradas: ' + nomes.join(', ') + '.';
      renderOutput();
      return;
    }

    /* 2ª leitura: só a aba Revenda */
    const wb = XLSX.read(buf, {type: 'array', sheets: alvo});
    const res = parseWorkbook(wb, alvo);
    items = res.list;
    servicos = res.servicos;

    const partes = ['aba "' + alvo + '"', items.length + (items.length === 1 ? ' item lido' : ' itens lidos')];
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
   TABELA DE SAÍDA
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
    const nome = it.pn || it.desc || ('Item ' + it.code);
    const sub = it.pn ? it.desc : '';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td class="left"><strong>' + escapeHtml(nome) + '</strong>' + (sub ? '<br>' + escapeHtml(sub) : '') + '</td>' +
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

/* impostos, estado e Zona Franca */
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
    estadoItem.textContent = 'Consideramos a compradora contribuinte de ICMS no estado de ' + estado + ';';
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
   COPIAR TABELA — mantém a formatação ao colar no Word
--------------------------------------------------------------------- */
document.getElementById('copy-table').addEventListener('click', () => {
  const statusEl = document.getElementById('copy-status');
  const wrap = document.getElementById('output-wrap');
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
});

renderOutput();

/* ---------------------------------------------------------------------
   EXTRAÇÃO LOCAL DE PDF
   O arquivo é lido no navegador com PDF.js; não é enviado a um servidor.
   PDFs digitalizados (somente imagem) exigem OCR, que não está incluído.
--------------------------------------------------------------------- */
(function initPdfExtractor(){
  const fileInput = document.getElementById('pdf-input');
  if(!fileInput) return;
  const extractBtn = document.getElementById('extract-pdf');
  const status = document.getElementById('pdf-status');
  const editor = document.getElementById('pdf-editor');
  const itemsHost = document.getElementById('pdf-items');
  const formatted = document.getElementById('pdf-formatted-text');
  const rawText = document.getElementById('pdf-raw-text');
  let extractedItems = [];

  if(window.pdfjsLib){
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }

  function clean(s){ return String(s || '').replace(/\s+/g, ' ').trim(); }
  function normalizeLine(s){ return clean(s).replace(/\s+([,.;:])/g, '$1'); }
  function parseMoney(s){ return String(s || '').replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, ''); }
  function field(id){ return document.getElementById(id).value.trim(); }

  async function extractLines(file){
    if(!window.pdfjsLib) throw new Error('A biblioteca PDF.js não carregou. Verifique a conexão com a internet e tente novamente.');
    const buffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: buffer}).promise;
    const lines = [];

    for(let pageNum = 1; pageNum <= pdf.numPages; pageNum++){
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      const grouped = [];

      content.items.forEach(item => {
        if(!item.str || !item.str.trim()) return;
        const x = item.transform ? item.transform[4] : 0;
        const y = item.transform ? item.transform[5] : 0;
        let line = grouped.find(g => Math.abs(g.y - y) < 2.2);
        if(!line){ line = {y:y, parts:[]}; grouped.push(line); }
        line.parts.push({x:x, text:item.str});
      });

      grouped.sort((a,b) => b.y-a.y);
      grouped.forEach(line => {
        line.parts.sort((a,b) => a.x-b.x);
        const text = normalizeLine(line.parts.map(p => p.text).join(' '));
        if(text) lines.push(text);
      });
      lines.push('');
    }
    return lines;
  }
  function findMatch(lines, regex){
    const line = lines.find(l => regex.test(l));
    if(!line) return '';
    const match = line.match(regex);
    return match && match[1] ? clean(match[1]) : '';
  }

  
function parsePdfText(lines) {
  const all = lines.join('\n');

  // Dados gerais da proposta
  const proposalMatch =
    all.match(/PROPOSTA\s+TÉCNICA\s+([A-Z0-9-]+(?:\s+REV\d+)?(?:\s*-\s*[^\n]+)?)/i) ||
    all.match(/Rotork\s+Quotation\s*[-–]\s*([^\s\n]+)/i) ||
    all.match(/Nossa refer[êe]ncia:\s*([^\n]+)/i);

  const proposal = proposalMatch ? clean(proposalMatch[1]) : '';

  const dateMatch = all.match(/\b(\d{2}\/\d{2}\/\d{4})\b/);
  const date = dateMatch ? dateMatch[1] : '';

  const clientMatch =
    all.match(/CLIENTE\s*\n\s*([^\n]+)/i) ||
    all.match(/Cliente:\s*([^\n]+)/i);

  const client = clientMatch ? clean(clientMatch[1]) : '';

  const validity = findMatch(
    lines,
    /Expira[çc][ãa]o da cota[çc][ãa]o:\s*(.+)$/i
  );

  const payment =
    lines.find(l => /^(100%|\d+%).*DDL.*$/i.test(l)) || '';

  const delivery = findMatch(
    lines,
    /^(\d+\s*dias[^\n]*)$/i
  );

  const ncm = findMatch(
    lines,
    /Classifica[çc][ãa]o fiscal:\s*(.+)$/i
  );

  const taxLines = lines.filter(l =>
    /^(?:Classifica[çc][ãa]o fiscal:|[-–]\s*(?:PIS|COFINS|ICMS|IPI):|Estamos considerando)/i.test(l)
  );

  // Captura o total indicado no documento
  let grandTotal = '';
  const totalMatches = all.match(
    /(?:Total Geral|Valor Total|Total da Proposta)[^\n]*?(?:BRL|R\$)?\s*([\d.,]+)/gi
  );

  if (totalMatches && totalMatches.length) {
    const last = totalMatches[totalMatches.length - 1];
    const m = last.match(/([\d.]+,\d{2}|\d{1,3}(?:,\d{3})*\.\d{2})\s*$/);
    if (m) grandTotal = m[1];
  }

  // Identificação dos itens
  const parsedItems = [];
  let current = null;

  function finishItem() {
    if (current && current.pn) {
      current.description = clean(current.description);
      current.specs = current.specs
        .map(s => clean(s.replace(/^[•●▪\-–]\s*/, '')))
        .filter(Boolean);

      // Mantém a descrição separada das especificações
      current.description += current.specs.length
        ? '\n' + current.specs.map(s => '• ' + s).join('\n')
        : '';

      parsedItems.push(current);
    }
    current = null;
  }

  const pnLine = /^P\s*\/\s*N\s*:\s*(.+)$/i;

  // Exemplo de linha de tabela:
  // 1 VRP-SS-1380-032-M38F-N38F-N-S 01
  const tableItemLine =
    /^\s*(\d+)\s+([A-Z0-9][A-Z0-9./_-]{4,})\s+(\d{1,4})\s*$/i;

  const stopItem = /^(?:\(?Item\s+\d+.*Oportunidade|VER DESVIOS|Item solicitado:|Item orçado:|Notas\/Desvios|Contatos|Índice|Índice de Revisões|\d+\.\s+(?:Apresentação|Objetivo|Documentos de Referência|Escopo de Fornecimento|Notas\/Desvios|Contatos))/i;

  for (const originalLine of lines) {
    const line = clean(originalLine);
    if (!line) continue;

    const pnMatch = line.match(pnLine);
    const tableMatch = line.match(tableItemLine);

    if (pnMatch) {
      finishItem();

      let rest = clean(pnMatch[1]);
      let qty = '';
      let unit = '';
      let total = '';

      const moneyTail = rest.match(
        /\s+(\d+(?:[.,]\d+)?)\s+([\d.,]+)\s+([\d.,]+)\s*$/
      );

      if (moneyTail) {
        qty = moneyTail[1];
        unit = moneyTail[2];
        total = moneyTail[3];
        rest = rest.slice(0, moneyTail.index).trim();
      }

      const split = rest.match(/^(.+?)\s*[–-]\s*\((.*)$/);
      const pn = split
        ? split[1].trim()
        : rest.split(/\s+[–-]\s+/)[0].trim();

      const description = split
        ? '(' + split[2]
        : rest.slice(pn.length).replace(/^\s*[–-]\s*/, '').trim();

      current = {
        pn,
        description,
        qty,
        unit,
        total,
        specs: []
      };
      continue;
    }

    if (tableMatch) {
      finishItem();

      current = {
        pn: tableMatch[2],
        description: '',
        qty: tableMatch[3],
        unit: '',
        total: '',
        specs: []
      };
      continue;
    }

    if (!current) continue;

    if (stopItem.test(line)) {
      finishItem();
      continue;
    }

    // Especificações técnicas que aparecem como bullets no PDF
    if (/^[•●▪]\s*/.test(line)) {
      current.specs.push(line.replace(/^[•●▪]\s*/, ''));
      continue;
    }

    // Também aceita especificações com hífen
    if (/^[-–]\s+\S/.test(line)) {
      current.specs.push(line.replace(/^[-–]\s+/, ''));
      continue;
    }

    // Ignora cabeçalhos e rodapés recorrentes
    if (/^(?:PROPOSTA TÉCNICA|Rua Espírito Santo|CEP:|Tel\.?:|www\.|Pág(?:ina)?\.?\s*\d+|\d+\/\d+)$/i.test(line)) {
      continue;
    }

    // Primeira linha comum após o PN é a descrição do produto
    if (!current.description) {
      current.description = line;
    } else {
      // Continuação de uma descrição quebrada em mais de uma linha
      current.description += ' ' + line;
    }
  }

  finishItem();

  return {
    proposal,
    client,
    date,
    validity,
    payment,
    delivery,
    ncm,
    grandTotal,
    tax: taxLines.join('\n'),
    items: parsedItems,
    raw: all
  };
}

  function setField(id, value){
    document.getElementById(id).value = value || '';
  }

  function renderItems(){
    itemsHost.innerHTML = extractedItems.map((it, idx) => `
      <article class="pdf-item-card" data-index="${idx}">
        <div class="pdf-item-top">
          <strong>Item ${idx+1}</strong>
          <button type="button" class="text-button remove-pdf-item" data-index="${idx}">Remover</button>
        </div>
        <div class="pdf-item-fields">
          <label>Part number
            <input data-field="pn" data-index="${idx}" value="${escapeHtml(it.pn)}">
          </label>
          <label>Quantidade
            <input data-field="qty" data-index="${idx}" value="${escapeHtml(it.qty)}">
          </label>
          <label>Preço unitário
            <input data-field="unit" data-index="${idx}" value="${escapeHtml(it.unit)}">
          </label>
          <label>Valor total
            <input data-field="total" data-index="${idx}" value="${escapeHtml(it.total)}">
          </label>
          <label class="span-all">Descrição técnica
            <textarea data-field="description" data-index="${idx}" rows="5">${escapeHtml(it.description)}</textarea>
          </label>
        </div>
      </article>`).join('');

    itemsHost.querySelectorAll('[data-field]').forEach(el =>
      el.addEventListener('input', () => {
        const idx = Number(el.dataset.index);
        extractedItems[idx][el.dataset.field] = el.value;
        buildFormattedText();
      })
    );

    itemsHost.querySelectorAll('.remove-pdf-item').forEach(btn =>
      btn.addEventListener('click', () => {
        extractedItems.splice(Number(btn.dataset.index), 1);
        renderItems();
        buildFormattedText();
      })
    );
  }

  
function buildFormattedText() {
  const lines = [];

  if (field('pdf-proposal')) {
    lines.push('PROPOSTA: ' + field('pdf-proposal'));
  }

  if (field('pdf-client')) {
    lines.push('CLIENTE: ' + field('pdf-client'));
  }

  if (field('pdf-date')) {
    lines.push('DATA: ' + field('pdf-date'));
  }

  if (field('pdf-validity')) {
    lines.push('VALIDADE: ' + field('pdf-validity'));
  }

  if (field('pdf-payment')) {
    lines.push('CONDIÇÕES DE PAGAMENTO: ' + field('pdf-payment'));
  }

  if (field('pdf-delivery')) {
    lines.push('PRAZO DE ENTREGA: ' + field('pdf-delivery'));
  }

  if (field('pdf-ncm')) {
    lines.push('CLASSIFICAÇÃO FISCAL / NCM: ' + field('pdf-ncm'));
  }

  if (field('pdf-tax')) {
    lines.push('\nIMPOSTOS E OBSERVAÇÕES FISCAIS');
    lines.push(field('pdf-tax'));
  }

  if (extractedItems.length) {
    lines.push('\nESCOPO DE FORNECIMENTO');

    extractedItems.forEach((it, idx) => {
      lines.push('\nITEM ' + (idx + 1));

      if (it.pn) {
        lines.push('PN: ' + it.pn);
      }

      const descriptionLines = String(it.description || '')
        .split(/\r?\n/)
        .map(s => s.trim())
        .filter(Boolean);

      if (descriptionLines.length) {
        lines.push('Descrição: ' + descriptionLines[0]);

        if (descriptionLines.length > 1) {
          lines.push('Especificações técnicas:');

          descriptionLines.slice(1).forEach(spec => {
            lines.push(
              '• ' + spec.replace(/^[•●▪\-–]\s*/, '')
            );
          });
        }
      }

      if (it.qty) {
        lines.push('Quantidade: ' + it.qty);
      }

      if (it.unit) {
        lines.push('Preço unitário: ' + it.unit);
      }

      if (it.total) {
        lines.push('Valor total do item: ' + it.total);
      }
    });
  }

  if (field('pdf-grand-total')) {
    lines.push('\nVALOR TOTAL DA PROPOSTA: ' + field('pdf-grand-total'));
  }

  formatted.value = lines.join('\n');
}

  document.getElementById('extract-pdf').addEventListener('click', async () => {
    const file = fileInput.files && fileInput.files[0];

    if(!file){
      status.className = 'err';
      status.textContent = 'Selecione um arquivo PDF antes de extrair.';
      return;
    }

    if(!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf'){
      status.className = 'err';
      status.textContent = 'O arquivo selecionado não parece ser um PDF.';
      return;
    }

    status.className = '';
    status.textContent = 'Lendo o PDF localmente no navegador...';
    extractBtn.disabled = true;

    try{
      const lines = await extractLines(file);
      const result = parsePdfText(lines);

      setField('pdf-proposal', result.proposal);
      setField('pdf-client', result.client);
      setField('pdf-date', result.date);
      setField('pdf-validity', result.validity);
      setField('pdf-payment', result.payment);
      setField('pdf-delivery', result.delivery);
      setField('pdf-ncm', result.ncm);
      setField('pdf-grand-total', result.grandTotal);
      setField('pdf-tax', result.tax);

      extractedItems = result.items;
      rawText.value = result.raw;
      editor.classList.remove('hidden');

      renderItems();
      buildFormattedText();

      status.className = result.raw.trim() ? 'ok' : 'err';
      status.textContent = result.raw.trim()
        ? 'Extração concluída: ' + lines.filter(Boolean).length + ' linhas de texto e ' + extractedItems.length + ' item(ns) identificados. Revise os campos antes de copiar.'
        : 'Não foi possível extrair texto. O PDF pode ser digitalizado como imagem; esta versão não inclui OCR.';

    } catch(err){
      status.className = 'err';
      status.textContent = 'Não foi possível ler o PDF: ' + err.message;
    } finally {
      extractBtn.disabled = false;
    }
  });

  document.getElementById('add-pdf-item').addEventListener('click', () => {
    extractedItems.push({
      pn: '',
      description: '',
      qty: '',
      unit: '',
      total: ''
    });
    renderItems();
    buildFormattedText();
  });

  [
    'pdf-proposal',
    'pdf-client',
    'pdf-date',
    'pdf-validity',
    'pdf-payment',
    'pdf-delivery',
    'pdf-ncm',
    'pdf-grand-total',
    'pdf-tax'
  ].forEach(id => {
    document.getElementById(id).addEventListener('input', buildFormattedText);
  });

  document.getElementById('copy-pdf-text').addEventListener('click', async () => {
    const copyStatus = document.getElementById('pdf-copy-status');

    try{
      await navigator.clipboard.writeText(formatted.value);
      copyStatus.textContent = 'Texto copiado.';
    } catch(err){
      formatted.focus();
      formatted.select();
      const ok = document.execCommand('copy');
      copyStatus.textContent = ok
        ? 'Texto copiado.'
        : 'Selecione o texto e use Ctrl+C.';
    }

    setTimeout(() => {
      copyStatus.textContent = '';
    }, 3500);
  });
})();