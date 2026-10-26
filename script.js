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
   Extrai exclusivamente o PN e a descrição técnica.
   Cada especificação separada por vírgula vira um bullet point.
--------------------------------------------------------------------- */
(function initPdfExtractor() {
  const fileInput = document.getElementById('pdf-input');
  if (!fileInput) return;

  const extractBtn = document.getElementById('extract-pdf');
  const status = document.getElementById('pdf-status');
  const editor = document.getElementById('pdf-editor');
  const itemsHost = document.getElementById('pdf-items');
  const formatted = document.getElementById('pdf-formatted-text');
  const rawText = document.getElementById('pdf-raw-text');

  let extractedItems = [];

  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }

  function clean(text) {
    return String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function escapeValue(text) {
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /* ---------------------------------------------------------------
     LEITURA DO TEXTO DO PDF
  ---------------------------------------------------------------- */
  async function extractLines(file) {
    if (!window.pdfjsLib) {
      throw new Error(
        'PDF.js não carregou. Verifique a conexão com a internet.'
      );
    }

    const buffer = await file.arrayBuffer();

    const pdf = await pdfjsLib.getDocument({
      data: buffer
    }).promise;

    const lines = [];

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const grouped = [];

      content.items.forEach(item => {
        if (!item.str || !item.str.trim()) return;

        const x = item.transform ? item.transform[4] : 0;
        const y = item.transform ? item.transform[5] : 0;

        let line = grouped.find(
          entry => Math.abs(entry.y - y) < 2.2
        );

        if (!line) {
          line = {
            y: y,
            parts: []
          };
          grouped.push(line);
        }

        line.parts.push({
          x: x,
          text: item.str
        });
      });

      grouped.sort((a, b) => b.y - a.y);

      grouped.forEach(line => {
        line.parts.sort((a, b) => a.x - b.x);

        const text = clean(
          line.parts.map(part => part.text).join(' ')
        );

        if (text) lines.push(text);
      });

      lines.push('');
    }

    return lines;
  }

  /* ---------------------------------------------------------------
     IDENTIFICA O PN E O INÍCIO DA DESCRIÇÃO

     Exemplo:
     P/N: VRT-SS-0800-030-C-C-N-U - (Alternative offering for: ...
  ---------------------------------------------------------------- */
  function parsePnLine(line) {
    const marker = line.match(/P\s*\/\s*N\s*:\s*(.*)/i);

    if (!marker) return null;

    const content = clean(marker[1]);
    if (!content) return null;

    // Procura o separador entre o PN e a descrição.
    const separator = content.match(/\s+[-–]\s+\(?/);

    let pn = '';
    let description = '';

    if (separator) {
      const position = separator.index;

      pn = content.slice(0, position).trim();

      description = content
        .slice(position + separator[0].length)
        .trim();
    } else {
      // Alternativa para PDFs que não usam " - " entre os campos.
      const firstSpace = content.search(/\s/);

      if (firstSpace === -1) {
        pn = content;
      } else {
        pn = content.slice(0, firstSpace).trim();
        description = content.slice(firstSpace).trim();
      }
    }

    // Remove parênteses externos da descrição.
    description = description
      .replace(/^[\s(]+/, '')
      .replace(/[\s)]+$/, '')
      .trim();

    if (!pn) return null;

    return {
      pn: pn,
      description: description
    };
  }

  /* ---------------------------------------------------------------
     IDENTIFICA SOMENTE OS ITENS QUE CONTÊM P/N

     As linhas de descrição continuam sendo adicionadas ao item atual
     até aparecer outro P/N ou o início de outra seção do documento.
  ---------------------------------------------------------------- */
  function parsePdfText(lines) {
    const parsedItems = [];
    let current = null;

    const stopSection = /^(?:NOTAS\/DESVIOS|CONTATOS|ÍNDICE(?:\s|$)|ÍNDICE DE REVISÕES|ATENCIOSAMENTE|DOCUMENTOS DE REFERÊNCIA|ESCOPO DE FORNECIMENTO|TOTAL GERAL|CLASSIFICAÇÃO FISCAL:|ESTAMOS CONSIDERANDO)/i;

    function finishItem() {
      if (!current || !current.pn) {
        current = null;
        return;
      }

      // Vírgula é o separador das especificações técnicas.
      const specs = current.description
        .replace(/\s+/g, ' ')
        .replace(/^[\s(]+/, '')
        .replace(/[\s),]+$/, '')
        .split(/,\s*/)
        .map(spec => spec.trim())
        .filter(Boolean);

      parsedItems.push({
        pn: current.pn,
        description: specs
      });

      current = null;
    }

    for (const originalLine of lines) {
      const line = clean(originalLine);

      if (!line) continue;

      const parsedPn = parsePnLine(line);

      if (parsedPn) {
        finishItem();

        current = {
          pn: parsedPn.pn,
          description: parsedPn.description
        };

        continue;
      }

      if (!current) continue;

      // Não incorporar títulos de outras seções nem rodapés.
      if (stopSection.test(line)) {
        finishItem();
        continue;
      }

      if (
        /^PROPOSTA TÉCNICA\b/i.test(line) ||
        /^Rua Espírito Santo\b/i.test(line) ||
        /^CEP:/i.test(line) ||
        /^Tel\.?:/i.test(line) ||
        /^www\./i.test(line) ||
        /^\d+\s*\/\s*\d+$/.test(line) ||
        /^Pág(?:ina)?\.?\s*\d+/i.test(line)
      ) {
        continue;
      }

      // Continua a descrição do item atual.
      current.description += ' ' + line;
    }

    finishItem();

    return parsedItems;
  }

  /* ---------------------------------------------------------------
     EXIBIÇÃO DOS ITENS EDITÁVEIS
     Apenas PN e descrição técnica.
  ---------------------------------------------------------------- */
  function renderItems() {
    itemsHost.innerHTML = extractedItems.map((item, index) => `
      <article class="pdf-item-card" data-index="${index}">
        <div class="pdf-item-top">
          <strong>Item ${index + 1}</strong>
          <button
            type="button"
            class="text-button remove-pdf-item"
            data-index="${index}">
            Remover
          </button>
        </div>

        <div class="pdf-item-fields">
          <label>
            PN
            <input
              data-field="pn"
              data-index="${index}"
              value="${escapeValue(item.pn)}">
          </label>

          <label class="span-all">
            Descrição técnica — uma especificação por linha
            <textarea
              data-field="description"
              data-index="${index}"
              rows="8">${escapeValue(item.description.join('\n'))}</textarea>
          </label>
        </div>
      </article>
    `).join('');

    itemsHost.querySelectorAll('[data-field]').forEach(input => {
      input.addEventListener('input', () => {
        const index = Number(input.dataset.index);
        const field = input.dataset.field;

        if (field === 'pn') {
          extractedItems[index].pn = input.value;
        }

        if (field === 'description') {
          extractedItems[index].description = input.value
            .split(/\r?\n|,\s*/)
            .map(spec => spec.replace(/^[•●▪\-–]\s*/, '').trim())
            .filter(Boolean);
        }

        buildFormattedText();
      });
    });

    itemsHost.querySelectorAll('.remove-pdf-item').forEach(button => {
      button.addEventListener('click', () => {
        extractedItems.splice(Number(button.dataset.index), 1);
        renderItems();
        buildFormattedText();
      });
    });
  }

  /* ---------------------------------------------------------------
     RESULTADO FINAL

     Exemplo:
     PN: VRT-SS-0800-030-C-C-N-U

     Descrição:
     • Alternative offering for: 14540-03
     • Product: Valve - Relief Thermal
     • Option: None
  ---------------------------------------------------------------- */
  function buildFormattedText() {
    const output = [];

    extractedItems.forEach((item, index) => {
      if (index > 0) {
        output.push('');
        output.push('');
      }

      output.push('PN: ' + item.pn);
      output.push('');
      output.push('Descrição:');

      item.description.forEach(specification => {
        output.push('• ' + specification);
      });
    });

    formatted.value = output.join('\n');
  }

  /* ---------------------------------------------------------------
     BOTÃO DE EXTRAÇÃO
  ---------------------------------------------------------------- */
  extractBtn.addEventListener('click', async () => {
    const file = fileInput.files && fileInput.files[0];

    if (!file) {
      status.className = 'err';
      status.textContent = 'Selecione um arquivo PDF.';
      return;
    }

    status.className = '';
    status.textContent = 'Extraindo PN e descrições técnicas...';
    extractBtn.disabled = true;

    try {
      const lines = await extractLines(file);
      extractedItems = parsePdfText(lines);

      if (rawText) rawText.value = '';

      editor.classList.remove('hidden');

      renderItems();
      buildFormattedText();

      if (extractedItems.length) {
        status.className = 'ok';
        status.textContent =
          'Extração concluída: ' +
          extractedItems.length +
          ' item(ns) identificado(s). Revise os PNs e as especificações.';
      } else {
        status.className = 'err';
        status.textContent =
          'Nenhum PN foi identificado. Verifique se o PDF contém texto selecionável e se utiliza o formato P/N:.';
      }
    } catch (error) {
      status.className = 'err';
      status.textContent =
        'Erro ao extrair o PDF: ' + error.message;
    } finally {
      extractBtn.disabled = false;
    }
  });

  /* Adiciona um item manualmente */
  const addButton = document.getElementById('add-pdf-item');

  if (addButton) {
    addButton.addEventListener('click', () => {
      extractedItems.push({
        pn: '',
        description: []
      });

      renderItems();
      buildFormattedText();
    });
  }

  /* Copia somente PN e descrição formatados */
  const copyButton = document.getElementById('copy-pdf-text');

  if (copyButton) {
    copyButton.addEventListener('click', async () => {
      const copyStatus = document.getElementById('pdf-copy-status');

      try {
        await navigator.clipboard.writeText(formatted.value);
        copyStatus.textContent = 'PN e descrição copiados.';
      } catch (error) {
        formatted.focus();
        formatted.select();

        const success = document.execCommand('copy');

        copyStatus.textContent = success
          ? 'PN e descrição copiados.'
          : 'Selecione o texto e pressione Ctrl+C.';
      }

      setTimeout(() => {
        copyStatus.textContent = '';
      }, 3500);
    });
  }
})();