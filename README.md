# Tabela de Preços e Impostos

Site estático (`index.html`, `style.css`, `script.js`), sem backend, que lê a
planilha de revenda (Ver. 3.0) e gera a seção "Preços e Impostos" da proposta,
pronta para copiar e colar no Word.

## Como usar
1. Abra o site e escolha o arquivo Excel (aba "Revenda").
2. Confira os itens lidos e os avisos exibidos.
3. Em "Receita Bruta e Petronect", escolha o preço (Revenda direta ou Portal), o estado
   do cliente, como o IPI aparece e se o destino é Zona Franca de Manaus.
4. Clique em "Copiar tabela" e cole (Ctrl+V) no Word.
5. Para o escopo técnico, vá em "Proposta técnica", escolha o PDF da cotação do
   fornecedor e clique em "Copiar proposta técnica".

## O que o site lê da planilha
O site NÃO recalcula preços: lê os valores já calculados pelo Excel
(o arquivo precisa ter sido salvo pelo Excel, com os resultados em cache).

O arquivo pode ter várias abas (Entradas, Lista de Preços, Resumo, NCM...). O
site procura **somente** a aba chamada "Revenda" (sem diferenciar maiúsculas,
acentos ou espaços nas pontas) e lê apenas ela. Se não a encontrar, para e
mostra a lista de abas do arquivo, sem tentar outra aba no lugar.

| Dado | Coluna |
|---|---|
| Item (código) | A |
| Descrição (nome do item, coluna "DESCRIÇÃO") | E |
| NCM | G |
| Quantidade | P |
| ICMS / PIS-COFINS / IPI de saída (%) | R / S / T |
| Preço unitário e total, revenda direta | W / X |
| Preço unitário e total, portal (líquido de ICMS e IPI) | Z / AA |

- Cabeçalho: a linha é detectada automaticamente (coluna A = "ITEM" e coluna X =
  "P. TOTAL", nas primeiras 60 linhas), então a tabela pode começar na linha 13,
  14, 15 etc. Os itens são lidos nas até 100 linhas abaixo dela.
- Proposta e cliente são achados pelo rótulo, acima do cabeçalho; o subtotal é o
  primeiro número da coluna X acima do cabeçalho.
- Itens sem quantidade ou com erro (`#REF!`) nos preços são ignorados, com aviso
  na tela. Se o cabeçalho não for encontrado, a leitura para com uma mensagem.
- Nome do item: o site procura a coluna **"DESCRIÇÃO"** no cabeçalho (qualquer
  posição) e a usa como identificação na coluna "Modelo" da tabela. Se não achar,
  usa a coluna E e avisa na tela.
- A soma dos itens é conferida com o subtotal da planilha quando todos os itens
  são lidos.

## Proposta técnica (PDF do fornecedor)
O site lê o PDF da cotação e monta a tabela **Escopo de Fornecimento**
(ITEM | DESCRIÇÃO | QTD), no formato do modelo Word.

- Cada item é reconhecido pela linha `P/N: <código> – (<descrição>)`; a
  quantidade é o primeiro número do trio "quantidade, preço unitário, valor" que
  vem depois da descrição.
- A descrição é separada em tópicos "Rótulo: valor" (uma vírgula seguida de um
  rótulo com dois-pontos inicia um novo tópico). Os textos ficam como o
  fornecedor escreveu, inclusive "N/A".
- Quebras de linha no meio do código (por exemplo `XW-00-` / `01580-...`) são
  juntadas; cabeçalho e rodapé repetidos em cada página são descartados (lista
  `PDF_NOISE`, no início da seção em `script.js`).
- O pdf.js só é baixado quando um PDF é escolhido.
- Os preços do PDF não são lidos nem exibidos: a proposta técnica não os usa.
- Layout testado: cotação da Rotork. Outros fornecedores precisam usar a mesma
  convenção "P/N:"; senão o site informa que não encontrou itens.

## Publicar no GitHub Pages
Suba os três arquivos juntos para a raiz do repositório e ative o Pages em
Settings → Pages (branch `main`, pasta `/ (root)`).

## Observações
- A biblioteca de leitura de Excel (SheetJS) é carregada por CDN: é preciso
  estar online.
- A formatação da tabela copiada segue o documento original (Open Sans, 8 pt,
  linhas horizontais finas, alinhamento centralizado, linhas de 2 cm).
