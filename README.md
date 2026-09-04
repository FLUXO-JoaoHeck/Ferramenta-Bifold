# Composição de Preço & Proposta

Ferramenta em página única (HTML/CSS/JS), sem backend, para:
1. Extrair dados de um PDF de fornecedor (texto)
2. Editar/conferir esses dados numa interface
3. Montar a composição de preço (compra → CMV → preço de saída)
4. Gerar uma proposta comercial em PDF para enviar ao cliente

Tudo roda no navegador — não precisa de servidor, banco de dados ou chave de API.

## Como publicar no GitHub Pages

1. Crie um repositório novo no GitHub (pode ser público ou privado, desde que o plano permita Pages).
2. Suba o arquivo `index.html` para a raiz do repositório (via upload no site do GitHub ou `git push`).
3. Vá em **Settings → Pages**.
4. Em "Source", selecione a branch `main` e a pasta `/ (root)`.
5. Salve. Em alguns minutos o GitHub mostra a URL pública, algo como:
   `https://seu-usuario.github.io/nome-do-repositorio/`

Qualquer atualização no `index.html` (novo commit) atualiza o site automaticamente.

## Como funciona a extração do PDF

A extração usa a biblioteca `pdf.js` para ler o texto do PDF e procura, por
palavras-chave, três informações: fornecedor, produto/descrição e um valor
em R$. Como você indicou que o PDF de entrada sempre segue o mesmo modelo,
vale a pena refinar isso:

No arquivo `index.html`, procure o bloco:

```js
const PATTERNS = {
  fornecedor: /.../,
  produto: /.../,
  precoCompra: /.../
};
```

Ajuste essas expressões regulares para o texto exato que aparece no seu
PDF (ex.: se o fornecedor sempre aparece depois de "Emitente:", troque o
regex correspondente). Se quiser, me envie um PDF de exemplo (ou o texto
extraído dele) que eu ajusto os padrões junto com você.

**Importante:** a extração é só um atalho — o passo 2 da ferramenta sempre
permite conferir e corrigir manualmente antes de gerar a proposta, então um
erro de extração nunca vira um erro na proposta final.

## Fórmula de composição de preço (markup divisor)

Método padrão de precificação "por dentro", comum no Brasil:

```
Créditos de entrada = Preço de compra × (ICMS% + PIS% + COFINS% de entrada)
CMV = Preço de compra − Créditos de entrada + (Preço de compra × IPI% não recuperável) + Outras despesas

% total de saída = ICMS% + PIS% + COFINS% + ISS% (saída) + Margem% desejada

Preço de saída = CMV / (1 − % total de saída)
Impostos de saída (R$) = Preço de saída × (ICMS% + PIS% + COFINS% + ISS% de saída)
Margem de saída (R$) = Preço de saída × Margem% desejada
```

Se a soma de impostos de saída + margem chegar a 100% ou mais, a ferramenta
avisa e não calcula (matematicamente o preço tenderia ao infinito).

## Personalizações comuns

- **Impostos diferentes por regime** (Simples Nacional, Lucro Presumido,
  Lucro Real): os campos de % são livres, então basta digitar a alíquota
  correta em cada campo — não há lógica de regime fixa no código.
- **Logotipo na proposta**: adicione uma tag `<img>` dentro de
  `#proposal-template .p-header` no `index.html`.
- **Mais itens por proposta**: hoje a proposta trata um item por vez; para
  múltiplos itens seria necessário guardar uma lista em vez de campos
  únicos — posso construir essa versão se fizer sentido para o seu uso.

## Estrutura do arquivo

Um único arquivo `index.html` contém HTML, CSS e JavaScript. As bibliotecas
externas (pdf.js para leitura do PDF, html2pdf.js para gerar o PDF da
proposta) são carregadas via CDN, então é preciso estar online para usar a
ferramenta.
