# Composição de Preço

Ferramenta em três arquivos (`index.html`, `style.css`, `script.js`), sem backend, para:
1. (Opcional) extrair o custo de um PDF de fornecedor
2. Definir as alíquotas de ICMS, PIS/COFINS, IPI, ISS e a margem desejada
3. Montar os itens (PN, descrição, NCM, custo, quantidade) — o preço é calculado automaticamente
4. Gerar uma tabela pronta para copiar e colar formatada em um Word

Tudo roda no navegador — não precisa de servidor, banco de dados ou chave de API.

## Como publicar no GitHub Pages

1. Crie um repositório novo no GitHub.
2. Suba os três arquivos juntos (`index.html`, `style.css`, `script.js`) para a raiz do repositório — os três precisam estar na mesma pasta.
3. Vá em **Settings → Pages**, selecione a branch `main` e a pasta `/ (root)`, salve.
4. Em alguns minutos o GitHub mostra a URL pública (`https://seu-usuario.github.io/nome-do-repositorio/`).

## Fórmula de composição de preço

```
Créditos de entrada:
  IPI(R$)        = Custo / (1 + IPI%_entrada) x IPI%_entrada
  ICMS(R$)       = (Custo - IPI(R$)) x ICMS%_entrada
  PIS/COFINS(R$) = (Custo - IPI(R$) - ICMS(R$)) x PIS/COFINS%_entrada

CMV = Servico(R$) + (Custo - IPI(R$) - ICMS(R$) - PIS/COFINS(R$))

RecLiq = CMV / (1 - Margem%)
DenomSaida = 1 - ICMS%_saida - ICMS%_saida x IPI%_saida - PIS/COFINS%_saida
Preco final = RecLiq / DenomSaida
```

"Servico(R$)" e o valor de frete (se ativado), dividido igualmente entre a
quantidade total de todos os itens.

## Copiar a tabela para o Word

O botao "Copiar tabela" seleciona a tabela final e copia via
`document.execCommand('copy')`, preservando negrito, bordas e alinhamento -
basta colar (Ctrl+V) direto no Word. Se o navegador bloquear a copia
automatica, selecione a tabela manualmente e use Ctrl+C.

## Como funciona a extracao do PDF

Best-effort: procura por "Valor Unitario/Preco" e "NCM" no texto do PDF via
`pdf.js`. Ajuste os regex em `PATTERNS`, no inicio de `script.js`, para o
layout real do seu fornecedor.

## Estrutura dos arquivos

- `index.html` - estrutura da pagina
- `style.css` - aparencia, modo claro/escuro, estilo da tabela final
- `script.js` - extracao do PDF, calculo, tabela de saida e copia para o Word
