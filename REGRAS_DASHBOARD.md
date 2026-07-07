# Regras do Dashboard Motiva — Referência Completa
> Atualizado em 02/07/2026 — correções críticas de GMV, cartões e filtro PE

---

## 1. Fontes de Dados

| Planilha | Aba | Uso |
|----------|-----|-----|
| Excel Lorena | Acompanhamento | GMV, recVT, recPAT, financiamento, 2°Via |
| Excel Lorena | Pedido de carga Completo (PCC) | Taxas VT, NF, comissões VT, bilhetagens |
| Excel Lorena | Pedido externo (PE) | Taxas PE (VT externo), comissões PE |
| Excel Lorena | Clientes | Base de clientes, cadastros por mês |
| Excel Lorena | GMV por Cliente | Segmentação, GMV por empresa |

---

## 2. Colunas Utilizadas por Aba

### PCC (Pedido de carga Completo)
| Coluna | Letra | Variável JS | Uso |
|--------|-------|-------------|-----|
| taxa_administrativa | P | vtTxAdm (parte) | Tx Adm VT |
| **taxa_aproveitamento** | **R** | **vtTxSucesso (parte)** | **Tx Sucesso VT — USAR COL R, não Col Q** |
| taxa_entrega | S | vtTxEntrega (via Acomp row29) | Tx Entrega (2°Via) |
| nota_fiscal | T | nfPedCarga | NF para composição de receita |
| Comissão | AA | comm_vt | Comissões dos vendedores |
| Codigo_bilhetagem | C | cartoesMes (contagem) | Total cartões processados |
| mes_pagamento | I | agrupamento | Mês de referência |

### PE (Pedido externo)
| Coluna | Letra | Uso |
|--------|-------|-----|
| Mês | **E** | **Referência de mês (coluna E)** |
| Taxa Adm | N | vtTxAdm (parte PE) |
| Tx de Sucesso | O | vtTxSucesso (parte PE) |
| Comissão | X | comm_pat (comissões vendedores) |

#### Regra BaseSK no PE:
| BaseSK | Tipo | GMV | Taxas | NF/Receita | Comissão |
|--------|------|-----|-------|------------|----------|
| "Sem SK" | VT externo | ✅ incluir no VD | ✅ incluir em vtTxAdm/vtTxSucesso | ✅ | ✅ |
| [número] | Já no PCC como VT | ❌ não duplicar no VD | ✅ incluir taxas | ✅ | ✅ |

**IMPORTANTE:** Todos os pedidos PE (com e sem SK) têm suas taxas somadas ao VT.
A coluna E (Mês) é a referência para agrupamento mensal no PE.

---

## 3. Cálculo das Variáveis JS

```
vtTxAdm     = PCC taxa_administrativa (col P) + PE Taxa Adm (col N) — TODOS pedidos PE
vtTxSucesso = PCC taxa_aproveitamento (col R) + PE Tx de Sucesso (col O) — SEM subtrair financiamento
              (Col R já exclui financiamento por natureza)
vtTxEntrega = Acompanhamento linha 29 (2°Via receitas)
financiamento = Acompanhamento linha 31 (separado, não entra no VT)

recVT       = vtTxAdm + vtTxSucesso + vtTxEntrega  ← calculado no Python, substituído via replace_const
recPAT      = Acompanhamento linha 36
recTotal    = recVT + recPAT + financiamento  ← calculado no JS: recVT.map((v,i)=>v+recPAT[i]+financiamento[i])

nfPedCarga  = PCC nota_fiscal (col T)
cartoesMes  = contagem de Codigo_bilhetagem por mes_pagamento no PCC
```

### Verificação: soma das barras deve = recVT
`vtTxAdm + vtTxSucesso + vtTxEntrega = recVT` ✓ (exceto meses parciais)

### Financiamento
- Vem **apenas** do Acompanhamento (linha 31)
- Exibido como card/barra **separado** nas abas GMV e Receitas
- **NÃO** entra em vtTxAdm nem vtTxSucesso
- **É** somado em recTotal (Rec. Total = VT + PAT + Financiamento)

---

## 4. Comissões (VD)

```python
# Comissões: PE completo (Sem SK + numerados) — todos têm comissão real
comm_pat = df_pe_full.pivot_table(index='mes_str', columns='Vendedor', values='Comissão', ...)

# VD por empresa: Sem SK = GMV incluso; [número] = só NF + comissão
is_sem_sk = 'SEM SK' in str(row.get('BaseSK', '')).upper()
if is_sem_sk: entries[key]['gmv'][i] += float(row['Valor do Boleto Motiva'])
entries[key]['nf'][i]   += float(row['NF'])
entries[key]['comm'][i] += float(row['Comissão'])
```

---

## 5. Aba Controle Cliente (_CC)

Dados extraídos de PCC + PE, agrupados por empresa/mês.

**PCC:**
- Valor Compra: col M (valor_recarga)
- Projeção Saldo: col N (projecao_saldo)
- Carga Realizada: col O (Carga Realizada)
- Economia Líquida: col N - col R (projecao_saldo - taxa_aproveitamento)
- Cartões: contagem de Codigo_bilhetagem

**PE (coluna E = Mês):**
- Valor Compra: col I (Valor Recarga (I))
- Projeção Saldo: col K (Projeção de Saldo)
- Carga Realizada: col J (Valor Recarga (F))
- Economia Líquida: col K - col O (Projeção de Saldo - Tx de Sucesso)
- Cartões: contagem de pedidos

**Estrutura _CC:** `[cnpj, empresa, mes_idx, compra, projecao, carga, economia, cartoes, origem]`
- origem: 0 = VT (PCC), 1 = VT Externo (PE)

---

## 6. Card Cartões Processados (Aba GMV)

```python
# cartoesMes = contagem de bilhetagens PCC por mes_pagamento
cartoesMes = [count por mês de Codigo_bilhetagem no PCC]
```

Inserir no script principal (antes de applyGmv), não no script CC.

---

## 7. Regras de Exibição (HTML/JS)

### Aba GMV
- recTotal no JS: `const recTotal = recVT.map((v,i) => v + (recPAT[i]||0) + (financiamento[i]||0))`
- Card "Cartões Processados": usa `cartoesMes[li]`

### Aba Receitas
- Card "Rec. VT": mostra `recVT[li]` diretamente (= vtTxAdm + vtTxSucesso + vtTxEntrega)
- Card "Rec. Total": mostra `recTotal[li]` (já inclui financiamento via JS)
- Tabela "Detalhe Mensal": Tx Adm | Tx Sucesso | Tx Entrega | Financiamento | Rec.VT | Rec.PAT | Rec.Total | Margem

### Aba Comissões
- Tabela empresas: sempre visível ao selecionar vendedor (try/catch)
- Linha TOTAL no Resumo por Vendedor

### Aba Controle Cliente
- Gráfico: 5 métricas — Valor Compra | Proj. Saldo | Carga | Economia | Cartões
- Tabela com comparação de meses: ambos os meses lado a lado + Δ Variação
- Cartões formatados como número inteiro (não R$)

---

## 8. Problema Conhecido: Truncamento do HTML

O HTML (~1.3MB) pode ser truncado ao escrever diretamente no workspace (OneDrive).

**Solução:** Sempre usar o git como fonte de verdade:
1. `git clone` do repositório
2. Modificar o clone
3. `git push`
4. Copiar para workspace apenas ao final

```python
GIT_DIR = "/tmp/motiva-update-TIMESTAMP"  # nunca reusar /tmp/motiva-dashboard (permissão)
```

---

## 9. Correções Aplicadas em 02/Jul/2026

### 9.1 GMV PCC — coluna correta
- ❌ **Errado (antigo):** `total_pedido` (col L) — inclui taxas administrativas
- ✅ **Correto:** `valor_recarga` (col M) — valor líquido de recarga
- Diferença: R$ 5,6 milhões de GMV inflado corrigido

### 9.2 PE filtro BaseSK para GMV
- ✅ `BaseSK = "Sem SK"` → conta como GMV e projeção de saldo
- ❌ `BaseSK = [número]` → **excluído do GMV**, mas receita/taxas mantidas
- Impacto: 572 linhas PE excluídas do GMV (R$ 1,748,277 removidos)

### 9.3 Cartões — soma real
- ❌ **Antigo:** `count(Codigo_bilhetagem)` = contagem de pedidos (linhas)
- ✅ **Correto:** `sum(Qtd Cartões)` = coluna PCC índice 29, soma de beneficiários por pedido

### 9.4 last_complete — critério de mês completo
- ❌ **Antigo:** `gmv_pat > 0` → nunca avançava (PAT encerrado)
- ✅ **Correto:** `gmv_total > 0` (VT + PAT) — funciona com PAT = 0

### 9.5 Motiva não opera mais PAT
- `gmvPAT` sempre = 0 a partir de 2026
- Dashboard mantém histórico PAT mas não recebe novos dados
- `recPAT` também = 0

### 9.6 Regras _RD (GMV por empresa / Resumo)
- `gmv_arr`: PCC `valor_recarga` (todos) + PE `Valor Recarga (I)` (só Sem SK)
- `nf_arr`: PCC `nota_fiscal` (todos) + PE `NF` (todos — receita inclui Com SK)

### 9.7 Regras _CC (Controle por Cliente)
- PCC: todos os pedidos, `valor_recarga` como GMV, `Qtd Cartões` como cartões
- PE: apenas `BaseSK = "Sem SK"`, `Valor Recarga (I)` como GMV
- Comissões/taxas: todo PE independente de SK

### 9.8 Novas funcionalidades
- Botão `📄 Exportar PDF` adicionado à aba Resumo
- Clientes mesclados por CNPJ em `_CC` e `_RD`

---

## 10. Valores de Referência — Jun/26 (último mês completo desde 02/Jul/2026)

| Métrica | Valor |
|---------|-------|
| GMV VT | R$ 5.021.844 |
| GMV PAT | R$ 0 (PAT encerrado) |
| Receita VT | R$ 216.245 |
| Take Rate VT | 4,31% |
| Cartões | 18.027 |
| Empresas únicas | 831 CNPJs |
| Entradas _CC | 9.496 |

---

## 11. Valores de Referência — Abr/26 (histórico)

| Métrica | Valor |
|---------|-------|
| GMV VT | R$ 4.352.421 |
| Tx Adm VT | R$ 78.493 (PCC + PE) |
| Tx Sucesso VT | R$ 106.147 (Col R + PE, sem subtrair Fin) |
| Tx Entrega | R$ 225 |
| Rec. VT Total | R$ 184.865 (= soma das 3 taxas) |
| Financiamento | R$ 12.319 |
| Rec. PAT | R$ 71.056 |
| Rec. Total | R$ 268.240 (VT + PAT + Fin) |
| Cartões | 1.721 |
| Karine Comissão | R$ 1.791,33 |
| Segmentação Abr/26 | D:18 O:76 P:93 BP:250 BPP:62 |

---

## 12. Regra de Segmentação de Clientes (GMV Rule Based)

Segmentação de cada cliente (CNPJ) pelo **maior GMV mensal** observado no período analisado
(`max(gmv[i] for i in periodo)`), usando a mesma base de dados do GMV VT oficial:

| Segmento | Faixa (maior GMV mensal) |
|----------|---------------------------|
| Diamante | >= R$ 50.000 |
| Ouro | R$ 10.000 – R$ 50.000 |
| Prata | R$ 5.000 – R$ 10.000 |
| Bronze-P | R$ 1.000 – R$ 5.000 |
| Bronze-PP | < R$ 1.000 |

### Fonte dos dados (`_GMVSEG`)
Calculado diretamente de PCC + PE — mesma Regra de GMV usada em `_RD` e `_CC` (§9.6/9.7):
- PCC: `valor_recarga` (col M), **todos** os pedidos
- PE: `Valor Recarga (I)` (col I), **apenas** linhas com `BaseSK = "Sem SK"`
- Agrupado por CNPJ (`pad_cnpj`) + mês (`mes_pagamento`)

Substitui a antiga fonte ("GMV por Cliente", aba separada da planilha), que apresentava pequenas
divergências de reconciliação frente ao `gmv_vt` oficial. Ao usar a mesma base de PCC+PE do
restante do pipeline, os números do card de segmentação ficam consistentes com os demais
cálculos de GMV do dashboard (`_RD`, `_CC`, cards da aba GMV).

### Nota de reconciliação (atualizado 07/Jul/2026)
**Correção validada pelo usuário:** o PCC deve **excluir pedidos com `tipo_pedido = "Deposit"`**
— esses não são recargas reais do cliente. Com essa exclusão, a reconciliação com `gmv_vt`
do Acompanhamento passa a ser **exata** (diferença residual de centavos, arredondamento):

| Mês | GMV via PCC(sem Deposit)+PE(Sem SK) | `gmv_vt` oficial | Diferença |
|-----|--------------------------------------|--------------------|-----------|
| Abr/26 | R$ 4.352.421,13 | R$ 4.352.421,00 | R$ 0,13 |
| Mai/26 | R$ 4.570.791,19 | R$ 4.570.791,00 | R$ 0,19 |
| Jun/26 | R$ 5.021.843,84 | R$ 5.021.844,00 | -R$ 0,16 |

**Fórmula definitiva do GMV Rule:**
```
GMV = PCC.valor_recarga [tipo_pedido != "Deposit"]  (todos os demais tipos: Dealer, Broker, SecondCopy, FirstCopy)
    + PE."Valor Recarga (I)" [BaseSK contém "Sem SK"]
```

Esta correção foi aplicada em `_GMVSEG` (card GMV VT por Segmento) e em `_CC` (Controle
Cliente) no `update_vars.py`. **Pendente:** `_RD` (tabela "Detalhe por Empresa" da aba
Receitas) ainda é um array estático no HTML, não gerado pelo pipeline — ainda não recebeu
esta correção (ver §8 sobre limitações de geração automática).

---

## 13. GitHub Pages
URL: https://vitorabl.github.io/motiva-dashboard/
Repo: config em `Automacao/github_config.txt` (REMOTE_URL)
Deploy: ~2 min após push. Hard refresh: Ctrl+Shift+R ou aba anônima.
