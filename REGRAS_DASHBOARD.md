# Regras do Dashboard Motiva — Referência Completa
> Atualizado em 03/06/2026 — compilado após sessão completa de refinamento

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

## 9. Valores de Referência — Abr/26 (último mês completo)

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

## 10. GitHub Pages
URL: https://vitorabl.github.io/motiva-dashboard/
Repo: config em `Automacao/github_config.txt` (REMOTE_URL)
Deploy: ~2 min após push. Hard refresh: Ctrl+Shift+R ou aba anônima.
