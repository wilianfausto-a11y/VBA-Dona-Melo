# Entrega de Código — Projeto DonaMelo
## Data: 31/05/2026

---

## Arquivos entregues

| Arquivo | Tipo | Descrição |
|---|---|---|
| `modGlobais.bas` | Módulo | Constantes de colunas, ValidarCPF, GravarMovimentacao centralizado |
| `EsteLivro.bas` | ThisWorkbook | Workbook_Open → abre frmLogin |
| `Módulo1.bas` | Módulo | CalcularTotaisProduto, GerarCodigoInternoUnico, GerarIDRomaneio |
| `frmLogin_corrigido.bas` | UserForm | Login refatorado → abre frmMenu |
| `frmMenu.bas` | UserForm | Tela de navegação pós-login |
| `frmRomaneio.bas` | UserForm | Romaneio completo — 3 abas |

---

## Como aplicar no projeto

> Todos os arquivos `.bas` devem ser importados no VBA Editor (Alt+F11) substituindo os módulos existentes de mesmo nome.

### Passo a passo:
1. Abra o arquivo `.xlsm` no Excel
2. Pressione `Alt + F11` para abrir o VBE
3. Para cada arquivo `.bas`:
   - Clique com botão direito em qualquer módulo → **Remover módulo** (do módulo antigo)
   - Menu **Arquivo → Importar Arquivo** → selecione o `.bas` correspondente
4. Para `EsteLivro.bas`: **não importe como módulo** — abra o módulo `EsteLivro` existente e **cole o código** dentro dele
5. Crie os UserForms `frmMenu` e `frmRomaneio` no VBE e cole os respectivos códigos

---

## Etapa 1 — Correções críticas aplicadas

### Bug 1 — Módulo1: nome de planilha errado (CRÍTICO)
```vba
' ANTES (quebrava com Subscript out of range):
Set wsMov = Sheets("MOVIMENTACOES")

' DEPOIS:
' Removido — lógica centralizada em modGlobais.GravarMovimentacao()
' que usa a constante M_* e o nome correto "MOVIMENTACAO_ESTOQUE"
```

### Bug 2 — frmMovimentacao: variável Responsavel sombreada (CRÍTICO)
```vba
' ANTES (Responsavel sempre "" → nunca salvava):
Dim Responsavel As String   ' variável local ofuscando a global

' DEPOIS:
' Linha removida — usa diretamente a global de modGlobais
```

### Bug 3 — Estoque Total não atualizado após movimentação
```vba
' ANTES: só atualizava colunas 11 e 12, esquecia coluna 10
wsP.Cells(linhaProd, 11).Value = novoMaletas
wsP.Cells(linhaProd, 12).Value = novoDisp
' coluna 10 (estoqueTotal) nunca era atualizada!

' DEPOIS (em GravarMovimentacao no modGlobais):
wsP.Cells(linhaProd, P_ESTOQUE_DISPONIVEL).Value = novoDisp   ' col 12
wsP.Cells(linhaProd, P_ESTOQUE_MALETAS).Value    = novoMal    ' col 11
wsP.Cells(linhaProd, P_ESTOQUE_TOTAL).Value      = novoTot    ' col 10 ✓
```

### Bug 4 — GerarCodigoInterno sem verificação de unicidade
```vba
' ANTES:
GerarCodigoInterno = Format(Int(Rnd * 90000000) + 10000000, "00000000")
' Podia gerar código duplicado

' DEPOIS (GerarCodigoInternoUnico em Módulo1):
Do
    codigo = Format(CLng(Rnd * 90000000) + 10000000, "00000000")
    Set achou = ws.Columns(P_CODIGO).Find(codigo, LookAt:=xlWhole)
Loop While Not achou Is Nothing
```

### Etapa 3 — frmMenu adicionado
- Abre automaticamente após login bem-sucedido
- Botões para: Produtos, Movimentação, Revendedoras, Romaneios, Sair

---

## Etapa 5 — Constantes de colunas (documentação)

Todas as referências numéricas foram substituídas por constantes nomeadas em `modGlobais.bas`:

```vba
' Exemplo — antes:
wsP.Cells(linhaProd, 12).Value = novoDisp  ' o que é a coluna 12?

' Depois — autoexplicativo:
wsP.Cells(linhaProd, P_ESTOQUE_DISPONIVEL).Value = novoDisp
```

### Planilha PRODUTOS
| Constante | Col | Campo |
|---|---|---|
| P_CODIGO | 1 | Código Interno |
| P_DESCRICAO | 2 | Descrição |
| P_CUSTO_BRUTO | 5 | Custo Bruto |
| P_CUSTO_BANHO | 6 | Custo Banho |
| P_CUSTO_TOTAL | 7 | Custo Total |
| P_PRECO_VENDA | 8 | Preço de Venda |
| P_MARGEM_LUCRO | 9 | Margem (%) |
| P_ESTOQUE_TOTAL | 10 | Estoque Total |
| P_ESTOQUE_MALETAS | 11 | Estoque Em Maletas |
| P_ESTOQUE_DISPONIVEL | 12 | Estoque Disponível |
| P_STATUS | 18 | Status (ATIVO/INATIVO) |

### Planilha MOVIMENTACAO_ESTOQUE
| Constante | Col | Campo |
|---|---|---|
| M_IDM | 1 | ID da Movimentação |
| M_CODIGO | 2 | Código do Produto |
| M_TIPO | 12 | Tipo (ENTRADA/SAIDA) |
| M_ORIGEM | 13 | Origem |
| M_MALETA_ORIGEM | 14 | Maleta Origem |
| M_MOTIVO | 15 | Motivo |
| M_QUANTIDADE | 16 | Quantidade |
| M_RESPONSAVEL | 17 | Responsável |
| M_DESTINO | 18 | Destino |
| M_MALETA_DESTINO | 19 | Nome da Revendedora |
| M_DATA | 23 | Data/Hora |
| M_ID_ROMANEIO | 25 | ID do Romaneio |

### Planilha ROMANEIOS_ITENS
| Constante | Col | Campo |
|---|---|---|
| RI_ID_ROMANEIO | 1 | ID do Romaneio |
| RI_STATUS | 11 | PENDENTE/CONFERIDO/FINALIZADO |
| RI_QTD_RETORNOU | 14 | Qtd que voltou da maleta |
| RI_QTD_VENDIDA | 15 | Qtd considerada vendida |

---

## frmRomaneio — Fluxo implementado

```
[cmdNovo] → gera ID ROM-YYYYMMDD-NNN + data
    ↓
Seleciona revendedora (código ↔ combo sincronizados)
    ↓
Informa código produto → sistema carrega dados → informa qtd → [cmdAdicionar]
Repete para cada produto. Totalizadores atualizam em tempo real.
    ↓
[cmdSalvarRomaneio] → grava em ROMANEIOS_ITENS, status=PENDENTE
    ↓
Aba CONSULTAR ROMANEIO → filtros → [cmdConsultarRom]
Duplo clique PENDENTE → carrega em modo edição
Duplo clique outros status → carrega somente leitura
    ↓
[cmdConfirmarSaida] (só PENDENTE):
  → grava movimentação por produto: SAIDA/REVENDA/ESTOQUE DISPONIVEL→EM MALETA
  → status → CONFERIDO
  → gera PDF (layout IMP_ROMANEIOS) → salvo em /ROMANEIOS/{idRom}.pdf
    ↓
Aba CONFERÊNCIA:
Informa ID romaneio → [cmdBuscarRomaneio] → carrega grid (só status CONFERIDO)
Duplo clique em cada item → InputBox para informar qtd retornou
  Sistema calcula: qtd vendida = qtd enviada - qtd retornou
[cmdPreConferencia] → resumo para validação antes de confirmar
[cmdSalvarConferencia]:
  → qtd retornou: ENTRADA/REVENDA/EM MALETA→ESTOQUE DISPONIVEL
  → qtd vendida: SAIDA/VENDA/EM MALETA→VENDA
  → status → FINALIZADO
  → atualiza ROMANEIOS_ITENS (qtd_retornou, qtd_vendida, data_retorno)
```

---

## Próximos passos recomendados

1. **Importar os módulos** no VBE e criar os UserForms `frmMenu` e `frmRomaneio` com os controles listados nos comentários de cada arquivo
2. **Adicionar coluna M_ID_ROMANEIO (col 25)** na aba MOVIMENTACAO_ESTOQUE se ainda não existir
3. **Adicionar colunas RI_QTD_RETORNOU (14), RI_QTD_VENDIDA (15), RI_DATA_RETORNO (16)** na aba ROMANEIOS_ITENS
4. **Criar UserForm frmMenu** com os controles: `lblUsuario`, `cmdProdutos`, `cmdMovimentacao`, `cmdRevendedoras`, `cmdRomaneio`, `cmdSair`
5. Verificar células do layout IMP_ROMANEIOS (B2=ID, B3=Revendedora, B4=Data, B5=Responsável, A9:F=itens) e ajustar em `GerarPDFRomaneio` se necessário
