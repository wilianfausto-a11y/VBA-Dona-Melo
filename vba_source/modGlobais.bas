Attribute VB_Name = "modGlobais"
Option Explicit

' ============================================================
' VARIÁVEIS GLOBAIS DE SESSÃO
' ============================================================
Public Responsavel As String   ' Nome do usuário logado

' ============================================================
' MAPEAMENTO DE COLUNAS — PRODUTOS (aba PRODUTOS)
' ============================================================
Public Const P_CODIGO               As Long = 1   ' Código Interno
Public Const P_DESCRICAO            As Long = 2   ' Descrição
Public Const P_CATEGORIA            As Long = 3   ' Categoria
Public Const P_SUBCATEGORIA         As Long = 4   ' Subcategoria
Public Const P_CUSTO_BRUTO          As Long = 5   ' Custo Bruto (sem banho)
Public Const P_CUSTO_BANHO          As Long = 6   ' Custo Banho galvânico
Public Const P_CUSTO_TOTAL          As Long = 7   ' Custo Total (bruto + banho)
Public Const P_PRECO_VENDA          As Long = 8   ' Preço de Venda
Public Const P_MARGEM_LUCRO         As Long = 9   ' Margem de Lucro (%)
Public Const P_ESTOQUE_TOTAL        As Long = 10  ' Estoque Total (disp + maletas)
Public Const P_ESTOQUE_MALETAS      As Long = 11  ' Estoque Em Maletas
Public Const P_ESTOQUE_DISPONIVEL   As Long = 12  ' Estoque Disponível
Public Const P_TOTAL_VENDIDO        As Long = 13  ' Total de Unidades Vendidas
Public Const P_VAL_TOTAL_ESTOQUE    As Long = 14  ' Valor Total Estoque (venda)
Public Const P_VAL_TOTAL_MALETAS    As Long = 15  ' Valor Total Em Maletas (venda)
Public Const P_VAL_TOTAL_GERAL      As Long = 16  ' Valor Total Geral (venda)
Public Const P_VAL_TOTAL_CUSTO      As Long = 17  ' Valor Total Custo
Public Const P_STATUS               As Long = 18  ' Status (ATIVO / INATIVO)
Public Const P_MARCA                As Long = 19  ' Marca
Public Const P_FORNECEDOR           As Long = 20  ' Fornecedor
Public Const P_COD_FORNECEDOR       As Long = 21  ' Código do Fornecedor
Public Const P_VAL_EST_DISPONIVEL   As Long = 22  ' Valor Estoque Disponível (venda)
Public Const P_IMAGEM               As Long = 23  ' Caminho/referência de imagem
Public Const P_DATA_ALT             As Long = 24  ' Data da última alteração

' ============================================================
' MAPEAMENTO DE COLUNAS — MOVIMENTACAO_ESTOQUE
' ============================================================
Public Const M_IDM                  As Long = 1   ' ID da Movimentação
Public Const M_CODIGO               As Long = 2   ' Código do Produto
Public Const M_DESCRICAO            As Long = 3   ' Descrição do Produto
Public Const M_CATEGORIA            As Long = 4   ' Categoria
Public Const M_SUBCATEGORIA         As Long = 5   ' Subcategoria
Public Const M_EST_DISP_ANTES       As Long = 6   ' Estoque Disponível Antes
Public Const M_EST_MAL_ANTES        As Long = 7   ' Estoque Maletas Antes
Public Const M_EST_TOT_ANTES        As Long = 8   ' Estoque Total Antes
Public Const M_MARCA                As Long = 9   ' Marca
Public Const M_FORNECEDOR           As Long = 10  ' Fornecedor
Public Const M_COD_FORNECEDOR       As Long = 11  ' Código Fornecedor
Public Const M_TIPO                 As Long = 12  ' Tipo (ENTRADA / SAIDA)
Public Const M_ORIGEM               As Long = 13  ' Origem do estoque
Public Const M_MALETA_ORIGEM        As Long = 14  ' Maleta de Origem
Public Const M_MOTIVO               As Long = 15  ' Motivo
Public Const M_QUANTIDADE           As Long = 16  ' Quantidade movimentada
Public Const M_RESPONSAVEL          As Long = 17  ' Responsável
Public Const M_DESTINO              As Long = 18  ' Destino do estoque
Public Const M_MALETA_DESTINO       As Long = 19  ' Maleta de Destino (nome revendedora)
Public Const M_EST_MAL_DEPOIS       As Long = 20  ' Estoque Maletas Depois
Public Const M_EST_DISP_DEPOIS      As Long = 21  ' Estoque Disponível Depois
Public Const M_EST_TOT_DEPOIS       As Long = 22  ' Estoque Total Depois
Public Const M_DATA                 As Long = 23  ' Data/Hora da movimentação
Public Const M_OBSERVACAO           As Long = 24  ' Observação
Public Const M_ID_ROMANEIO          As Long = 25  ' ID do Romaneio vinculado

' ============================================================
' MAPEAMENTO DE COLUNAS — REVENDAS
' ============================================================
Public Const R_CODIGO               As Long = 1   ' Código da Revendedora
Public Const R_CPF                  As Long = 2   ' CPF
Public Const R_NOME                 As Long = 3   ' Nome
Public Const R_TELEFONE             As Long = 4   ' Telefone
Public Const R_EMAIL                As Long = 5   ' E-mail
Public Const R_WHATSAPP             As Long = 6   ' WhatsApp
Public Const R_RUA                  As Long = 7   ' Rua
Public Const R_NUMERO               As Long = 8   ' Número
Public Const R_BAIRRO               As Long = 9   ' Bairro
Public Const R_CIDADE               As Long = 10  ' Cidade
Public Const R_ESTADO               As Long = 11  ' Estado (UF)
Public Const R_CEP                  As Long = 12  ' CEP
Public Const R_STATUS               As Long = 13  ' Status (ATIVO / INATIVO)
Public Const R_DATA_CAD             As Long = 14  ' Data de Cadastro

' ============================================================
' MAPEAMENTO DE COLUNAS — ROMANEIOS_ITENS
' ============================================================
Public Const RI_ID_ROMANEIO         As Long = 1   ' ID do Romaneio
Public Const RI_DATA                As Long = 2   ' Data de criação
Public Const RI_ID_REVENDA          As Long = 3   ' Código da Revendedora
Public Const RI_NOME_REVENDEDORA    As Long = 4   ' Nome da Revendedora
Public Const RI_CODIGO_PRODUTO      As Long = 5   ' Código do Produto
Public Const RI_DESCRICAO           As Long = 6   ' Descrição do Produto
Public Const RI_CATEGORIA           As Long = 7   ' Categoria
Public Const RI_QUANTIDADE          As Long = 8   ' Quantidade Enviada
Public Const RI_PRECO_VENDA         As Long = 9   ' Preço de Venda unitário
Public Const RI_TOTAL               As Long = 10  ' Total (qtd × preço)
Public Const RI_STATUS              As Long = 11  ' Status (PENDENTE/CONFERIDO/FINALIZADO)
Public Const RI_RESPONSAVEL         As Long = 12  ' Responsável pelo romaneio
Public Const RI_DATA_CONFERENCIA    As Long = 13  ' Data da conferência de saída
Public Const RI_QTD_RETORNOU        As Long = 14  ' Quantidade que retornou
Public Const RI_QTD_VENDIDA         As Long = 15  ' Quantidade considerada vendida
Public Const RI_DATA_RETORNO        As Long = 16  ' Data do retorno/finalização

' ============================================================
' FUNÇÕES UTILITÁRIAS GLOBAIS
' ============================================================

' Valida CPF com dígito verificador
Public Function ValidarCPF(ByVal cpf As String) As Boolean
    Dim nums As String
    Dim soma As Long
    Dim resto As Long
    Dim d1 As Long, d2 As Long
    Dim i As Long
    
    nums = ""
    Dim c As String
    Dim j As Long
    For j = 1 To Len(cpf)
        c = Mid(cpf, j, 1)
        If c >= "0" And c <= "9" Then nums = nums & c
    Next j
    
    If Len(nums) <> 11 Then Exit Function
    
    ' Rejeita sequências inválidas (111.111.111-11 etc.)
    If nums = String(11, Left(nums, 1)) Then Exit Function
    
    ' Primeiro dígito verificador
    soma = 0
    For i = 1 To 9
        soma = soma + CLng(Mid(nums, i, 1)) * (11 - i)
    Next i
    resto = soma Mod 11
    d1 = IIf(resto < 2, 0, 11 - resto)
    
    If d1 <> CLng(Mid(nums, 10, 1)) Then Exit Function
    
    ' Segundo dígito verificador
    soma = 0
    For i = 1 To 10
        soma = soma + CLng(Mid(nums, i, 1)) * (12 - i)
    Next i
    resto = soma Mod 11
    d2 = IIf(resto < 2, 0, 11 - resto)
    
    If d2 <> CLng(Mid(nums, 11, 1)) Then Exit Function
    
    ValidarCPF = True
End Function

' Próximo IDM disponível na aba MOVIMENTACAO_ESTOQUE
Public Function ProximoIDM() As Long
    Dim ws As Worksheet
    Dim ult As Long
    Dim i As Long
    Dim maior As Long
    Dim v As Variant
    
    Set ws = ThisWorkbook.Worksheets("MOVIMENTACAO_ESTOQUE")
    ult = ws.Cells(ws.Rows.Count, M_IDM).End(xlUp).Row
    maior = 0
    
    For i = 2 To ult
        v = ws.Cells(i, M_IDM).Value
        If IsNumeric(v) Then
            If CLng(v) > maior Then maior = CLng(v)
        End If
    Next i
    
    ProximoIDM = maior + 1
End Function

' Grava uma movimentação de estoque e atualiza PRODUTOS
' Retorna True se bem-sucedido
Public Function GravarMovimentacao( _
    ByVal codigoProduto As String, _
    ByVal tipoMov As String, _
    ByVal motivo As String, _
    ByVal origem As String, _
    ByVal destino As String, _
    ByVal maletaOrigem As String, _
    ByVal maletaDestino As String, _
    ByVal quantidade As Long, _
    ByVal observacao As String, _
    Optional ByVal idRomaneio As String = "") As Boolean

    Dim wsP As Worksheet
    Dim wsM As Worksheet
    Dim linhaProd As Long
    Dim ultMov As Long
    
    Dim estDisp As Long, estMal As Long, estTot As Long
    Dim novoDisp As Long, novoMal As Long, novoTot As Long
    
    Set wsP = ThisWorkbook.Worksheets("PRODUTOS")
    Set wsM = ThisWorkbook.Worksheets("MOVIMENTACAO_ESTOQUE")
    
    ' Localiza produto
    Dim r As Range
    Set r = wsP.Columns(P_CODIGO).Find(codigoProduto, LookAt:=xlWhole)
    If r Is Nothing Then
        MsgBox "Produto '" & codigoProduto & "' não encontrado.", vbCritical
        Exit Function
    End If
    linhaProd = r.Row
    
    estDisp = CLng(wsP.Cells(linhaProd, P_ESTOQUE_DISPONIVEL).Value)
    estMal  = CLng(wsP.Cells(linhaProd, P_ESTOQUE_MALETAS).Value)
    estTot  = CLng(wsP.Cells(linhaProd, P_ESTOQUE_TOTAL).Value)
    
    novoDisp = estDisp
    novoMal  = estMal
    novoTot  = estTot
    
    ' Calcula novos estoques conforme tipo/motivo
    Select Case UCase(Trim(tipoMov))
    
        Case "ENTRADA"
            Select Case UCase(Trim(motivo))
                Case "REVENDA"          ' retorno de maleta
                    novoMal  = novoMal - quantidade
                    novoDisp = novoDisp + quantidade
                Case "COMPRA", "AJUSTE MANUAL", "BANHO"
                    novoDisp = novoDisp + quantidade
                Case Else
                    novoDisp = novoDisp + quantidade
            End Select
            novoTot = novoDisp + novoMal
            
        Case "SAIDA", "SAÍDA"
            Select Case UCase(Trim(motivo))
                Case "REVENDA"          ' envio para maleta
                    novoDisp = novoDisp - quantidade
                    novoMal  = novoMal  + quantidade
                Case "VENDA"
                    novoMal  = novoMal  - quantidade
                Case Else
                    novoDisp = novoDisp - quantidade
            End Select
            novoTot = novoDisp + novoMal
    End Select
    
    ' Validação de negativos
    If novoDisp < 0 Or novoMal < 0 Or novoTot < 0 Then
        MsgBox "Estoque insuficiente para o produto " & codigoProduto & ".", vbCritical
        Exit Function
    End If
    
    ' Grava movimentação
    ultMov = wsM.Cells(wsM.Rows.Count, M_IDM).End(xlUp).Row + 1
    
    With wsM
        .Cells(ultMov, M_IDM).Value            = ProximoIDM()
        .Cells(ultMov, M_CODIGO).Value          = codigoProduto
        .Cells(ultMov, M_DESCRICAO).Value       = wsP.Cells(linhaProd, P_DESCRICAO).Value
        .Cells(ultMov, M_CATEGORIA).Value       = wsP.Cells(linhaProd, P_CATEGORIA).Value
        .Cells(ultMov, M_SUBCATEGORIA).Value    = wsP.Cells(linhaProd, P_SUBCATEGORIA).Value
        .Cells(ultMov, M_EST_DISP_ANTES).Value  = estDisp
        .Cells(ultMov, M_EST_MAL_ANTES).Value   = estMal
        .Cells(ultMov, M_EST_TOT_ANTES).Value   = estTot
        .Cells(ultMov, M_MARCA).Value           = wsP.Cells(linhaProd, P_MARCA).Value
        .Cells(ultMov, M_FORNECEDOR).Value      = wsP.Cells(linhaProd, P_FORNECEDOR).Value
        .Cells(ultMov, M_COD_FORNECEDOR).Value  = wsP.Cells(linhaProd, P_COD_FORNECEDOR).Value
        .Cells(ultMov, M_TIPO).Value            = UCase(Trim(tipoMov))
        .Cells(ultMov, M_ORIGEM).Value          = UCase(Trim(origem))
        .Cells(ultMov, M_MALETA_ORIGEM).Value   = maletaOrigem
        .Cells(ultMov, M_MOTIVO).Value          = UCase(Trim(motivo))
        .Cells(ultMov, M_QUANTIDADE).Value      = quantidade
        .Cells(ultMov, M_RESPONSAVEL).Value     = Responsavel
        .Cells(ultMov, M_DESTINO).Value         = UCase(Trim(destino))
        .Cells(ultMov, M_MALETA_DESTINO).Value  = maletaDestino
        .Cells(ultMov, M_EST_MAL_DEPOIS).Value  = novoMal
        .Cells(ultMov, M_EST_DISP_DEPOIS).Value = novoDisp
        .Cells(ultMov, M_EST_TOT_DEPOIS).Value  = novoTot
        .Cells(ultMov, M_DATA).Value            = Now
        .Cells(ultMov, M_OBSERVACAO).Value      = observacao
        .Cells(ultMov, M_ID_ROMANEIO).Value     = idRomaneio
    End With
    
    ' Atualiza PRODUTOS
    wsP.Cells(linhaProd, P_ESTOQUE_DISPONIVEL).Value = novoDisp
    wsP.Cells(linhaProd, P_ESTOQUE_MALETAS).Value    = novoMal
    wsP.Cells(linhaProd, P_ESTOQUE_TOTAL).Value      = novoTot
    wsP.Cells(linhaProd, P_DATA_ALT).Value           = Now
    
    GravarMovimentacao = True
End Function
