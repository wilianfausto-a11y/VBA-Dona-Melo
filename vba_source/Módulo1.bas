Attribute VB_Name = "Módulo1"
Option Explicit

' ============================================================
' MÓDULO DE CÁLCULOS E UTILITÁRIOS DE PRODUTO
' Nota: Movimentações de estoque agora são feitas via
' modGlobais.GravarMovimentacao() que centraliza toda a lógica.
' ============================================================

' Recalcula margem e valores derivados de um produto
' Chamado após qualquer alteração de preços na aba PRODUTOS
Public Sub CalcularTotaisProduto(ByVal linha As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("PRODUTOS")
    
    Dim custoBruto  As Double: custoBruto  = ws.Cells(linha, P_CUSTO_BRUTO).Value
    Dim custoBanho  As Double: custoBanho  = ws.Cells(linha, P_CUSTO_BANHO).Value
    Dim custoTotal  As Double: custoTotal  = custoBruto + custoBanho
    Dim precoVenda  As Double: precoVenda  = ws.Cells(linha, P_PRECO_VENDA).Value
    Dim estTotal    As Long:   estTotal    = ws.Cells(linha, P_ESTOQUE_TOTAL).Value
    Dim estMaletas  As Long:   estMaletas  = ws.Cells(linha, P_ESTOQUE_MALETAS).Value
    Dim estDisp     As Long:   estDisp     = ws.Cells(linha, P_ESTOQUE_DISPONIVEL).Value
    
    ' Custo total
    ws.Cells(linha, P_CUSTO_TOTAL).Value = custoTotal
    
    ' Margem de lucro
    If custoTotal > 0 Then
        ws.Cells(linha, P_MARGEM_LUCRO).Value = (precoVenda - custoTotal) / custoTotal
    Else
        ws.Cells(linha, P_MARGEM_LUCRO).Value = 0
    End If
    
    ' Valores monetários de estoque
    ws.Cells(linha, P_VAL_TOTAL_ESTOQUE).Value   = estTotal   * precoVenda
    ws.Cells(linha, P_VAL_TOTAL_MALETAS).Value   = estMaletas * precoVenda
    ws.Cells(linha, P_VAL_TOTAL_CUSTO).Value     = estTotal   * custoTotal
    ws.Cells(linha, P_VAL_EST_DISPONIVEL).Value  = estDisp    * precoVenda
    ws.Cells(linha, P_VAL_TOTAL_GERAL).Value     = estTotal   * precoVenda
End Sub

' Gera código interno único de 8 dígitos para produto
Public Function GerarCodigoInternoUnico() As String
    Dim ws As Worksheet
    Dim codigo As String
    Dim achou As Range
    
    Set ws = ThisWorkbook.Worksheets("PRODUTOS")
    
    Randomize
    Do
        codigo = Format(CLng(Rnd * 90000000) + 10000000, "00000000")
        Set achou = ws.Columns(P_CODIGO).Find(What:=codigo, LookAt:=xlWhole)
    Loop While Not achou Is Nothing
    
    GerarCodigoInternoUnico = codigo
End Function

' Próximo ID de Romaneio (formato ROM-YYYYMMDD-NNN)
Public Function GerarIDRomaneio() As String
    Dim ws As Worksheet
    Dim ult As Long
    Dim i As Long
    Dim maior As Long
    Dim partes() As String
    Dim sufixo As Long
    Dim prefixo As String
    
    prefixo = "ROM-" & Format(Date, "YYYYMMDD") & "-"
    
    Set ws = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    maior = 0
    
    For i = 2 To ult
        Dim val As String
        val = CStr(ws.Cells(i, RI_ID_ROMANEIO).Value)
        If Left(val, Len(prefixo)) = prefixo Then
            On Error Resume Next
            sufixo = CLng(Mid(val, Len(prefixo) + 1))
            On Error GoTo 0
            If sufixo > maior Then maior = sufixo
        End If
    Next i
    
    GerarIDRomaneio = prefixo & Format(maior + 1, "000")
End Function
