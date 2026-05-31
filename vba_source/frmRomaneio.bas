Attribute VB_Name = "frmRomaneio"
Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================
' frmRomaneio — Gestão completa de romaneios de maletas
'
' ABAS (MultiPage1):
'   0 = ROMANEIO           (criação e edição)
'   1 = CONSULTAR ROMANEIO (filtros + ListView)
'   2 = CONFERÊNCIA        (retorno de maletas)
'
' CONTROLES — ABA ROMANEIO (pgRomaneio):
'   txtIDRomaneio         — ID gerado (somente leitura)
'   txtDataRomaneio       — Data (somente leitura)
'   txtCodRevendedora     — Código da revendedora
'   cboNomeRevendedora    — Combo com nomes
'   txtCodigoProduto      — Código do produto a adicionar
'   txtDescricao          — Descrição (auto)
'   txtCategoria          — Categoria (auto)
'   txtPrecoVenda         — Preço de venda (auto)
'   txtQtdAdicionar       — Quantidade a adicionar
'   cmdConsultarProduto   — Busca produto pelo código
'   cmdAdicionar          — Adiciona item ao ListView
'   cmdRemoverItem        — Remove item selecionado
'   cmdNovo               — Inicia novo romaneio
'   cmdSalvarRomaneio     — Salva romaneio como PENDENTE
'   cmdLimparRomaneio     — Limpa tudo
'   cmdMenuRomaneio       — Voltar ao menu
'   lvItens               — ListView com itens do romaneio
'   lblTotalItens         — Label total de itens
'   lblTotalValor         — Label valor total
'   lblStatusRomaneio     — Label com status atual
'
' CONTROLES — ABA CONSULTAR (pgConsultar):
'   cboFiltroStatus       — Filtro por status
'   cboFiltroRevenda      — Filtro por revendedora
'   txtFiltroID           — Filtro por ID do romaneio
'   cmdConsultarRom       — Executa consulta
'   cmdLimparFiltros      — Limpa filtros
'   lvRomaneios           — ListView com romaneios
'   cmdConfirmarSaida     — Confirma romaneio (PENDENTE→CONFERIDO + PDF + movimentação)
'
' CONTROLES — ABA CONFERÊNCIA (pgConferencia):
'   txtRomaneioConf       — ID do romaneio para conferir
'   cmdBuscarRomaneio     — Busca itens do romaneio
'   lvConferencia         — ListView com itens + campo qtd retornou
'   cmdPreConferencia     — Pré-conferência (valida antes de salvar)
'   cmdSalvarConferencia  — Salva conferência definitiva
'   cmdLimparConferencia  — Limpa aba
'   lblInfoConferencia    — Label com revendedora/data/status
' ============================================================

' --- ESTADO INTERNO ---
Private mModoRomaneio    As Integer   ' 0=Novo, 1=Edicao, 2=SoLeitura
Private mIDRomaneioAtual As String    ' ID em edição
Private mLinhaRevAtual   As Long      ' Linha da revendedora selecionada
Private mConfIDs()       As String    ' IDs únicos de romaneios no lvRomaneios
Private Const COR_BLOQUEADO As Long = &HE0E0E0
Private Const COR_LIBERADO  As Long = vbWhite

' ============================================================
' INICIALIZAÇÃO
' ============================================================
Private Sub UserForm_Initialize()
    ConfigurarListViewItens
    ConfigurarListViewConsulta
    ConfigurarListViewConferencia
    CarregarRevendedorasCombo
    CarregarFiltroStatus
    LimparAbaRomaneio
    LimparFiltros
    LimparConferencia
    MultiPage1.Value = 0
End Sub

' ============================================================
' ABA 0 — ROMANEIO (Criação / Edição)
' ============================================================

Private Sub ConfigurarListViewItens()
    With lvItens
        .View = 3                   ' lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Código",      70
        .ColumnHeaders.Add , , "Descrição",   160
        .ColumnHeaders.Add , , "Categoria",   80
        .ColumnHeaders.Add , , "Qtd",         45
        .ColumnHeaders.Add , , "Preço Unit.", 70
        .ColumnHeaders.Add , , "Total",       80
    End With
End Sub

Private Sub CarregarRevendedorasCombo()
    Dim ws As Worksheet
    Dim ult As Long, i As Long
    
    Set ws = ThisWorkbook.Worksheets("REVENDAS")
    ult = ws.Cells(ws.Rows.Count, R_CODIGO).End(xlUp).Row
    
    cboNomeRevendedora.Clear
    If Not Not cboFiltroRevenda Is Nothing Then cboFiltroRevenda.Clear
    
    For i = 2 To ult
        If UCase(Trim(ws.Cells(i, R_STATUS).Value)) = "ATIVO" Then
            cboNomeRevendedora.AddItem Trim(ws.Cells(i, R_NOME).Value)
            On Error Resume Next
            cboFiltroRevenda.AddItem Trim(ws.Cells(i, R_NOME).Value)
            On Error GoTo 0
        End If
    Next i
End Sub

' Seleção pelo combo de nome → preenche código
Private Sub cboNomeRevendedora_Change()
    Dim ws As Worksheet
    Dim ult As Long, i As Long
    Dim nome As String
    
    nome = Trim(cboNomeRevendedora.Value)
    If nome = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("REVENDAS")
    ult = ws.Cells(ws.Rows.Count, R_CODIGO).End(xlUp).Row
    
    For i = 2 To ult
        If UCase(Trim(ws.Cells(i, R_NOME).Value)) = UCase(nome) Then
            txtCodRevendedora.Value = ws.Cells(i, R_CODIGO).Value
            mLinhaRevAtual = i
            Exit Sub
        End If
    Next i
End Sub

' Seleção pelo código → preenche combo
Private Sub txtCodRevendedora_AfterUpdate()
    Dim ws As Worksheet
    Dim ult As Long, i As Long
    Dim cod As String
    
    cod = Trim(txtCodRevendedora.Value)
    If cod = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("REVENDAS")
    ult = ws.Cells(ws.Rows.Count, R_CODIGO).End(xlUp).Row
    
    For i = 2 To ult
        If Trim(ws.Cells(i, R_CODIGO).Value) = cod Then
            cboNomeRevendedora.Value = ws.Cells(i, R_NOME).Value
            mLinhaRevAtual = i
            Exit Sub
        End If
    Next i
    
    MsgBox "Revendedora não encontrada.", vbExclamation
    txtCodRevendedora.Value = ""
    cboNomeRevendedora.Value = ""
    mLinhaRevAtual = 0
End Sub

' Consulta produto pelo código
Private Sub cmdConsultarProduto_Click()
    Dim ws As Worksheet
    Dim r As Range
    Dim cod As String
    
    cod = Trim(txtCodigoProduto.Value)
    If cod = "" Then
        MsgBox "Informe o código do produto.", vbExclamation
        txtCodigoProduto.SetFocus
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("PRODUTOS")
    Set r = ws.Columns(P_CODIGO).Find(cod, LookAt:=xlWhole)
    
    If r Is Nothing Then
        MsgBox "Produto não encontrado.", vbExclamation
        txtCodigoProduto.Value = ""
        txtCodigoProduto.SetFocus
        Exit Sub
    End If
    
    txtDescricao.Value  = ws.Cells(r.Row, P_DESCRICAO).Value
    txtCategoria.Value  = ws.Cells(r.Row, P_CATEGORIA).Value
    txtPrecoVenda.Value = Format(ws.Cells(r.Row, P_PRECO_VENDA).Value, "R$ #,##0.00")
    txtQtdAdicionar.SetFocus
End Sub

Private Sub txtCodigoProduto_AfterUpdate()
    If Trim(txtCodigoProduto.Value) <> "" Then cmdConsultarProduto_Click
End Sub

' Adiciona item ao ListView
Private Sub cmdAdicionar_Click()
    Dim cod     As String
    Dim desc    As String
    Dim cat     As String
    Dim precoS  As String
    Dim preco   As Double
    Dim qtd     As Long
    Dim total   As Double
    Dim item    As ListItem
    
    ' Validações
    If Trim(txtCodRevendedora.Value) = "" Or Trim(cboNomeRevendedora.Value) = "" Then
        MsgBox "Selecione a revendedora antes de adicionar itens.", vbExclamation
        Exit Sub
    End If
    
    cod = Trim(txtCodigoProduto.Value)
    If cod = "" Then
        MsgBox "Informe o código do produto.", vbExclamation
        txtCodigoProduto.SetFocus
        Exit Sub
    End If
    
    If txtDescricao.Value = "" Then
        MsgBox "Consulte o produto antes de adicionar.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtQtdAdicionar.Value) Or CLng(txtQtdAdicionar.Value) <= 0 Then
        MsgBox "Informe uma quantidade válida maior que zero.", vbExclamation
        txtQtdAdicionar.SetFocus
        Exit Sub
    End If
    
    qtd = CLng(txtQtdAdicionar.Value)
    
    ' Verifica estoque disponível
    Dim ws As Worksheet
    Dim r As Range
    Set ws = ThisWorkbook.Worksheets("PRODUTOS")
    Set r = ws.Columns(P_CODIGO).Find(cod, LookAt:=xlWhole)
    
    If r Is Nothing Then
        MsgBox "Produto não encontrado em PRODUTOS.", vbExclamation
        Exit Sub
    End If
    
    Dim estDisp As Long
    estDisp = CLng(ws.Cells(r.Row, P_ESTOQUE_DISPONIVEL).Value)
    
    If qtd > estDisp Then
        MsgBox "Quantidade solicitada (" & qtd & ") superior ao estoque disponível (" & estDisp & ").", vbExclamation
        txtQtdAdicionar.SetFocus
        Exit Sub
    End If
    
    ' Verifica se produto já está na lista e soma
    Dim i As Long
    For i = 1 To lvItens.ListItems.Count
        If lvItens.ListItems(i).Text = cod Then
            Dim qtdExist As Long
            qtdExist = CLng(lvItens.ListItems(i).SubItems(3))
            Dim novaQtd As Long
            novaQtd = qtdExist + qtd
            
            If novaQtd > estDisp Then
                MsgBox "Quantidade total (" & novaQtd & ") supera o estoque disponível (" & estDisp & ").", vbExclamation
                Exit Sub
            End If
            
            preco = CDbl(Replace(Replace(lvItens.ListItems(i).SubItems(4), "R$", ""), " ", ""))
            lvItens.ListItems(i).SubItems(3) = novaQtd
            lvItens.ListItems(i).SubItems(5) = Format(novaQtd * preco, "R$ #,##0.00")
            AtualizarTotalizadores
            LimparCamposProduto
            Exit Sub
        End If
    Next i
    
    ' Novo item
    desc   = txtDescricao.Value
    cat    = txtCategoria.Value
    precoS = Replace(Replace(txtPrecoVenda.Value, "R$", ""), " ", "")
    preco  = CDbl(precoS)
    total  = qtd * preco
    
    Set item = lvItens.ListItems.Add(Text:=cod)
    item.SubItems(1) = desc
    item.SubItems(2) = cat
    item.SubItems(3) = CStr(qtd)
    item.SubItems(4) = Format(preco, "R$ #,##0.00")
    item.SubItems(5) = Format(total, "R$ #,##0.00")
    
    AtualizarTotalizadores
    LimparCamposProduto
    txtCodigoProduto.SetFocus
End Sub

' Remove item selecionado do ListView
Private Sub cmdRemoverItem_Click()
    If lvItens.SelectedItem Is Nothing Then
        MsgBox "Selecione um item para remover.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Remover item '" & lvItens.SelectedItem.SubItems(1) & "'?", vbYesNo) = vbNo Then Exit Sub
    
    lvItens.ListItems.Remove lvItens.SelectedItem.Index
    AtualizarTotalizadores
End Sub

' Atualiza totalizadores de quantidade e valor
Private Sub AtualizarTotalizadores()
    Dim totalQtd    As Long
    Dim totalValor  As Double
    Dim i As Long
    Dim precoS As String
    
    For i = 1 To lvItens.ListItems.Count
        totalQtd = totalQtd + CLng(lvItens.ListItems(i).SubItems(3))
        precoS = Replace(Replace(lvItens.ListItems(i).SubItems(5), "R$", ""), " ", "")
        On Error Resume Next
        totalValor = totalValor + CDbl(precoS)
        On Error GoTo 0
    Next i
    
    lblTotalItens.Caption = "Total de itens: " & totalQtd & " unidades | " & lvItens.ListItems.Count & " produto(s)"
    lblTotalValor.Caption = "Valor total: " & Format(totalValor, "R$ #,##0.00")
End Sub

' Novo romaneio
Private Sub cmdNovo_Click()
    LimparAbaRomaneio
    
    txtIDRomaneio.Value    = GerarIDRomaneio()
    txtDataRomaneio.Value  = Format(Now, "dd/mm/yyyy hh:mm")
    mIDRomaneioAtual       = txtIDRomaneio.Value
    mModoRomaneio          = 0
    lblStatusRomaneio.Caption = "Status: NOVO"
    
    txtCodRevendedora.SetFocus
End Sub

' Salva romaneio como PENDENTE
Private Sub cmdSalvarRomaneio_Click()
    ' Validações
    If Trim(Responsavel) = "" Then
        MsgBox "Nenhum usuário logado.", vbCritical
        Exit Sub
    End If
    
    If Trim(txtIDRomaneio.Value) = "" Then
        MsgBox "Clique em NOVO para iniciar um romaneio.", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCodRevendedora.Value) = "" Then
        MsgBox "Selecione a revendedora.", vbExclamation
        Exit Sub
    End If
    
    If lvItens.ListItems.Count = 0 Then
        MsgBox "Adicione pelo menos um item ao romaneio.", vbExclamation
        Exit Sub
    End If
    
    Dim idRom        As String
    Dim codRev       As String
    Dim nomeRev      As String
    Dim ws           As Worksheet
    Dim i As Long
    
    idRom   = Trim(txtIDRomaneio.Value)
    codRev  = Trim(txtCodRevendedora.Value)
    nomeRev = Trim(cboNomeRevendedora.Value)
    
    Set ws = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    
    Application.ScreenUpdating = False
    On Error GoTo ErroSalvar
    
    ' Se estiver editando (modo 1), remove registros anteriores PENDENTES
    If mModoRomaneio = 1 Then
        Call RemoverItensPendentesRomaneio(idRom, ws)
    End If
    
    ' Grava cada item
    For i = 1 To lvItens.ListItems.Count
        Dim ultL As Long
        ultL = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row + 1
        
        Dim precoS As String
        precoS = Replace(Replace(lvItens.ListItems(i).SubItems(4), "R$", ""), " ", "")
        Dim totalS As String
        totalS = Replace(Replace(lvItens.ListItems(i).SubItems(5), "R$", ""), " ", "")
        
        With ws
            .Cells(ultL, RI_ID_ROMANEIO).Value      = idRom
            .Cells(ultL, RI_DATA).Value             = Now
            .Cells(ultL, RI_ID_REVENDA).Value       = codRev
            .Cells(ultL, RI_NOME_REVENDEDORA).Value = nomeRev
            .Cells(ultL, RI_CODIGO_PRODUTO).Value   = lvItens.ListItems(i).Text
            .Cells(ultL, RI_DESCRICAO).Value        = lvItens.ListItems(i).SubItems(1)
            .Cells(ultL, RI_CATEGORIA).Value        = lvItens.ListItems(i).SubItems(2)
            .Cells(ultL, RI_QUANTIDADE).Value       = CLng(lvItens.ListItems(i).SubItems(3))
            .Cells(ultL, RI_PRECO_VENDA).Value      = CDbl(precoS)
            .Cells(ultL, RI_TOTAL).Value            = CDbl(totalS)
            .Cells(ultL, RI_STATUS).Value           = "PENDENTE"
            .Cells(ultL, RI_RESPONSAVEL).Value      = Responsavel
        End With
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Romaneio " & idRom & " salvo com sucesso!" & vbNewLine & "Status: PENDENTE", vbInformation
    LimparAbaRomaneio
    Exit Sub

ErroSalvar:
    Application.ScreenUpdating = True
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
End Sub

' Remove itens PENDENTES de um romaneio (para reedição)
Private Sub RemoverItensPendentesRomaneio(ByVal idRom As String, ByVal ws As Worksheet)
    Dim ult As Long, i As Long
    ult = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    
    For i = ult To 2 Step -1
        If CStr(ws.Cells(i, RI_ID_ROMANEIO).Value) = idRom And _
           UCase(Trim(ws.Cells(i, RI_STATUS).Value)) = "PENDENTE" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Private Sub cmdLimparRomaneio_Click()
    If MsgBox("Limpar todos os dados do romaneio atual?", vbYesNo) = vbYes Then
        LimparAbaRomaneio
    End If
End Sub

Private Sub cmdMenuRomaneio_Click()
    Unload Me
End Sub

Private Sub LimparAbaRomaneio()
    txtIDRomaneio.Value    = ""
    txtDataRomaneio.Value  = ""
    txtCodRevendedora.Value = ""
    cboNomeRevendedora.Value = ""
    LimparCamposProduto
    lvItens.ListItems.Clear
    lblTotalItens.Caption      = "Total de itens: 0 unidades | 0 produto(s)"
    lblTotalValor.Caption      = "Valor total: R$ 0,00"
    lblStatusRomaneio.Caption  = "Status: —"
    mIDRomaneioAtual           = ""
    mLinhaRevAtual             = 0
    mModoRomaneio              = 0
End Sub

Private Sub LimparCamposProduto()
    txtCodigoProduto.Value = ""
    txtDescricao.Value     = ""
    txtCategoria.Value     = ""
    txtPrecoVenda.Value    = ""
    txtQtdAdicionar.Value  = ""
End Sub

' ============================================================
' ABA 1 — CONSULTAR ROMANEIO
' ============================================================

Private Sub ConfigurarListViewConsulta()
    With lvRomaneios
        .View = 3
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID Romaneio",    100
        .ColumnHeaders.Add , , "Data",            80
        .ColumnHeaders.Add , , "Revendedora",    150
        .ColumnHeaders.Add , , "Itens",           45
        .ColumnHeaders.Add , , "Valor Total",     90
        .ColumnHeaders.Add , , "Status",          80
        .ColumnHeaders.Add , , "Responsável",     90
    End With
End Sub

Private Sub CarregarFiltroStatus()
    cboFiltroStatus.Clear
    cboFiltroStatus.AddItem ""
    cboFiltroStatus.AddItem "PENDENTE"
    cboFiltroStatus.AddItem "CONFERIDO"
    cboFiltroStatus.AddItem "FINALIZADO"
End Sub

Private Sub cmdConsultarRom_Click()
    Dim ws       As Worksheet
    Dim ult As Long, i As Long
    Dim filtStatus As String
    Dim filtRev    As String
    Dim filtID     As String
    
    filtStatus = UCase(Trim(cboFiltroStatus.Value))
    filtRev    = UCase(Trim(cboFiltroRevenda.Value))
    filtID     = UCase(Trim(txtFiltroID.Value))
    
    Set ws = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    
    lvRomaneios.ListItems.Clear
    
    ' Agrega por ID de romaneio usando Dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To ult
        Dim idRom   As String
        Dim status  As String
        Dim nomeRev As String
        
        idRom   = CStr(ws.Cells(i, RI_ID_ROMANEIO).Value)
        status  = UCase(Trim(ws.Cells(i, RI_STATUS).Value))
        nomeRev = UCase(Trim(ws.Cells(i, RI_NOME_REVENDEDORA).Value))
        
        If filtStatus <> "" And status <> filtStatus Then GoTo ProxLinha
        If filtRev <> "" And nomeRev <> filtRev Then GoTo ProxLinha
        If filtID <> "" And UCase(idRom) <> filtID Then GoTo ProxLinha
        
        If Not dict.Exists(idRom) Then
            ' [data, nomeRev, qtdItens, valorTotal, status, responsavel]
            dict(idRom) = Array( _
                ws.Cells(i, RI_DATA).Value, _
                ws.Cells(i, RI_NOME_REVENDEDORA).Value, _
                0, 0, _
                ws.Cells(i, RI_STATUS).Value, _
                ws.Cells(i, RI_RESPONSAVEL).Value)
        End If
        
        Dim arr As Variant
        arr = dict(idRom)
        arr(2) = arr(2) + CLng(ws.Cells(i, RI_QUANTIDADE).Value)
        arr(3) = arr(3) + CDbl(ws.Cells(i, RI_TOTAL).Value)
        dict(idRom) = arr
        
ProxLinha:
    Next i
    
    ' Preenche ListView
    Dim k As Variant
    For Each k In dict.Keys
        arr = dict(k)
        Dim itm As ListItem
        Set itm = lvRomaneios.ListItems.Add(Text:=CStr(k))
        itm.SubItems(1) = Format(arr(0), "dd/mm/yyyy")
        itm.SubItems(2) = arr(1)
        itm.SubItems(3) = CStr(arr(2))
        itm.SubItems(4) = Format(arr(3), "R$ #,##0.00")
        itm.SubItems(5) = arr(4)
        itm.SubItems(6) = arr(5)
    Next k
End Sub

Private Sub cmdLimparFiltros_Click()
    LimparFiltros
End Sub

Private Sub LimparFiltros()
    cboFiltroStatus.Value = ""
    On Error Resume Next
    cboFiltroRevenda.Value = ""
    On Error GoTo 0
    txtFiltroID.Value = ""
    lvRomaneios.ListItems.Clear
End Sub

' Duplo clique na consulta → carrega na aba ROMANEIO
Private Sub lvRomaneios_DblClick()
    If lvRomaneios.SelectedItem Is Nothing Then Exit Sub
    
    Dim idRom  As String
    Dim status As String
    
    idRom  = lvRomaneios.SelectedItem.Text
    status = UCase(Trim(lvRomaneios.SelectedItem.SubItems(5)))
    
    ' Carrega dados do romaneio na aba ROMANEIO
    CarregarRomaneioNaAba idRom, status
    MultiPage1.Value = 0
End Sub

Private Sub CarregarRomaneioNaAba(ByVal idRom As String, ByVal status As String)
    Dim ws As Worksheet
    Dim ult As Long, i As Long
    
    Set ws = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    
    LimparAbaRomaneio
    
    mIDRomaneioAtual = idRom
    txtIDRomaneio.Value   = idRom
    lblStatusRomaneio.Caption = "Status: " & status
    
    ' Bloqueia edição se não for PENDENTE
    If status <> "PENDENTE" Then
        mModoRomaneio = 2
        txtCodRevendedora.Locked = True
        cboNomeRevendedora.Enabled = False
        txtCodigoProduto.Locked = True
        txtQtdAdicionar.Locked = True
        cmdAdicionar.Enabled = False
        cmdRemoverItem.Enabled = False
        cmdSalvarRomaneio.Enabled = False
    Else
        mModoRomaneio = 1
    End If
    
    For i = 2 To ult
        If CStr(ws.Cells(i, RI_ID_ROMANEIO).Value) = idRom Then
            ' Cabeçalho: pega da primeira linha encontrada
            If txtDataRomaneio.Value = "" Then
                txtDataRomaneio.Value  = Format(ws.Cells(i, RI_DATA).Value, "dd/mm/yyyy")
                txtCodRevendedora.Value = ws.Cells(i, RI_ID_REVENDA).Value
                cboNomeRevendedora.Value = ws.Cells(i, RI_NOME_REVENDEDORA).Value
            End If
            
            ' Adiciona item ao ListView
            Dim itm As ListItem
            Dim preco As Double, qtd As Long, total As Double
            preco = CDbl(ws.Cells(i, RI_PRECO_VENDA).Value)
            qtd   = CLng(ws.Cells(i, RI_QUANTIDADE).Value)
            total = CDbl(ws.Cells(i, RI_TOTAL).Value)
            
            Set itm = lvItens.ListItems.Add(Text:=ws.Cells(i, RI_CODIGO_PRODUTO).Value)
            itm.SubItems(1) = ws.Cells(i, RI_DESCRICAO).Value
            itm.SubItems(2) = ws.Cells(i, RI_CATEGORIA).Value
            itm.SubItems(3) = CStr(qtd)
            itm.SubItems(4) = Format(preco, "R$ #,##0.00")
            itm.SubItems(5) = Format(total, "R$ #,##0.00")
        End If
    Next i
    
    AtualizarTotalizadores
End Sub

' Confirma saída do romaneio: PENDENTE → CONFERIDO + movimentação + PDF
Private Sub cmdConfirmarSaida_Click()
    If lvRomaneios.SelectedItem Is Nothing Then
        MsgBox "Selecione um romaneio para confirmar a saída.", vbExclamation
        Exit Sub
    End If
    
    Dim idRom  As String
    Dim status As String
    idRom  = lvRomaneios.SelectedItem.Text
    status = UCase(Trim(lvRomaneios.SelectedItem.SubItems(5)))
    
    If status <> "PENDENTE" Then
        MsgBox "Apenas romaneios com status PENDENTE podem ser confirmados.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Confirmar saída do romaneio " & idRom & "?" & vbNewLine & _
              "Isso irá gerar as movimentações de estoque e o PDF.", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Dim ws As Worksheet
    Dim ult As Long, i As Long
    Dim nomeRev As String
    Dim ok As Boolean
    
    Set ws = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    
    Application.ScreenUpdating = False
    ok = True
    
    For i = 2 To ult
        If CStr(ws.Cells(i, RI_ID_ROMANEIO).Value) = idRom Then
            nomeRev = ws.Cells(i, RI_NOME_REVENDEDORA).Value
            
            Dim codProd As String
            Dim qtd As Long
            codProd = CStr(ws.Cells(i, RI_CODIGO_PRODUTO).Value)
            qtd     = CLng(ws.Cells(i, RI_QUANTIDADE).Value)
            
            ' Grava movimentação: SAIDA / REVENDA / ESTOQUE DISPONIVEL → EM MALETA
            Dim gravou As Boolean
            gravou = GravarMovimentacao( _
                codProd, "SAIDA", "REVENDA", _
                "ESTOQUE DISPONIVEL", "EM MALETA", _
                "", nomeRev, qtd, _
                "Romaneio " & idRom, idRom)
            
            If Not gravou Then
                ok = False
                MsgBox "Erro na movimentação do produto " & codProd & ". Processo interrompido.", vbCritical
                GoTo Finalizar
            End If
            
            ' Atualiza status do item para CONFERIDO
            ws.Cells(i, RI_STATUS).Value           = "CONFERIDO"
            ws.Cells(i, RI_DATA_CONFERENCIA).Value = Now
        End If
    Next i
    
    If ok Then
        ' Gera PDF com layout da aba IMP_ROMANEIOS
        GerarPDFRomaneio idRom
        Application.ScreenUpdating = True
        MsgBox "Romaneio " & idRom & " confirmado!" & vbNewLine & "Movimentações geradas e PDF exportado.", vbInformation
        cmdConsultarRom_Click   ' Atualiza lista
    End If
    
Finalizar:
    Application.ScreenUpdating = True
End Sub

' Gera PDF a partir do layout da aba IMP_ROMANEIOS
Private Sub GerarPDFRomaneio(ByVal idRom As String)
    Dim wsImp As Worksheet
    Dim wsRom As Worksheet
    Dim ult As Long, i As Long
    Dim linhaDados As Long
    
    On Error GoTo ErroImp
    
    Set wsImp = ThisWorkbook.Worksheets("IMP_ROMANEIOS")
    Set wsRom = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = wsRom.Cells(wsRom.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    
    ' Preenche cabeçalho (ajuste as células conforme seu layout)
    wsImp.Range("B2").Value = idRom
    wsImp.Range("B3").Value = ""
    
    For i = 2 To ult
        If CStr(wsRom.Cells(i, RI_ID_ROMANEIO).Value) = idRom Then
            wsImp.Range("B3").Value = wsRom.Cells(i, RI_NOME_REVENDEDORA).Value
            wsImp.Range("B4").Value = Format(wsRom.Cells(i, RI_DATA_CONFERENCIA).Value, "dd/mm/yyyy")
            wsImp.Range("B5").Value = Responsavel
            Exit For
        End If
    Next i
    
    ' Limpa área de itens (ajuste o range conforme seu layout)
    Dim rngItens As Range
    Set rngItens = wsImp.Range("A9:F200")
    rngItens.ClearContents
    
    linhaDados = 9
    Dim totalVal As Double
    Dim totalQtd As Long
    
    For i = 2 To ult
        If CStr(wsRom.Cells(i, RI_ID_ROMANEIO).Value) = idRom Then
            wsImp.Cells(linhaDados, 1).Value = wsRom.Cells(i, RI_CODIGO_PRODUTO).Value
            wsImp.Cells(linhaDados, 2).Value = wsRom.Cells(i, RI_DESCRICAO).Value
            wsImp.Cells(linhaDados, 3).Value = wsRom.Cells(i, RI_CATEGORIA).Value
            wsImp.Cells(linhaDados, 4).Value = wsRom.Cells(i, RI_QUANTIDADE).Value
            wsImp.Cells(linhaDados, 5).Value = wsRom.Cells(i, RI_PRECO_VENDA).Value
            wsImp.Cells(linhaDados, 6).Value = wsRom.Cells(i, RI_TOTAL).Value
            totalQtd = totalQtd + CLng(wsRom.Cells(i, RI_QUANTIDADE).Value)
            totalVal = totalVal + CDbl(wsRom.Cells(i, RI_TOTAL).Value)
            linhaDados = linhaDados + 1
        End If
    Next i
    
    ' Totais (ajuste a linha conforme layout)
    wsImp.Cells(linhaDados + 1, 4).Value = totalQtd
    wsImp.Cells(linhaDados + 1, 6).Value = totalVal
    
    ' Exporta PDF
    Dim caminhoPDF As String
    caminhoPDF = ThisWorkbook.Path & "\ROMANEIOS\" & idRom & ".pdf"
    
    ' Cria pasta se não existir
    If Dir(ThisWorkbook.Path & "\ROMANEIOS\", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\ROMANEIOS\"
    End If
    
    wsImp.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminhoPDF, Quality:=xlQualityStandard
    
    MsgBox "PDF gerado em: " & caminhoPDF, vbInformation
    Exit Sub
    
ErroImp:
    MsgBox "Aviso: PDF não pôde ser gerado (" & Err.Description & ")." & vbNewLine & _
           "As movimentações foram gravadas normalmente.", vbExclamation
End Sub

' ============================================================
' ABA 2 — CONFERÊNCIA (Retorno de Maletas)
' ============================================================

Private Sub ConfigurarListViewConferencia()
    With lvConferencia
        .View = 3
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Código",        70
        .ColumnHeaders.Add , , "Descrição",     150
        .ColumnHeaders.Add , , "Qtd Enviada",   70
        .ColumnHeaders.Add , , "Qtd Retornou",  80
        .ColumnHeaders.Add , , "Qtd Vendida",   70
        .ColumnHeaders.Add , , "Preço Unit.",   70
        .ColumnHeaders.Add , , "Val. Retorno",  80
        .ColumnHeaders.Add , , "Val. Venda",    80
    End With
End Sub

' Busca itens do romaneio para conferência
Private Sub cmdBuscarRomaneio_Click()
    Dim idRom As String
    idRom = Trim(txtRomaneioConf.Value)
    
    If idRom = "" Then
        MsgBox "Informe o número do romaneio.", vbExclamation
        txtRomaneioConf.SetFocus
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim ult As Long, i As Long
    Dim encontrou As Boolean
    
    Set ws = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = ws.Cells(ws.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    
    lvConferencia.ListItems.Clear
    lblInfoConferencia.Caption = ""
    encontrou = False
    
    For i = 2 To ult
        If UCase(CStr(ws.Cells(i, RI_ID_ROMANEIO).Value)) = UCase(idRom) Then
            Dim statusItem As String
            statusItem = UCase(Trim(ws.Cells(i, RI_STATUS).Value))
            
            If Not encontrou Then
                ' Valida status: só CONFERIDO pode ser finalizado
                If statusItem = "PENDENTE" Then
                    MsgBox "Este romaneio ainda não foi confirmado (status PENDENTE)." & vbNewLine & _
                           "Confirme a saída na aba CONSULTAR ROMANEIO antes de fazer a conferência de retorno.", vbExclamation
                    Exit Sub
                End If
                
                If statusItem = "FINALIZADO" Then
                    MsgBox "Este romaneio já foi finalizado.", vbInformation
                    Exit Sub
                End If
                
                lblInfoConferencia.Caption = "Revendedora: " & ws.Cells(i, RI_NOME_REVENDEDORA).Value & _
                                             "   |   Data Saída: " & Format(ws.Cells(i, RI_DATA_CONFERENCIA).Value, "dd/mm/yyyy") & _
                                             "   |   Status: " & ws.Cells(i, RI_STATUS).Value
                encontrou = True
            End If
            
            Dim qtdEnv As Long
            Dim preco  As Double
            qtdEnv = CLng(ws.Cells(i, RI_QUANTIDADE).Value)
            preco  = CDbl(ws.Cells(i, RI_PRECO_VENDA).Value)
            
            Dim itm As ListItem
            Set itm = lvConferencia.ListItems.Add(Text:=ws.Cells(i, RI_CODIGO_PRODUTO).Value)
            itm.SubItems(1) = ws.Cells(i, RI_DESCRICAO).Value
            itm.SubItems(2) = CStr(qtdEnv)
            itm.SubItems(3) = ""          ' Qtd retornou — a preencher
            itm.SubItems(4) = ""          ' Qtd vendida — calculado
            itm.SubItems(5) = Format(preco, "R$ #,##0.00")
            itm.SubItems(6) = ""          ' Valor retorno
            itm.SubItems(7) = ""          ' Valor venda
        End If
    Next i
    
    If Not encontrou Then
        MsgBox "Romaneio não encontrado ou sem itens elegíveis.", vbExclamation
    End If
End Sub

' Clique num item do lvConferencia → usuário edita a qtd no txtQtdRetornou
' A lógica real de preenchimento é feita pelo usuário direto na célula
' (para simplificar sem controle externo complexo, usamos InputBox ao dar duplo clique)
Private Sub lvConferencia_DblClick()
    If lvConferencia.SelectedItem Is Nothing Then Exit Sub
    
    Dim qtdEnv As Long
    Dim qtdRet As Variant
    Dim cod    As String
    
    cod    = lvConferencia.SelectedItem.Text
    qtdEnv = CLng(lvConferencia.SelectedItem.SubItems(2))
    
    qtdRet = InputBox("Produto: " & lvConferencia.SelectedItem.SubItems(1) & vbNewLine & _
                      "Quantidade enviada: " & qtdEnv & vbNewLine & vbNewLine & _
                      "Quantas unidades RETORNARAM?", "Conferência de Retorno", 0)
    
    If qtdRet = "" Then Exit Sub
    
    If Not IsNumeric(qtdRet) Then
        MsgBox "Informe um número válido.", vbExclamation
        Exit Sub
    End If
    
    Dim qRet As Long
    qRet = CLng(qtdRet)
    
    If qRet < 0 Or qRet > qtdEnv Then
        MsgBox "Quantidade inválida. Deve ser entre 0 e " & qtdEnv & ".", vbExclamation
        Exit Sub
    End If
    
    Dim preco As Double
    Dim precoS As String
    precoS = Replace(Replace(lvConferencia.SelectedItem.SubItems(5), "R$", ""), " ", "")
    preco  = CDbl(precoS)
    
    Dim qVend As Long
    qVend = qtdEnv - qRet
    
    lvConferencia.SelectedItem.SubItems(3) = CStr(qRet)
    lvConferencia.SelectedItem.SubItems(4) = CStr(qVend)
    lvConferencia.SelectedItem.SubItems(6) = Format(qRet  * preco, "R$ #,##0.00")
    lvConferencia.SelectedItem.SubItems(7) = Format(qVend * preco, "R$ #,##0.00")
End Sub

' Pré-conferência: valida se todos os itens foram preenchidos
Private Sub cmdPreConferencia_Click()
    If lvConferencia.ListItems.Count = 0 Then
        MsgBox "Busque um romaneio primeiro.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    Dim pendentes As Long
    Dim totalEnv As Long, totalRet As Long, totalVend As Long
    Dim totalValRet As Double, totalValVend As Double
    
    For i = 1 To lvConferencia.ListItems.Count
        If lvConferencia.ListItems(i).SubItems(3) = "" Then
            pendentes = pendentes + 1
        Else
            totalEnv  = totalEnv  + CLng(lvConferencia.ListItems(i).SubItems(2))
            totalRet  = totalRet  + CLng(lvConferencia.ListItems(i).SubItems(3))
            totalVend = totalVend + CLng(lvConferencia.ListItems(i).SubItems(4))
            
            Dim valRet As String, valVend As String
            valRet  = Replace(Replace(lvConferencia.ListItems(i).SubItems(6), "R$", ""), " ", "")
            valVend = Replace(Replace(lvConferencia.ListItems(i).SubItems(7), "R$", ""), " ", "")
            On Error Resume Next
            totalValRet  = totalValRet  + CDbl(valRet)
            totalValVend = totalValVend + CDbl(valVend)
            On Error GoTo 0
        End If
    Next i
    
    Dim msg As String
    msg = "=== PRÉ-CONFERÊNCIA ===" & vbNewLine & vbNewLine
    msg = msg & "Itens com quantidade preenchida: " & (lvConferencia.ListItems.Count - pendentes) & "/" & lvConferencia.ListItems.Count & vbNewLine
    If pendentes > 0 Then
        msg = msg & "⚠ " & pendentes & " item(ns) sem quantidade — preencha antes de salvar!" & vbNewLine
    End If
    msg = msg & vbNewLine
    msg = msg & "Total enviado:    " & totalEnv & " un." & vbNewLine
    msg = msg & "Total retornou:   " & totalRet & " un. (" & Format(totalValRet, "R$ #,##0.00") & ")" & vbNewLine
    msg = msg & "Total vendido:    " & totalVend & " un. (" & Format(totalValVend, "R$ #,##0.00") & ")" & vbNewLine
    
    MsgBox msg, vbInformation, "Resumo da Conferência"
End Sub

' Salva conferência definitiva e finaliza o romaneio
Private Sub cmdSalvarConferencia_Click()
    If Trim(Responsavel) = "" Then
        MsgBox "Nenhum usuário logado.", vbCritical
        Exit Sub
    End If
    
    If lvConferencia.ListItems.Count = 0 Then
        MsgBox "Busque um romaneio primeiro.", vbExclamation
        Exit Sub
    End If
    
    ' Valida se todos os itens têm quantidade preenchida
    Dim i As Long
    For i = 1 To lvConferencia.ListItems.Count
        If lvConferencia.ListItems(i).SubItems(3) = "" Then
            MsgBox "Preencha a quantidade retornada de todos os itens antes de salvar." & vbNewLine & _
                   "Dê duplo clique em cada item para informar a quantidade.", vbExclamation
            Exit Sub
        End If
    Next i
    
    Dim idRom As String
    idRom = Trim(txtRomaneioConf.Value)
    
    If MsgBox("Confirmar conferência de retorno do romaneio " & idRom & "?" & vbNewLine & _
              "Esta ação é irreversível e irá finalizar o romaneio.", _
              vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim wsRom As Worksheet
    Dim ult As Long
    Dim linhaPorCodigo As Object
    
    Set wsRom = ThisWorkbook.Worksheets("ROMANEIOS_ITENS")
    ult = wsRom.Cells(wsRom.Rows.Count, RI_ID_ROMANEIO).End(xlUp).Row
    Set linhaPorCodigo = CreateObject("Scripting.Dictionary")
    
    ' Mapeia linha na aba por código de produto para atualizar depois
    Dim j As Long
    For j = 2 To ult
        If CStr(wsRom.Cells(j, RI_ID_ROMANEIO).Value) = idRom Then
            linhaPorCodigo(wsRom.Cells(j, RI_CODIGO_PRODUTO).Value) = j
        End If
    Next j
    
    Dim nomeRev As String
    nomeRev = lblInfoConferencia.Caption
    ' Extrai nome da revendedora do caption
    If InStr(nomeRev, "Revendedora:") > 0 Then
        nomeRev = Trim(Mid(nomeRev, InStr(nomeRev, "Revendedora:") + 12))
        Dim pipePos As Long
        pipePos = InStr(nomeRev, "|")
        If pipePos > 0 Then nomeRev = Trim(Left(nomeRev, pipePos - 1))
    End If
    
    Dim erros As Long
    erros = 0
    
    For i = 1 To lvConferencia.ListItems.Count
        Dim cod     As String
        Dim qRet    As Long
        Dim qVend   As Long
        Dim qtdEnv  As Long
        
        cod    = lvConferencia.ListItems(i).Text
        qtdEnv = CLng(lvConferencia.ListItems(i).SubItems(2))
        qRet   = CLng(lvConferencia.ListItems(i).SubItems(3))
        qVend  = CLng(lvConferencia.ListItems(i).SubItems(4))
        
        ' Retorno: ENTRADA / REVENDA / EM MALETA → ESTOQUE DISPONIVEL
        If qRet > 0 Then
            Dim gravouRet As Boolean
            gravouRet = GravarMovimentacao( _
                cod, "ENTRADA", "REVENDA", _
                "EM MALETA", "ESTOQUE DISPONIVEL", _
                nomeRev, "", qRet, _
                "Retorno romaneio " & idRom, idRom)
            If Not gravouRet Then erros = erros + 1
        End If
        
        ' Vendas: SAIDA / VENDA / EM MALETA → VENDA
        If qVend > 0 Then
            Dim gravouVend As Boolean
            gravouVend = GravarMovimentacao( _
                cod, "SAIDA", "VENDA", _
                "EM MALETA", "VENDA", _
                nomeRev, "", qVend, _
                "Venda romaneio " & idRom, idRom)
            If Not gravouVend Then erros = erros + 1
        End If
        
        ' Atualiza linha na aba ROMANEIOS_ITENS
        If linhaPorCodigo.Exists(cod) Then
            Dim linhaR As Long
            linhaR = linhaPorCodigo(cod)
            wsRom.Cells(linhaR, RI_STATUS).Value       = "FINALIZADO"
            wsRom.Cells(linhaR, RI_QTD_RETORNOU).Value = qRet
            wsRom.Cells(linhaR, RI_QTD_VENDIDA).Value  = qVend
            wsRom.Cells(linhaR, RI_DATA_RETORNO).Value = Now
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If erros = 0 Then
        MsgBox "Romaneio " & idRom & " finalizado com sucesso!" & vbNewLine & _
               "Estoque atualizado.", vbInformation
        LimparConferencia
    Else
        MsgBox erros & " erro(s) ao gravar movimentações. Verifique os estoques.", vbExclamation
    End If
End Sub

Private Sub cmdLimparConferencia_Click()
    LimparConferencia
End Sub

Private Sub LimparConferencia()
    txtRomaneioConf.Value = ""
    lvConferencia.ListItems.Clear
    lblInfoConferencia.Caption = ""
End Sub
