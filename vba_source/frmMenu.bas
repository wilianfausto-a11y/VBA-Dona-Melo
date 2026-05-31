Attribute VB_Name = "frmMenu"
Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================
' MENU PRINCIPAL — pós-login
' Controles esperados no UserForm:
'   lblUsuario      — Label: "Usuário: [nome]"
'   cmdProdutos     — Botão: Cadastro de Produtos
'   cmdMovimentacao — Botão: Movimentação de Estoque
'   cmdRevendedoras — Botão: Cadastro de Revendedoras
'   cmdRomaneio     — Botão: Romaneios
'   cmdSair         — Botão: Sair / Encerrar sessão
' ============================================================

Private Sub UserForm_Initialize()
    Me.lblUsuario.Caption = "Usuário logado: " & Responsavel
    Me.Caption = "DonaMelo — Menu Principal"
End Sub

Private Sub cmdProdutos_Click()
    Me.Hide
    frmProdutos.Show
    Me.Show
End Sub

Private Sub cmdMovimentacao_Click()
    Me.Hide
    frmMovimentacao.Show
    Me.Show
End Sub

Private Sub cmdRevendedoras_Click()
    Me.Hide
    frmRevendedoras.Show
    Me.Show
End Sub

Private Sub cmdRomaneio_Click()
    Me.Hide
    frmRomaneio.Show
    Me.Show
End Sub

Private Sub cmdSair_Click()
    If MsgBox("Deseja encerrar a sessão?", vbYesNo + vbQuestion, "Sair") = vbYes Then
        Responsavel = ""
        Unload Me
    End If
End Sub
