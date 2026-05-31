Attribute VB_Name = "frmLogin"
Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================
' FORMULÁRIO DE LOGIN
' Controles esperados:
'   txtUsuario  — TextBox: nome de usuário
'   txtSenha    — TextBox: senha (PasswordChar = "*")
'   cmdEntrar   — CommandButton: entrar
'   lblErro     — Label: mensagem de erro (Visible=False por padrão)
' ============================================================

Private Sub UserForm_Initialize()
    Me.Caption = "DonaMelo — Acesso ao Sistema"
    lblErro.Visible = False
    txtUsuario.SetFocus
    Responsavel = ""
End Sub

Private Sub cmdEntrar_Click()
    AutenticarUsuario
End Sub

Private Sub txtSenha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then AutenticarUsuario
End Sub

Private Sub txtUsuario_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSenha.SetFocus
End Sub

Private Sub AutenticarUsuario()
    Dim ws As Worksheet
    Dim ult As Long
    Dim i As Long
    Dim usuario As String
    Dim senha As String
    
    usuario = Trim(txtUsuario.Value)
    senha   = Trim(txtSenha.Value)
    
    If usuario = "" Then
        MostrarErro "Informe o nome de usuário."
        txtUsuario.SetFocus
        Exit Sub
    End If
    
    If senha = "" Then
        MostrarErro "Informe a senha."
        txtSenha.SetFocus
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("USUARIOS")
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ult
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(usuario) Then
            If Trim(ws.Cells(i, 2).Value) = senha Then
                If UCase(Trim(ws.Cells(i, 3).Value)) = "ATIVO" Then
                    ' Login bem-sucedido
                    Responsavel = Trim(ws.Cells(i, 4).Value)
                    If Responsavel = "" Then Responsavel = usuario
                    Unload Me
                    frmMenu.Show
                    Exit Sub
                Else
                    MostrarErro "Usuário inativo. Contate o administrador."
                    Exit Sub
                End If
            Else
                MostrarErro "Senha incorreta."
                txtSenha.Value = ""
                txtSenha.SetFocus
                Exit Sub
            End If
        End If
    Next i
    
    MostrarErro "Usuário não encontrado."
    txtUsuario.SetFocus
End Sub

Private Sub MostrarErro(ByVal msg As String)
    lblErro.Caption = msg
    lblErro.Visible = True
End Sub
