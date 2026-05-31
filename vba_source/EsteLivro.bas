Attribute VB_Name = "EsteLivro"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    ' Abre o login ao iniciar o arquivo
    Application.Visible = True
    frmLogin.Show
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Limpa usuário logado ao fechar
    Responsavel = ""
End Sub
