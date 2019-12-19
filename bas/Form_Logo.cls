VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database

Dim fondo As New FondoAccess.CMDIWindow

Private Sub Form_Close()

 fondo.Unhook

End Sub

Private Sub Form_Open(Cancel As Integer)

Dim fArquivo As New FileSystemObject
Dim MyApl As String

MyApl = Application.CurrentProject.Path
                
If Not fArquivo.FileExists("c:\FondoAccess.dll") Then
    If fArquivo.FileExists(MyApl & "\FondoAccess.dll") Then
        fArquivo.CopyFile MyApl & "\FondoAccess.dll", "c:\FondoAccess.dll"
    End If
End If

If fArquivo.FileExists(MyApl & "\Logo.bmp") Then
    fondo.DrawMode = 1
    fondo.ImagePath = MyApl & "\Logo.bmp"
    fondo.Hook Application.hWndAccessApp
End If

End Sub

Private Sub Form_Timer()

    Me.Visible = False

End Sub
