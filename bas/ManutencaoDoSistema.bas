Attribute VB_Name = "ManutencaoDoSistema"
Option Compare Database

Public Function BKP()
Dim strMDB As String

strMDB = Dialogo("Backup da Base de Dados", False, "Microsoft Office Access", "*.MDB;*.MDE")

'Se selecionou arquivo, atualiza os vínculos
If strMDB <> "" Then
    'Gera bkp da base de dados
    Backup strMDB
End If

End Function

Public Function Alteracao()
Dim strSCR As String
Dim strMDB As String

strMDB = Dialogo("Backup da Base de Dados", False, "Microsoft Office Access", "*.MDB;*.MDE")

'Se selecionou arquivo, atualiza os vínculos
If strMDB <> "" Then
    'Gera bkp da base de dados
    Backup strMDB
    
    strSCR = Dialogo("Selecione o script de manutenção", False, "Script De Manutenção", "*.scr")
    
    If strSCR <> "" Then
        'Alterar a base de dados atravez de um script
        Application.Screen.MousePointer = 11
        ScriptManutencao strSCR, strMDB
        Application.Screen.MousePointer = 0
        MsgBox "Alteração(ões) realizada(s) com sucesso!", vbInformation, "Alteração da Base de Dados"
    End If

End If

End Function

Public Sub ScriptManutencao(Script As String, Base As String)
Dim objAccess As Object
Dim strArq As String
Dim strSQL As String

Set objAccess = CreateObject("Access.Application")
objAccess.OpenCurrentDatabase Base

Open Script For Input As #1
Do Until EOF(1)
    Line Input #1, strSQL
    objAccess.DoCmd.RunSQL (strSQL)
Loop
Close #1

objAccess.CloseCurrentDatabase
Set objAccess = Nothing

'apaga o arquivo temp se existir
If Dir(Script) <> "" Then Kill Script

End Sub
Public Function Backup(sFileName As String)
'===================================================================
'   Funções agregadas a esta função:
'   > CompactarRepararDatabase
'   > CriarPasta
'   > getPath
'   > getFileName
'   > getFileExt
''===================================================================

Dim oFSO As New FileSystemObject
Dim oPasta As New FileSystemObject
Dim oSHL
Dim tmp, p1, p2, p3, p4, p5
Dim Origem As String
Dim sOrigem As String
Dim sDestino As String
Dim sArquivo As String
Dim sExtencao As String

sDestino = "Backup"
sOrigem = getPath(sFileName)
sArquivo = getFileName(sFileName)
sExtencao = getFileExt(sFileName)

On Error Resume Next
Err.Clear

Origem = sOrigem & "\" & sArquivo & sExtencao

'Começa o bkp se o arquivo existir na origem
If Dir(Origem) <> "" Then
   
    Application.Screen.MousePointer = 11
   
    p1 = Right("00" & Year(Now()), 2)
    p2 = Right("00" & Month(Now()), 2)
    p3 = Right("00" & Day(Now()), 2)
    p4 = Right("00" & Hour(Now()), 2)
    p5 = Right("00" & Minute(Now()), 2)
     
    tmp = ("_" & p1 & p2 & p3 & "_" & p4 & p5)
    
    CompactarRepararDatabase sOrigem & "\" & sArquivo & sExtencao
    
    sOrigem = sOrigem & "\"
    
    oFSO.CopyFile sOrigem & sArquivo & sExtencao, CriarPasta(sDestino) & sArquivo & tmp & sExtencao, True
    
    If Err <> 0 Then
        MsgBox "Error: " & Err & " " & Err.Description, vbExclamation, "Backup"
    Else
        Set oSHL = CreateObject("WScript.Shell")
        oSHL.PopUp "Backup completo!", 1, "Backup", 0 + 64
    End If
     
    Application.Screen.MousePointer = 0
    
Else
    
    MsgBox "ATENÇÃO: Execute esta operação apartir do computador que contém os dados do sistema", vbInformation + vbOKOnly, "Backup"
    
End If

End Function

Public Function CompactarRepararDatabase(DatabasePath As String, Optional Password As String, Optional TempFile As String = "c:\tmp.mdb")
'===================================================================
' Se a versao DAO for anterior a 3.6 , entao devemos usar o método RepairDatabase
' Se a versao DAO for a 3.6 ou superior basta usar a função CompactDatabase
'===================================================================

If DBEngine.Version < "3.6" Then DBEngine.RepairDatabase DatabasePath

'se nao informou um arquivo temporario usa "c:\tmp.mdb"
If TempFile = "" Then TempFile = "c:\tmp.mdb"

'apaga o arquivo temp se existir
If Dir(TempFile) <> "" Then Kill TempFile

'formata a senha no formato ";pwd=PASSWORD" se a mesma existir
If Password <> "" Then Password = ";pwd=" & Password

'compacta a base criando um novo banco de dados
DBEngine.CompactDatabase DatabasePath, TempFile, , , Password

'apaga o primeiro banco de dados
Kill DatabasePath

'move a base compactada para a origem
FileCopy TempFile, DatabasePath

'apaga o arquivo temporario
Kill TempFile

End Function

Public Function CriarPasta(sPasta As String) As String
'Cria pasta apartir da origem do sistema

Dim fPasta As New FileSystemObject
Dim MyApl As String

MyApl = Application.CurrentProject.Path
        
If Not fPasta.FolderExists(MyApl & "\" & sPasta) Then
   fPasta.CreateFolder (MyApl & "\" & sPasta)
End If

CriarPasta = MyApl & "\" & sPasta & "\"

End Function

Public Function getPath(sPathIn As String) As String
'Esta função irá retornar apenas o path de uma string que contenha o path e o nome do arquivo:
Dim I As Integer

  For I = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, I, 1)) Then Exit For
  Next
  
  getPath = Left$(sPathIn, I)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next
  
  getFileName = Left(Mid$(sFileIn, I + 1, Len(sFileIn) - I), Len(Mid$(sFileIn, I + 1, Len(sFileIn) - I)) - 4)

End Function

Public Function getFileExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next
  
  getFileExt = Right(Mid$(sFileIn, I + 1, Len(sFileIn) - I), 4)

End Function



Public Function Dialogo(Titulo As String, MultiSelecao As Boolean, FiltroTitulo As String, FiltroExtencao As String) As String

Dim lngCount As Long

' Open the file dialog
With Application.FileDialog(msoFileDialogOpen)
    .AllowMultiSelect = True
    .Filters.Add FiltroTitulo, FiltroExtencao
    .Title = Titulo
    .AllowMultiSelect = MultiSelecao
    .Show
    
    ' Display paths of each file selected
    For lngCount = 1 To .SelectedItems.Count
        Dialogo = .SelectedItems(lngCount)
    Next lngCount
    
End With

End Function

