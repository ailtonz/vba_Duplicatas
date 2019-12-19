Attribute VB_Name = "modMouseHook"
Option Compare Database
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" _
Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
(ByVal hLibModule As Long) As Long

Private Declare Function StopMouseWheel Lib "MouseHook" _
(ByVal hWnd As Long, ByVal AccessThreadID As Long, _
Optional ByVal bNoSubformScroll As Boolean = False, Optional ByVal blIsGlobal As Boolean = False) As Boolean

Private Declare Function StartMouseWheel Lib "MouseHook" _
(ByVal hWnd As Long) As Boolean

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

' Instance returned from LoadLibrary call
Private hLib As Long


Public Function MouseWheelON() As Boolean
MouseWheelON = StartMouseWheel(Application.hWndAccessApp)
If hLib <> 0 Then
    hLib = FreeLibrary(hLib)
End If
End Function

Public Function MouseWheelOFF(Optional NoSubFormScroll As Boolean = False, Optional GlobalHook As Boolean = False) As Boolean
Dim s As String
Dim blRet As Boolean
Dim AccessThreadID As Long

On Error Resume Next
' Our error string
s = "Sorry...cannot find the MouseHook.dll file" & vbCrLf
s = s & "Please copy the MouseHook.dll file to your Windows System folder or into the same folder as this Access MDB."

' OK Try to load the DLL assuming it is in the Window System folder
hLib = LoadLibrary("MouseHook.dll")
If hLib = 0 Then
    ' See if the DLL is in the same folder as this MDB
    ' CurrentDB works with both A97 and A2K or higher
    hLib = LoadLibrary(CurrentDBDir() & "MouseHook.dll")
    If hLib = 0 Then
        MsgBox s, vbOKOnly, "MISSING MOUSEHOOK.dll FILE"
        MouseWheelOFF = False
        Exit Function
    End If
End If

' Get the ID for this thread
AccessThreadID = GetCurrentThreadId()
' Call our MouseHook function in the MouseHook dll.
' Please not the Optional GlobalHook BOOLEAN parameter
' Several developers asked for the MouseHook to be able to work with
' multiple instances of Access. In order to accomodate this request I
' have modified the function to allow the caller to
' specify a thread specific(this current instance of Access only) or
' a global(all applications) MouseWheel Hook.
' Only use the GlobalHook if you will be running multiple instances of Access!
MouseWheelOFF = StopMouseWheel(Application.hWndAccessApp, AccessThreadID, NoSubFormScroll, GlobalHook)

End Function


'******************** Code Begin ****************
'Code courtesy of
'Terry Kreft & Ken Getz
'
Function CurrentDBDir() As String
Dim strDBPath As String
Dim strDBFile As String
    strDBPath = CurrentDb.Name
    strDBFile = Dir(strDBPath)
    CurrentDBDir = Left$(strDBPath, Len(strDBPath) - Len(strDBFile))
End Function
'******************** Code End ****************

