VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

#If Win64 Then
  Private oWnd As Object
#End If

Private oSC As Object
Private hDB As Long
Private oDynWrapper As Variant


Private Sub Class_Initialize()
  Dim objFSO, scriptPath, sDLLPath As String

  hDB = 0
  #If Win64 Then
    Dim bRunning As Boolean
    If oWnd Is Nothing Then
      Dim sSignature, oShellWnd, oProc
      
      On Error Resume Next
      Do Until Len(sSignature) = 32
          sSignature = sSignature & Hex(Int(Rnd * 16))
      Loop
      CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
      Do
          For Each oShellWnd In CreateObject("Shell.Application").Windows
              Set oWnd = oShellWnd.GetProperty(sSignature)
              If Err.Number = 0 Then Exit Do
              Err.Clear
          Next
      Loop
      
      oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
    End If
    Set oSC = oWnd.CreateObjectx86("ScriptControl")
  #Else
    Set oSC = CreateObject("ScriptControl")
  #End If
    
  oSC.language = "vbscript"
  
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  scriptPath = objFSO.BuildPath(objFSO.GetParentFolderName(CurrentDb().Name), "SQLite.vbs")
  
  oSC.AddCode objFSO.OpenTextFile(scriptPath).readAll()
  
  sDLLPath = objFSO.BuildPath(objFSO.GetParentFolderName(CurrentDb().Name), "sqlite3.dll")
    
  Set oDynWrapper = oSC.Run("InitDLL", sDLLPath)
End Sub

Public Sub LogToFile(spFileName)
  Dim fl
  Set fl = oSC.Run("GetFileLogger", spFileName)
  oSC.Run "SetPrintProvider", fl
End Sub

Public Sub LogToImmediate()
  oSC.Run "SetPrintProvider", Me
End Sub

Public Sub OffLog()
  oSC.Run "SetPrintProvider", Nothing
End Sub

Public Sub Output(text As Variant)
  Debug.Print text
End Sub

Private Sub Class_Terminate()
  Set oDynWrapper = Nothing
  oSC.Run "FreeResource"
  Set oSC = Nothing
  #If Win64 Then
    If Not oWnd Is Nothing Then oWnd.Close
  #End If
End Sub

Public Function OpenDataBase(spFileName As String)
  Set OpenDataBase = oSC.Run("OpenDataBase", spFileName, Me, Application)
End Function

Public Property Get ScriptControl() As Object
  Set ScriptControl = oSC
End Property


Public Property Get DynamicWrapperX() As Object
  Set DynamicWrapperX = oDynWrapper
End Property
