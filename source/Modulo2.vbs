Public cnn As New Connection
Public rs As New Recordset
Public rs2 As New Recordset

Public strSQL As String

Public Const NOMBRE_HOJA_L As String = "L"

Public g_objFSO As Scripting.FileSystemObject
Public g_scrText As Scripting.TextStream

Sub Start()
    UserForm1.Show
End Sub

Public Sub OpenDB(ruta As String)
    If cnn.State = adStateOpen Then cnn.Close
    On Error GoTo Handle
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & _
        ruta
        cnn.Open
    Exit Sub
Handle:
    If cnn.Errors.Count > 0 Then
        Call Error_Handle(cnn.Errors.Item(0).Source, "Módulo2", strSQL, cnn.Errors.Item(0).Number, cnn.Errors.Item(0).Description)
    End If
    cnn.Errors.Clear
    closeRS
End Sub

Public Sub closeRS()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If cnn.State = adStateOpen Then cnn.Close
    Set cnn = Nothing
End Sub

Public Function LogFile_WriteError(ByVal sRoutineName As String, _
                             ByVal sMessage As String)
    Dim sText As String
    Dim pathLog As String
    pathLog = ActiveWorkbook.Path & "\log.txt"
   On Error GoTo ErrorHandler
   If (g_objFSO Is Nothing) Then
      Set g_objFSO = New FileSystemObject
   End If
   If (g_scrText Is Nothing) Then
      If (g_objFSO.FileExists(pathLog) = False) Then
         Set g_scrText = g_objFSO.OpenTextFile(pathLog, IOMode.ForWriting, True)
      Else
         Set g_scrText = g_objFSO.OpenTextFile(pathLog, IOMode.ForAppending)
      End If
   End If
   sText = sText & Format(Date, "DD/MM/YYYY") & " " & Time() & "|"
   sText = sText & sRoutineName & "|"
   sText = sText & sMessage & "|"
   g_scrText.WriteLine sText
   g_scrText.Close
   Set g_scrText = Nothing
   MsgBox "Log en: " & pathLog
   Exit Function
ErrorHandler:
   Set g_scrText = Nothing
   Call MsgBox("No se pudo escribir en el fichero log", vbCritical, "LogFile_WriteError")
End Function

Public Sub Error_Handle(ByVal sRoutineName As String, _
                         ByVal sObject As String, _
                         ByVal currentStrSQL As String, _
                         ByVal sErrorNo As String, _
                         ByVal sErrorDescription As String)
Dim sMessage As String
   sMessage = sObject & "|" & currentStrSQL & "|" & sErrorNo & "|" & sErrorDescription & "|" & Application.UserName
   Call MsgBox(sErrorNo & vbCrLf & sErrorDescription, vbCritical, sRoutineName & " - " & sObject & " - Error")
   Call LogFile_WriteError(sRoutineName, sMessage)
End Sub


