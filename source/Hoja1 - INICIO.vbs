Option Explicit

Private Sub btBuscarAccess_Click()
    Dim path_rep As String
    path_rep = openDialog
    If path_rep <> "FALSO" Then
        [ACCDB_PATH] = path_rep
    End If
End Sub

Private Sub btEjecutarQ_Click()
    If [STRQUERY] <> "" Then
        OpenDB ([ACCDB_PATH])
        rs.Open Hoja1.Range("STRQUERY").Value, cnn, adOpenKeyset, adLockOptimistic
        If rs.State = adStateOpen Then
            Hoja1.Range("dataSet").CopyFromRecordset rs
        End If
        closeRS
    End If
End Sub

Private Sub btExportar_Click()
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Dim strTemp As String
    Dim excluded(14) As String
    Dim sh As Worksheet
    Dim newsh As Worksheet
    Dim i As Integer
    
    excluded(0) = "MSysAccessStorage"
    excluded(1) = "MSysACEs"
    excluded(2) = "MSysComplexColumns"
    excluded(3) = "MSysNameMap"
    excluded(4) = "MSysNavPaneGroupCategories"
    excluded(5) = "MSysNavPaneGroups"
    excluded(6) = "MSysNavPaneGroupToObjects"
    excluded(7) = "MSysNavPaneObjectIDs"
    excluded(8) = "MSysObjects"
    excluded(9) = "MSysQueries"
    excluded(10) = "MSysRelationships"
    excluded(11) = "MSysResources"
    excluded(12) = "MSysAccessXML"
    excluded(13) = "MSysIMEXColumns"
    excluded(14) = "MSysIMEXSpecs"
    
    btLimpiar_Click
    
    OpenDB ([ACCDB_PATH])
    If cnn.State <> 0 Then
        Set rs = cnn.OpenSchema(adSchemaTables)
        Do While Not rs.EOF
            If Not IsInArray(rs.Fields("TABLE_NAME").Value, excluded) Then
                Set newsh = Sheets.Add(after:=Sheets(ThisWorkbook.Worksheets.Count))
                newsh.Name = rs.Fields("TABLE_NAME").Value
                
                strSQL = "SELECT * FROM " & rs.Fields("TABLE_NAME").Value
                
                On Error Resume Next
                rs2.Open strSQL, cnn, adOpenKeyset, adLockOptimistic
                On Error GoTo 0
                
                For i = 0 To rs2.Fields.Count - 1
                    newsh.Cells(1, 1 + i) = rs2.Fields(i).Name
                Next i
                
                newsh.Cells(2, 1).CopyFromRecordset rs2
                
                Set rs2 = Nothing
                
            End If
            rs.MoveNext
        Loop
        closeRS
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Private Function openDialog() As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

   With fd
        .InitialFileName = "C:\"
      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Por favor selecciones el reporte de conciliacion."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "ARCHIVO ACCDB (.accdb)", "*.accdb"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        openDialog = .SelectedItems(1)
    Else
        openDialog = "FALSO"
      End If
   End With
End Function

Private Sub btLimpiar_Click()
    Dim sh As Worksheet
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> "INICIO" Then
            sh.Delete
        End If
    Next sh
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub
