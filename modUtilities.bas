Attribute VB_Name = "modUtilities"
Option Explicit

Sub save_as_pdf(output_folder As String, _
    workbook_to_save As Workbook)
    
    If Len(Dir(Application.ThisWorkbook.Path & "\" & output_folder, vbDirectory)) = 0 _
        Then MkDir (Application.ThisWorkbook.Path & "\" & output_folder)
    
    Dim ith_sheet As Long
    Dim wsh As Worksheet
     
    Dim pdf_collection As Collection
    Set pdf_collection = New Collection
    
    For ith_sheet = 1 To workbook_to_save.Sheets.Count
        pdf_collection.Add workbook_to_save.Sheets(ith_sheet).Name
    Next
    
    Dim ith_template As Long
    Dim template As String
    For ith_template = 1 To pdf_collection.Count
        template = CStr(pdf_collection.Item(ith_template))
        Set wsh = workbook_to_save.Sheets(template)
        
        wsh.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
             Application.ThisWorkbook.Path & "\" & output_folder & "\" & template & ".pdf", _
                Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Next ith_template

End Sub

Sub get_data_from_db(sourcePathConn As String, _
    destPath As Worksheet, _
    destRange As Range, _
    sqlStatement As String)

    Dim conn As Object
    Dim rs As Object
    Dim fld As Object
    
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    conn.ConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePathConn & ";Persist Security Info=False;"
    conn.Open
    
    With rs
        .ActiveConnection = conn
        .Source = sqlStatement
        .CursorType = 3
        .LockType = 1
        .Open
    End With
    
    destPath.Cells.Clear
    destPath.Cells.ClearContents
    
    destRange.CopyFromRecordset rs
    
    destPath.Activate
    destRange.Offset(-1, 0).Select
    For Each fld In rs.Fields
        ActiveCell.Value = fld.Name
        ActiveCell.Offset(0, 1).Select
    Next fld
    
    rs.Close
    If CBool(conn.State And 1) Then
        conn.Close
    End If
    
    Set rs = Nothing
    Set conn = Nothing
End Sub

Sub set_pivot_connection_new()
    Dim sourcePathConn As String
    sourcePathConn = Replace(ThisWorkbook.Name, ".xlsm", "")
    
    'Application.DisplayAlerts = False
    
    With ThisWorkbook.Connections(sourcePathConn).OLEDBConnection
        .CommandText = "query_result$"
        .CommandType = xlCmdTable
        .Connection = _
            "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";"
    End With
    
    ThisWorkbook.Connections(sourcePathConn).Refresh
    
    'Application.DisplayAlerts = True
End Sub

