Attribute VB_Name = "modGenerateReport"
Option Explicit

Sub get_data_and_connect()
    Dim sql_statement As String
    sql_statement = "SELECT * FROM qry_production_report_data WHERE ranking <= 20;"
    
    Dim source_path_conn As String
    source_path_conn = ThisWorkbook.Path & "\production_lineDB.accdb"
    
    Dim dest_path As Worksheet
    Dim dest_range As Range
    
    Set dest_path = ThisWorkbook.Sheets("query_result")
    dest_path.Cells.ClearContents
    dest_path.Cells.Clear
    
    Set dest_range = ThisWorkbook.Sheets("query_result").Range("A2")
    
    Call get_data_from_db(source_path_conn, dest_path, dest_range, sql_statement)
    
    Call set_pivot_connection_new
End Sub

Sub copy_and_paste_template(sht_to_copy As Worksheet, dest_wkb As Workbook)
    Dim sht_to_paste As Worksheet
    Set sht_to_paste = dest_wkb.Sheets(dest_wkb.Sheets.Count)
    sht_to_copy.Copy after:=sht_to_paste
End Sub

Sub run_report()
    'get data and connect pivot tables
    Call get_data_and_connect
    
    Dim csht As Worksheet
    Set csht = ThisWorkbook.Worksheets("output")
    
    Dim ith_site As Long
    ith_site = ThisWorkbook.Sheets("production_sites").Range("A" & Rows.Count).End(xlUp).Row
    
    Dim site_range As Range
    Dim site As Range
    Set site_range = ThisWorkbook.Sheets("production_sites").Range("A2:A" & ith_site)
    
    Dim dest_wkb As Workbook
    Set dest_wkb = Workbooks.Add
    
    'copy template, name
    Dim psht As Worksheet
    Dim pvt As PivotTable
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For Each site In site_range
        Call copy_and_paste_template(csht, dest_wkb)
        Set psht = dest_wkb.Sheets(dest_wkb.Sheets.Count)
        psht.Name = LCase("production_site_" + site.Value)
        
        For Each pvt In psht.PivotTables
            pvt.PivotFields("production_site").ClearAllFilters
            pvt.PivotFields("production_site").CurrentPage = site.Value
        Next pvt
    Next site
 
    dest_wkb.Sheets(1).Activate
    Dim sht As Worksheet
    
    'delete blank sheets
    For Each sht In dest_wkb.Sheets
        If sht.Name Like ("Sheet*") Then sht.Delete
    Next sht
    
    'save template
    Dim wkb_path As String
    wkb_path = ThisWorkbook.Path
    If Len(Dir(wkb_path & "\excel_template", vbDirectory)) = 0 Then
        MkDir (wkb_path & "\excel_template")
    End If
    dest_wkb.SaveAs wkb_path & "\excel_template" & "\Production Report Template " & Format(Now, "dd-Mmm-yyyy hh-mm-ss") & ".xlsx"

    Call save_as_pdf(output_folder:="pdf_template", workbook_to_save:=dest_wkb)
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
