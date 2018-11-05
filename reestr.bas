Attribute VB_Name = "Module1"
' C:\Users\E.Kozodoy\AppData\Roaming\Microsoft\Excel\XLSTART
Dim WS As Worksheet
Dim WS2 As Worksheet
Dim y As Boolean
Dim i As Byte


Public Sub reestr()
    order
    saveas
    copsh
End Sub

Private Sub copsh()
    
    Dim WB1 As Workbook
    Dim WB2 As Workbook
    Set WB1 = Workbooks.Open("D:\OneDrive\Business Intelligence\Sources\Reestr\FinReestr.xlsm")
    Set WB2 = Workbooks.Open("D:\OneDrive\Business Intelligence\Sources\Reestr\sources\Reestr.xlsm")
        Application.DisplayAlerts = False
        For Each WS2 In WB2.Worksheets
            i = 1
            y = False
            Do While y = False And i <= WB1.Worksheets.Count
                If WS2.Name = WB1.Worksheets(i).Name Then y = True Else i = i + 1
            Loop
                
                If y = False Then WS2.Copy After:=WB1.Worksheets(WB1.Worksheets.Count)
                'sh.Copy After:=wb.Sheets(wb.sheets.count)
                'Workbooks(WB1).Worksheets (Workshhets.Count)
                'If y = False Then WS2.Copy After:=Worksheets(WB1.Worksheets.Count)
                
                    
        Next WS2
        
        WB1.Save
        WB2.Save
        WB2.Close
        WB1.Activate
        Application.DisplayAlerts = True
End Sub


Private Sub order()
 Dim WSname As String
    'Dim WS As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    For Each WS In Worksheets
        If WS.Name = "черновик" Then
        WS.Delete
        Else
        y = False
        i = 1
        
            Do While y = False And i < 30
            If WS.Cells(i, 1) = "№ пп" Then y = True Else i = i + 1
            Loop
            If y = True Then
                    Do While WS.Cells(1, 1) <> "№ пп"
                        If Trim(WS.Cells(1, 1)) = "за" Then
                        WSname = WS.Cells(1, 2)
                        WS.Rows(1).Delete
                        WS.Name = WSname
                        Else
                        WS.Rows(1).Delete
                        End If
                    
                    Loop
            Else: Exit Sub
            End If
        End If
    Next WS
    
   
End Sub

Private Sub saveas()
    Dim fname As String
    Dim oname As String
    
    Dim n As Integer
    
    
    fname = "D:\OneDrive\Business Intelligence\Sources\Reestr\sources\Reestr"
    oname = fname & ".xlsm"
    'x = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    Do While Dir(oname) <> ""
        n = n + 1
        oname = fname & n & ".xlsm"
    Loop
    On Error Resume Next
    Application.DisplayAlerts = False
        If ActiveWorkbook.Path & "\" & ActiveWorkbook.Name = fname & ".xlsm" Then
        ActiveWorkbook.saveas Filename:=oname, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ActiveWorkbook.saveas Filename:=fname, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Else
        FileCopy Source:=fname & ".xlsm", Destination:=oname
        ActiveWorkbook.saveas Filename:=fname, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        End If
        ActiveWorkbook.Close
    
    
    '    FileCopy Source:=fname & ".xlsm", Destination:=oname
     '   ActiveWorkbook.SaveAs Filename:=fname, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    'Kill (fname & ".*")
    
    
End Sub



