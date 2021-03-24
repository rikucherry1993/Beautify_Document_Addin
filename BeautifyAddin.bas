Option Explicit

'-------------------------------------------------
'TASK1: Set home position, rate 100, breakPageView
'TASK2: Set header to file name, footer to page number
'TASK3: Delete all vertical page break lines
'TASK4: Delete personal information and document properties
'-------------------------------------------------

    Dim WS As Worksheet
    Dim WD As Window
    Dim sw As Boolean
    Dim WB As Workbook

Sub formatDocuments()

    On Error GoTo e

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbooks!"
        Exit Sub
    End If
    
    sw = False
    If Application.ScreenUpdating Then
        sw = True
    End If
    
    If sw Then
        Application.ScreenUpdating = False
    End If
  
    Set WB = ActiveWorkbook
  
    For Each WS In WB.Worksheets
        If WS.visible = xlSheetVisible Then
            WS.Activate
            WS.Range("A1").Select
            WB.Windows(1).ScrollRow = 1
            WB.Windows(1).ScrollColumn = 1
            WB.Windows(1).View = xlPageBreakPreview
            WB.Windows(1).Zoom = 100
            WS.PageSetup.CenterHeader =  WB.Name
            WS.PageSetup.CenterFooter = "&P / &N"
            WS.PageSetup.Zoom = False
            WS.PageSetup.FitToPagesWide = 1
            WS.PageSetup.FitToPagesTall = False
        End If
    Next

    For Each WS In WB.Worksheets
        If WS.visible = xlSheetVisible Then
            WS.Select
            Exit For
        End If
    Next


    'Delete personal information and document properties
    Application.DisplayAlerts = False
    WB.RemovePersonalInformation = True
    Call WB.RemoveDocumentInformation(xlRDIAll)

    Set WB = Nothing
    
    If sw Then
        Application.ScreenUpdating = True
    End If
    
    If ActiveWorkbook.ReadOnly Then
        MsgBox "Cannot save read-only files"
        GoTo pass
    End If
    
    If InStr(ActiveWorkbook.FullName, ".") = 0 Then
        MsgBox "File name begins with . "
        GoTo pass
    End If
    
    ActiveWorkbook.Save
    
pass:
    Exit Sub
    
e:
    MsgBox "Saving failed"
    
End Sub