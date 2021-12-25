Sub ExportHiddenSheetsToPDFs()

    'www.contextures.com
    'for Excel 2010 and later
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strName As String
    Dim strFolder As String
    Dim strFile As String
    Dim strPathFile As String
    Dim pdfGeneratedCount As Integer
    
    Set wbA = ActiveWorkbook
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            strFolder = .SelectedItems(1)
        End If
    End With
    
    If strFolder = "" Then
        Exit Sub
    End If
    
    pdfGeneratedCount = 0

    For Each wsA In wbA.Sheets
        If wsA.Visible = False Then
            wsA.Visible = True
            
            'replace spaces and periods in sheet name
            strName = Replace(wsA.Name, " ", "")
            strName = Replace(strName, ".", "_")
            
            'create name for savng file
            strFile = strName & ".pdf"
            strPathFile = strFolder & "\" & strFile
            
            wsA.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=strPathFile, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            
            wsA.Visible = False
            
            pdfGeneratedCount = pdfGeneratedCount + 1
        End If
    Next wsA
    
    'confirmation message with file info
    MsgBox Str(pdfGeneratedCount) & " PDF files are created"

End Sub
