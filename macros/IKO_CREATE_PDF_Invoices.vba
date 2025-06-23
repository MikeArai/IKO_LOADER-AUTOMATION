Sub IKO_CREATE_PDF_Invoices()

    Dim wsInfo As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim wbPath As String
    Dim pdfName As String
    Dim sheetsArray As Variant
    Dim sheetExists As Boolean
    Dim shtName As Variant

    Set wsInfo = ThisWorkbook.Sheets("Information")
    wbPath = ThisWorkbook.Path

    lastRow = wsInfo.Cells(wsInfo.Rows.Count, "AN").End(xlUp).Row

    For i = 1 To lastRow
        If Not IsEmpty(wsInfo.Cells(i, "AN")) Then
            sheetsArray = Split(wsInfo.Cells(i, "AN"), ",")

            For j = LBound(sheetsArray) To UBound(sheetsArray)
                sheetsArray(j) = Trim(sheetsArray(j))
            Next j

            sheetExists = True
            For Each shtName In sheetsArray
                If Not Evaluate("ISREF('" & shtName & "'!A1)") Then
                    sheetExists = False
                    Exit For
                End If
            Next shtName

            If sheetExists Then
                invoiceText = ThisWorkbook.Sheets(sheetsArray(0)).Range("G5").Value
                pdfName = wbPath & "\" & invoiceText & ".pdf"

                ThisWorkbook.Sheets(sheetsArray).Select

                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfName, _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, OpenAfterPublish:=False
            Else
                MsgBox "Group in row " & i & " has an invalid sheet name. Skipped.", vbExclamation
            End If
        End If
    Next i

    ThisWorkbook.Sheets("Information").Activate
    MsgBox "All groups exported successfully."
End Sub
