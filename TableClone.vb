Sub Macro1()
'
' Macro1 Macro
'
'
'
' TableCleaner Macro
'
'
    Dim dataTable As Table
    Dim sourceDocument As Document
    Set sourceDocument = ActiveDocument
    Dim UserChoice As String
    Dim QuestionToMessageBox As String
    Set dataTable = sourceDocument.Tables(2)
    QuestionToMessageBox = "Delete row with Text?"

    Dim targetRows() As Row
    ReDim targetRows(1 To 1) As Row
    
    With dataTable
        If .Uniform = False Then ' Checking table for merged columns
            For Each oRow In dataTable.Rows
                On Error GoTo ErrHandler
                cellText = oRow.Cells(2).Range.Text
                cellText = Trim(Left(cellText, Len(cellText) - 2))  'trim extra characters
                
                If Len(oRow.Cells(3).Range.Text) = 0 Then
                    cellText = "Header Text: " & vbCrLf & Trim(Left(oRow.Cells(1).Range.Text, Len(cellText) - 2))
                End If
                oRow.Select
                ActiveWindow.ScrollIntoView Selection.Range, True 'Scroll document to focus on selected row
                UserChoice = MsgBox(cellText, vbYesNoCancel, "Delete Row?")
                If UserChoice = vbYes Then
                    Set targetRows(UBound(targetRows)) = oRow
                    ReDim Preserve targetRows(1 To UBound(targetRows) + 1) As Row
                    'oRow.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightGreen
                ElseIf UserChoice = vbCancel Then
                    End
                End If
                
            Next oRow
        End If
    End With
    
    Confirmation = MsgBox("Are you sure?", vbYesNo, "Confirm?")
    If Confirmation = vbYes Then
        For Each targetRow In targetRows
            targetRow.Delete
        Next targetRow

	' Save file as new, The default locatioin is MyDocuments
        
        ActiveDocument.SaveAs2 FileName:="Updated_File.docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
        
    End If

ErrHandler:
   Select Case Err.Number ' Evaluate error number.
    Case 5941 ' "Cell has no text.
        Resume Next
    Case Else
     'ErrorResp = MsgBox(Err.Number, vbOK, "Error")
     Resume Next
    End Select

End Sub
