Sub PasteAndSplitAddressHyperlink()
    Dim ws As Worksheet
    Dim clipboardData As String
    Dim addressLines() As String
    Dim nextRow As Long
    Dim cityStateZip As String
    Dim i As Long
    Dim hyperlinkText As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Get data from the clipboard
    On Error Resume Next
    clipboardData = Application.ClipboardFormats(xlClipboardFormatText)
    On Error GoTo 0
    
    ' Check if clipboard data is valid
    If clipboardData = "" Then
        MsgBox "No data found in the clipboard!", vbExclamation
        Exit Sub
    End If
    
    ' Extract plain text from hyperlink
    hyperlinkText = GetPlainTextFromClipboard()
    
    ' Split the plain text data into lines
    addressLines = Split(hyperlinkText, vbCrLf)
    
    ' Find the next empty row in Column N
    nextRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row + 1
    
    ' Assign Name (first line)
    ws.Cells(nextRow, "N").Value = Trim(addressLines(0)) ' Name
    
    ' Loop through the remaining lines to extract Address 1, Address 2, and City/State/Zip
    For i = 1 To UBound(addressLines)
        Dim currentLine As String
        currentLine = Trim(addressLines(i))
        
        ' Check if the line contains City, State, and Zip Code
        If InStr(currentLine, " ") > 0 And Len(currentLine) > 5 Then
            ' Extract Zip Code (last 5 characters)
            If IsNumeric(Right(currentLine, 5)) Then
                cityStateZip = currentLine
                Exit For
            End If
        End If
        
        ' Assign Address 1 and Address 2
        If ws.Cells(nextRow, "P").Value = "" Then
            ws.Cells(nextRow, "P").Value = currentLine ' Address 1
        ElseIf ws.Cells(nextRow, "Q").Value = "" Then
            ws.Cells(nextRow, "Q").Value = currentLine ' Address 2
        End If
    Next i
    
    ' Handle City, State, and Zip Code
    If cityStateZip <> "" Then
        ' Extract Zip Code (last 5 characters)
        ws.Cells(nextRow, "U").Value = Right(cityStateZip, 5) ' Zip Code
        
        ' Extract State (2 characters before the zip code)
        ws.Cells(nextRow, "T").Value = Mid(cityStateZip, Len(cityStateZip) - 7, 2) ' State
        
        ' Extract City (everything before the state and zip code)
        ws.Cells(nextRow, "S").Value = Trim(Left(cityStateZip, Len(cityStateZip) - 8)) ' City
    Else
        ' Handle cases where City, State, or Zip Code is missing
        ws.Cells(nextRow, "S").Value = "Invalid Format"
        ws.Cells(nextRow, "T").Value = "Invalid Format"
        ws.Cells(nextRow, "U").Value = "Invalid Format"
    End If
    
    MsgBox "Address pasted and split successfully!", vbInformation
End Sub

Function GetPlainTextFromClipboard() As String
    Dim html As Object
    Dim clipboardData As String
    
    ' Get HTML data from the clipboard
    On Error Resume Next
    Set html = CreateObject("htmlfile")
    clipboardData = html.ParentWindow.ClipboardData.GetData("text")
    On Error GoTo 0
    
    ' If HTML data is not available, get plain text
    If clipboardData = "" Then
        clipboardData = Application.ClipboardFormats(xlClipboardFormatText)
    End If
    
    GetPlainTextFromClipboard = clipboardData
End Function