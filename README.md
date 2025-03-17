Sub PasteAndSplitAddressHyperlink()
    Dim ws As Worksheet
    Dim clipboardData As String
    Dim addressLines() As String
    Dim nextRow As Long
    Dim cityStateZip As String
    Dim i As Long
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name
    
    ' Get data from the clipboard
    clipboardData = GetClipboardText()
    
    ' Check if clipboard data is valid
    If clipboardData = "" Then
        MsgBox "No data found in the clipboard!", vbExclamation
        Exit Sub
    End If
    
    ' Debug: Print clipboard data
    Debug.Print "Clipboard Data: " & clipboardData
    
    ' Split the clipboard data into lines
    addressLines = Split(clipboardData, vbCrLf)
    
    ' Debug: Print address lines
    Debug.Print "Address Lines: " & Join(addressLines, " | ")
    
    ' Find the next empty row in Column N
    nextRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row + 1
    
    ' Assign Name (first line)
    ws.Cells(nextRow, "N").Value = Trim(addressLines(0)) ' Name
    Debug.Print "Name: " & ws.Cells(nextRow, "N").Value
    
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
            Debug.Print "Address 1: " & ws.Cells(nextRow, "P").Value
        ElseIf ws.Cells(nextRow, "Q").Value = "" Then
            ws.Cells(nextRow, "Q").Value = currentLine ' Address 2
            Debug.Print "Address 2: " & ws.Cells(nextRow, "Q").Value
        End If
    Next i
    
    ' Handle City, State, and Zip Code
    If cityStateZip <> "" Then
        ' Extract Zip Code (last 5 characters)
        ws.Cells(nextRow, "U").Value = Right(cityStateZip, 5) ' Zip Code
        Debug.Print "Zip Code: " & ws.Cells(nextRow, "U").Value
        
        ' Extract State (2 characters before the zip code)
        ws.Cells(nextRow, "T").Value = Mid(cityStateZip, Len(cityStateZip) - 7, 2) ' State
        Debug.Print "State: " & ws.Cells(nextRow, "T").Value
        
        ' Extract City (everything before the state and zip code)
        ws.Cells(nextRow, "S").Value = Trim(Left(cityStateZip, Len(cityStateZip) - 8)) ' City
        Debug.Print "City: " & ws.Cells(nextRow, "S").Value
    Else
        ' Handle cases where City, State, or Zip Code is missing
        ws.Cells(nextRow, "S").Value = "Invalid Format"
        ws.Cells(nextRow, "T").Value = "Invalid Format"
        ws.Cells(nextRow, "U").Value = "Invalid Format"
    End If
    
    MsgBox "Address pasted and split successfully!", vbInformation
End Sub

Function GetClipboardText() As String
    Dim objData As Object
    On Error Resume Next
    ' Use Microsoft Forms 2.0 Object Library to access clipboard
    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objData.GetFromClipboard
    GetClipboardText = objData.GetText
    On Error GoTo 0
End Function