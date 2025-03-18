Sub SplitAndMoveAddress_RowWise()
    Dim ws As Worksheet
    Dim selectedRange As Range
    Dim addrLines() As String
    Dim cityParts() As String
    Dim namePart As String
    Dim addr1Part As String
    Dim addr2Part As String
    Dim cityPart As String
    Dim statePart As String
    Dim pincodePart As String
    Dim totalLines As Integer
    Dim i As Integer
    Dim startRow As Long
    
    ' Set active worksheet
    Set ws = ActiveSheet

    ' Ensure user selects multiple rows in Column N
    If Selection.Columns.Count = 1 And Selection.Column = 14 Then ' Column N = 14
        Set selectedRange = Selection
        
        ' Get the first row number of the selection
        startRow = selectedRange.Row
        totalLines = selectedRange.Rows.Count
        
        ' Store the selected values into an array
        ReDim addrLines(1 To totalLines)
        
        For i = 1 To totalLines
            addrLines(i) = Trim(ws.Cells(startRow + i - 1, 14).Value)
        Next i
        
        ' Ensure there are at least 2 lines (Name + Address)
        If totalLines >= 2 Then
            namePart = addrLines(1) ' First row is always the name
            
            ' Handle Address 1 & 2
            If totalLines = 3 Then
                addr2Part = addrLines(2) ' If only one address line, move to Q
            ElseIf totalLines >= 4 Then
                addr1Part = addrLines(2) ' Address Line 1 (Column P)
                addr2Part = addrLines(3) ' Address Line 2 (Column Q)
            End If
            
            ' Extract last line for City, State, and Pincode
            cityParts = Split(addrLines(totalLines), " ")
            pincodePart = cityParts(UBound(cityParts)) ' Last word = Pincode
            statePart = cityParts(UBound(cityParts) - 1) ' Second last word = State
            
            ' Extract City (everything before State)
            cityPart = ""
            For i = 0 To UBound(cityParts) - 2
                cityPart = cityPart & cityParts(i) & " "
            Next i
            cityPart = Trim(cityPart) ' Remove extra space
            
            ' Move data to respective columns
            ws.Cells(startRow, 16).Value = addr1Part  ' Column P (Address 1)
            ws.Cells(startRow, 17).Value = addr2Part  ' Column Q (Address 2)
            ws.Cells(startRow, 19).Value = cityPart   ' Column S (City)
            ws.Cells(startRow, 20).Value = statePart  ' Column T (State)
            ws.Cells(startRow, 21).Value = pincodePart ' Column U (Pincode)
            
            ' Clear original address data (keep name in Column N)
            ws.Range(ws.Cells(startRow + 1, 14), ws.Cells(startRow + totalLines - 1, 14)).ClearContents

            MsgBox "Address split and moved successfully (Row-wise)!", vbInformation
        Else
            MsgBox "Invalid selection! Ensure it follows the correct format.", vbExclamation
        End If
    Else
        MsgBox "Please select a valid address block in Column N.", vbExclamation
    End If
End Sub