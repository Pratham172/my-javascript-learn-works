Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim selectedRange As Range
    Dim addrLines() As String
    Dim cityParts() As String
    Dim addr1Part As String
    Dim addr2Part As String
    Dim cityPart As String
    Dim statePart As String
    Dim pincodePart As String
    Dim totalLines As Integer
    Dim i As Integer
    Dim startRow As Long
    
    ' Set the active worksheet
    Set ws = Me
    
    ' Check if change occurred in Column Q
    If Not Intersect(Target, ws.Columns(17)) Is Nothing Then ' Column Q = 17
        Application.EnableEvents = False ' Prevent infinite loop
        
        ' Get the first row number of the change
        startRow = Target.Row
        totalLines = Target.Rows.Count
        
        ' Store values into an array
        ReDim addrLines(1 To totalLines)
        
        For i = 1 To totalLines
            addrLines(i) = Trim(ws.Cells(startRow + i - 1, 17).Value)
        Next i
        
        ' Ensure at least two lines exist (Address + City, State, Zip)
        If totalLines >= 2 Then
            addr1Part = addrLines(1) ' First row is Address Line 1
            
            ' If 3 rows, second row is Address 2; else, keep empty
            If totalLines = 3 Then
                addr2Part = addrLines(2) ' Address Line 2
            Else
                addr2Part = "" ' If missing, leave blank
            End If
            
            ' Extract last row for City, State, and Pincode
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
            ws.Cells(startRow, 16).Value = addr2Part  ' Column P (Address 2) - Empty if missing
            ws.Cells(startRow, 17).Value = addr1Part  ' Column Q (Address 1)
            ws.Cells(startRow, 19).Value = cityPart   ' Column S (City)
            ws.Cells(startRow, 20).Value = statePart  ' Column T (State)
            ws.Cells(startRow, 21).Value = pincodePart ' Column U (Pincode)
            
            ' Clear original address data in Column Q
            ws.Range(ws.Cells(startRow, 17), ws.Cells(startRow + totalLines - 1, 17)).ClearContents
        End If
        
        Application.EnableEvents = True ' Re-enable events
    End If
End Sub