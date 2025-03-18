Sub SplitAndMoveAddress_Vertical()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim addrParts As Variant
    Dim cityPart As String
    Dim statePart As String
    Dim pincodePart As String
    Dim totalParts As Integer
    Dim i As Integer
    Dim nextRow As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Check if the user has selected a single cell in Column N
    If Selection.Cells.Count = 1 And Selection.Column = 14 Then ' Column N = 14
        Set selectedCell = Selection
        
        ' Split the address by spaces
        addrParts = Split(Trim(selectedCell.Value), " ")
        totalParts = UBound(addrParts) ' Get the last index of the array
        
        ' Ensure there are at least two parts (State & Pincode)
        If totalParts >= 1 Then
            ' Extract Pincode (last word)
            pincodePart = addrParts(totalParts)
            ' Extract State (second last word)
            statePart = addrParts(totalParts - 1)
            ' Extract City (everything before State)
            cityPart = ""
            For i = 0 To totalParts - 2
                cityPart = cityPart & addrParts(i) & " "
            Next i
            cityPart = Trim(cityPart) ' Remove trailing space
            
            ' Find the next empty row in Column S (City column)
            nextRow = ws.Cells(ws.Rows.Count, 19).End(xlUp).Row + 1 ' Column S = 19
            
            ' Move data to respective columns (stacked downwards)
            ws.Cells(nextRow, 19).Value = cityPart  ' Column S (City)
            ws.Cells(nextRow, 20).Value = statePart ' Column T (State)
            ws.Cells(nextRow, 21).Value = pincodePart ' Column U (Pincode)
            
            ' Clear the original cell
            selectedCell.ClearContents
            
            MsgBox "Address split and moved successfully!", vbInformation
        Else
            MsgBox "Invalid address format! Ensure it follows 'City State Pincode'.", vbExclamation
        End If
    Else
        MsgBox "Please select a single address cell in Column N.", vbExclamation