Sub SplitAndMoveAddress()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim addrParts As Variant
    Dim cityPart As String
    Dim statePart As String
    Dim pincodePart As String
    Dim totalParts As Integer
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Check if the user has selected a single cell in Column N
    If Selection.Cells.Count = 1 And Selection.Column = 14 Then ' Column N = 14
        Set selectedCell = Selection
        
        ' Split the address by spaces
        addrParts = Split(selectedCell.Value, " ")
        totalParts = UBound(addrParts) ' Zero-based index
        
        ' Ensure there are at least two parts (State & Pincode)
        If totalParts >= 1 Then
            ' Extract Pincode (last word)
            pincodePart = Trim(addrParts(totalParts))
            ' Extract State (second last word)
            statePart = Trim(addrParts(totalParts - 1))
            ' Everything left is City (joins all words before the state)
            If totalParts > 1 Then
                cityPart = Join(Application.Index(addrParts, 0, 0 To totalParts - 2), " ")
            Else
                cityPart = "N/A" ' If no city is present, mark as N/A
            End If
            
            ' Move data to respective columns
            ws.Cells(selectedCell.Row, 19).Value = cityPart  ' Column S (City)
            ws.Cells(selectedCell.Row, 20).Value = statePart ' Column T (State)
            ws.Cells(selectedCell.Row, 21).Value = pincodePart ' Column U (Pincode)
            
            ' Clear the original cell
            selectedCell.ClearContents
            
            MsgBox "Address split and moved successfully!", vbInformation
        Else
            MsgBox "Invalid address format! Ensure it follows 'City State Pincode'.", vbExclamation
        End If
    Else
        MsgBox "Please select a single address cell in Column N.", vbExclamation
    End If
End Sub