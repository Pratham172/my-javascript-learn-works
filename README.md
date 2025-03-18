Sub CutAndSplitZipState_FromS()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim addrParts As Variant
    Dim cityPart As String
    Dim statePart As String
    Dim pincodePart As String
    Dim totalParts As Integer
    Dim i As Integer
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Check if the user has selected a single cell in Column S
    If Selection.Cells.Count = 1 And Selection.Column = 19 Then ' Column S = 19
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
            ' Extract City (everything before State) and update Column S
            cityPart = ""
            For i = 0 To totalParts - 2
                cityPart = cityPart & addrParts(i) & " "
            Next i
            cityPart = Trim(cityPart) ' Remove trailing space
            
            ' Move data to respective columns
            ws.Cells(selectedCell.Row, 20).Value = statePart ' Column T (State)
            ws.Cells(selectedCell.Row, 21).Value = pincodePart ' Column U (Pincode)
            
            ' Update Column S with only City (cut out Pincode & State)
            selectedCell.Value = cityPart
            
            MsgBox "State and Pincode moved successfully!", vbInformation
        Else
            MsgBox "Invalid address format! Ensure it follows 'City State Pincode'.", vbExclamation
        End If
    Else
        MsgBox "Please select a single address cell in Column S.", vbExclamation
    End If
End Sub