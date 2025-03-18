Sub SplitCityStateZip()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim cityStateZip As String
    Dim city As String
    Dim state As String
    Dim zipCode As String
    
    ' Set the worksheet
    Set ws = ActiveSheet ' Use the active sheet
    
    ' Get the selected cell
    Set selectedCell = Selection
    If selectedCell.Cells.Count > 1 Then
        MsgBox "Please select a single cell!", vbExclamation
        Exit Sub
    End If
    
    ' Get the value from the selected cell
    cityStateZip = Trim(selectedCell.Value)
    
    ' Check if the cell contains valid data
    If cityStateZip = "" Then
        MsgBox "Selected cell is empty!", vbExclamation
        Exit Sub
    End If
    
    ' Extract Zip Code (last 5 characters)
    zipCode = Right(cityStateZip, 5)
    
    ' Extract State (2 characters before the zip code)
    state = Mid(cityStateZip, Len(cityStateZip) - 7, 2)
    
    ' Extract City (everything before the state and zip code)
    city = Trim(Left(cityStateZip, Len(cityStateZip) - 8))
    
    ' Clean up extra spaces in the city name
    city = WorksheetFunction.Trim(city)
    
    ' Assign values to respective columns
    selectedCell.Offset(0, 1).Value = city   ' City (Column S)
    selectedCell.Offset(0, 2).Value = state  ' State (Column T)
    selectedCell.Offset(0, 3).Value = zipCode ' Zip Code (Column U)
    
    MsgBox "City, State, and Zip Code split successfully!", vbInformation
End Sub