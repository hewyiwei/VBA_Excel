Attribute VB_Name = "fill_down"
Sub fillDownwards()
 
full_Column = 6
column_To_Be_Filled = 1
first_Row_Of_Data = 2
 
lastRow = Cells(Rows.Count, full_Column).End(xlUp).Row
 
subCounter = 0
 
For rowCounter = first_Row_Of_Data To lastRow
 
    If IsEmpty(Cells(rowCounter, column_To_Be_Filled).Value) Then
    
        Cells(rowCounter, column_To_Be_Filled).Value = Cells(rowCounter - 1, column_To_Be_Filled).Value
        
    End If
    
Next rowCounter
 
End Sub

