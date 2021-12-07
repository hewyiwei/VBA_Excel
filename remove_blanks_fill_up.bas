Attribute VB_Name = "remove_blanks_fill_up"
Sub removeBlanks()
 
sRow = Selection.Row
sColumn = Selection.Column
cRow = Selection.Rows.Count
cColumn = Selection.Columns.Count
 
arrCounter = 0
 
For columnCounter = sColumn To (sColumn + cColumn - 1)
 
    Dim arrValues() As Long
 
    ReDim arrValues(1 To cRow) As Long
 
    For rowCounter = sRow To (sRow + cRow - 1)
    
        If Not (IsEmpty(Cells(rowCounter, columnCounter).Value)) Then
        
            arrCounter = arrCounter + 1
            
            arrValues(arrCounter) = Cells(rowCounter, columnCounter).Value
            
            Cells(rowCounter, columnCounter).ClearContents
            
            Cells(rowCounter, columnCounter).ClearFormats
            
        End If
        
    Next rowCounter
    
    pasteRow = sRow
    
    For insertValues = 1 To arrCounter
        
        Cells(pasteRow, columnCounter).Value = arrValues(insertValues)
        
        pasteRow = pasteRow + 1
        
    Next insertValues
    
    arrCounter = 0
    
Next columnCounter
 
End Sub

