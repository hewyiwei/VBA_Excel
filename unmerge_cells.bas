Attribute VB_Name = "unmerge_cells"
Sub undoMerged()
 
Application.ScreenUpdating = False
 
Dim rCount As Integer
Dim cCount As Integer
Dim rCountS As Integer
Dim cCountS As Integer
Dim subR As Integer
Dim subC As Integer
Dim subRs As Integer
Dim subCs As Integer
 
rCount = Selection.Rows.Count
cCount = Selection.Columns.Count
rCountS = Selection.Row
cCountS = Selection.Column
 
For cStarter = cCountS To (cCountS + cCount - 1)
 
    For rStarter = rCountS To (rCountS + rCount - 1)
    
        If (Cells(rStarter, cStarter).MergeCells = True) = True Then
            
            Cells(rStarter, cStarter).Select
            
            Cells(rStarter, cStarter).UnMerge
            
            subR = Selection.Rows.Count
            subC = Selection.Columns.Count
            subRs = Selection.Row
            subCs = Selection.Column
            
            Cells(subRs, subCs).Copy
            
            For cSubStarter = subCs To (subCs + subC - 1)
            
                For rSubStarter = subRs To (subRs + subR - 1)
                
                    Cells(rSubStarter, cSubStarter).PasteSpecial xlPasteValuesAndNumberFormats
                
                Next rSubStarter
                
            Next cSubStarter
            
        End If
        
    Next rStarter
    
Next cStarter
         
End Sub

