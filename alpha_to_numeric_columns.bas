Attribute VB_Name = "alpha_to_numeric_columns"
Function Alpha_Numeric(inputNumber)
 
lengthOfNumber = Len(Trim(inputNumber))
 
Dim columnArr(1 To 3) As Variant
 
For arrCounter = 1 To 3
 
    columnArr(arrCounter) = Mid(Trim(inputNumber), arrCounter, 1)
    
Next arrCounter
For arrCounter2 = 1 To 3
 
    If UCase(columnArr(arrCounter2)) = "A" Then
        columnArr(arrCounter2) = 1
    ElseIf UCase(columnArr(arrCounter2)) = "B" Then
        columnArr(arrCounter2) = 2
    ElseIf UCase(columnArr(arrCounter2)) = "C" Then
        columnArr(arrCounter2) = 3
    ElseIf UCase(columnArr(arrCounter2)) = "D" Then
        columnArr(arrCounter2) = 4
    ElseIf UCase(columnArr(arrCounter2)) = "E" Then
        columnArr(arrCounter2) = 5
    ElseIf UCase(columnArr(arrCounter2)) = "F" Then
        columnArr(arrCounter2) = 6
    ElseIf UCase(columnArr(arrCounter2)) = "G" Then
        columnArr(arrCounter2) = 7
    ElseIf UCase(columnArr(arrCounter2)) = "H" Then
        columnArr(arrCounter2) = 8
    ElseIf UCase(columnArr(arrCounter2)) = "I" Then
        columnArr(arrCounter2) = 9
    ElseIf UCase(columnArr(arrCounter2)) = "J" Then
        columnArr(arrCounter2) = 10
    ElseIf UCase(columnArr(arrCounter2)) = "K" Then
        columnArr(arrCounter2) = 11
    ElseIf UCase(columnArr(arrCounter2)) = "L" Then
        columnArr(arrCounter2) = 12
    ElseIf UCase(columnArr(arrCounter2)) = "M" Then
        columnArr(arrCounter2) = 13
    ElseIf UCase(columnArr(arrCounter2)) = "N" Then
        columnArr(arrCounter2) = 14
    ElseIf UCase(columnArr(arrCounter2)) = "O" Then
        columnArr(arrCounter2) = 15
    ElseIf UCase(columnArr(arrCounter2)) = "P" Then
        columnArr(arrCounter2) = 16
    ElseIf UCase(columnArr(arrCounter2)) = "Q" Then
        columnArr(arrCounter2) = 17
    ElseIf UCase(columnArr(arrCounter2)) = "R" Then
        columnArr(arrCounter2) = 18
    ElseIf UCase(columnArr(arrCounter2)) = "S" Then
        columnArr(arrCounter2) = 19
    ElseIf UCase(columnArr(arrCounter2)) = "T" Then
        columnArr(arrCounter2) = 20
    ElseIf UCase(columnArr(arrCounter2)) = "U" Then
        columnArr(arrCounter2) = 21
    ElseIf UCase(columnArr(arrCounter2)) = "V" Then
        columnArr(arrCounter2) = 22
    ElseIf UCase(columnArr(arrCounter2)) = "W" Then
        columnArr(arrCounter2) = 23
    ElseIf UCase(columnArr(arrCounter2)) = "X" Then
        columnArr(arrCounter2) = 24
    ElseIf UCase(columnArr(arrCounter2)) = "Y" Then
        columnArr(arrCounter2) = 25
    ElseIf UCase(columnArr(arrCounter2)) = "Z" Then
        columnArr(arrCounter2) = 26
    ElseIf UCase(columnArr(arrCounter2)) = "" Then
        columnArr(arrCounter2) = 0
    Else
        MsgBox "Invalid input, please re-try.", vbOKOnly + vbCritical, "Error!"
    End If
    
Next arrCounter2
 
If lengthOfNumber = 1 Then
    
    Alpha_Numeric = (columnArr(1) + columnArr(2) * 26 + columnArr(3) * 26 * 26)
 
Else
 
    Alpha_Numeric = (columnArr(3) + columnArr(2) * 26 + columnArr(1) * 26 * 26)
    
End If
 
End Function
