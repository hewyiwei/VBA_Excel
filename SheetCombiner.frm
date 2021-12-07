VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Combine Sheets"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   OleObjectBlob   =   "SheetCombiner.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

For checkSelection = 0 To ListBox1.ListCount - 1

    If ListBox1.Selected(checkSelection) Then
    
        ListBox2.AddItem ListBox1.List(checkSelection)
    
    End If
    
Next checkSelection
    
End Sub

Private Sub CommandButton2_Click()

removedCounter = 0

For checkSelection = 0 To (ListBox2.ListCount - 1)

    If ListBox2.Selected(checkSelection - removedCounter) Then
    
        ListBox2.RemoveItem (checkSelection - removedCounter)
        
        removedCounter = removedCounter + 1
    
    End If
    
Next checkSelection

End Sub

Private Sub CommandButton3_Click()

Application.ScreenUpdating = False

noSheets = ThisWorkbook.Worksheets.Count

ThisWorkbook.Worksheets.Add after:=Worksheets(noSheets)

ActiveSheet.Name = "Combined" & " " & Date & " " & Time

newSheetPos = ThisWorkbook.Worksheets.Count

ThisWorkbook.Worksheets(ListBox2.List(0)).Activate

Rows(1).EntireRow.Copy

ThisWorkbook.Worksheets(newSheetPos).Activate

If OptionBox1.Value = True Then

    Rows(1).PasteSpecial xlPasteAll
    
ElseIf OptionBox2.Value = True Then

    Rows(1).PasteSpecial xlPasteValuesAndNumberFormats
    
End If

startingColumn = Alpha_Numeric(TextBox2.Value)

endingColumn = Alpha_Numeric(TextBox3.Value)

noOfColumns = endingColumn - startingColumn + 1

For Each Value In ListBox2

    ThisWorkbook.Worksheets(Value).Activate
    
    Dim rowsArr(startingColumn To endingColumn) As Long
    
    For columnCounter = startingColumn To endingColumn
    
        rowsArr(columnCounter) = Cells(Rows.Count, startingColumn).End(xlUp).Row
        
    Next columnCounter
    
    While checker < (noOfColumns - 1)
    
        checker = 0
    
        For rankingCounter = startingColumn To endingColumn
        
            
        
        
End Sub


Private Sub UserForm_Initialize()

Application.ScreenUpdating = False

ListBox1.MultiSelect = 1
ListBox2.MultiSelect = 1

noSheets = ThisWorkbook.Worksheets.Count

For sheetCounter = 1 To noSheets

    ThisWorkbook.Worksheets(sheetCounter).Activate

    ListBox1.AddItem ActiveSheet.Name
    
Next sheetCounter

End Sub

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
        'TextBox1.Value = Null
        'TextBox2.Value = Null
        'Call UserForm_Initialize
    End If
    
Next arrCounter2

If lengthOfNumber = 1 Then
    
    Alpha_Numeric = (columnArr(1) + columnArr(2) * 26 + columnArr(3) * 26 * 26)

Else

    Alpha_Numeric = (columnArr(3) + columnArr(2) * 26 + columnArr(1) * 26 * 26)
    
End If

End Function


