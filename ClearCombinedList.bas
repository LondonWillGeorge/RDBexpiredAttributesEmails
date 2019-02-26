Attribute VB_Name = "ClearCombinedList"
Sub ClearList_Click()
    ' Clear the cells in the sheet where this button is, from Cell H1 to U10,000
    ActiveSheet.Range("H1:U10000").Clear
    
End Sub
