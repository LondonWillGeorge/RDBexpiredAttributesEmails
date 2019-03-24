Sub ClearList_Click()
    ' Clear the cells in the sheet where this button is, from Cell H1 to U10,000
    ' which should be all the cells with candidate names and missing certificate names
    ActiveSheet.Range("H1:U10000").Clear
    
End Sub
