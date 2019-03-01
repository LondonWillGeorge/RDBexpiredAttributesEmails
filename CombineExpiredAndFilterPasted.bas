Attribute VB_Name = "CombineExpiredAndFilterPasted"
Sub CombineExpired_Click()
        
        Dim Ws As Worksheet
        Dim iLastRow As Integer
        Dim i As Integer
        Dim j As Integer
        
        Dim WsNames(20) As String 'fixed size array up to 21 worksheets fine
        ' populate array
        
        Dim currLength As Integer
        
        Dim X As Long
        X = 0
        ' ReDim Preserve namesArray(X)
        
        ' Dim namesDict As New Scripting.Dictionary
        
        Dim nameString As String
        nameString = ""
        
        Dim allTheNames As String
        allTheNames = ""
        
        Dim colorsArray As Variant
        Dim thisCol As Variant
        
        Dim colorIndex As Integer
       
        colorsArray = Array(Array(0, 0, 255), Array(0, 200, 0), Array(255, 0, 0), Array(255, 0, 255), Array(128, 0, 128), Array(128, 0, 0), Array(0, 150, 150))

        Dim namesArray() As Variant
        Dim tempArray() As Variant
        ' put first element of array as integer colour. If first fuzzy match, then change namesArray(X)(0) value
        ' if 2nd or more, change to colour of element it matches
        ' Redim outerArray to incremented (or ubound + 1?) each time name added to list
        ReDim Preserve namesArray(0 To 2)
 
        namesArray(0) = Array("", "")
        
        sheetInd = 1
 
        For Each Ws In Worksheets
            
            ' sheet index starts at 1, not zero!
            ' Worksheets.count returns number sheets
            
            If Ws.Name <> "START" Then
                WsNames(sheetInd) = Ws.Name
                sheetInd = sheetInd + 1
      
                iLastRow = Ws.Range("a10000").End(xlUp).Row
                
                ' MsgBox (Ws.Name + " last row number " + CStr(iLastRow) + ". index is " + _
                CStr(Ws.Index) + " and now worksheets number is " + CStr(Worksheets.Count))
                
                ' Append name if not in list
                For i = 2 To iLastRow
                    ' if Ws.Cells(i,1) not in array already...
                    
                    ' iterate over current final array
                    For j = 0 To X
                        ' trim both of them to be safe
                        ' if there IS a match, then add this sheet name to namesArray(j), which is itself a 1-D array
                        ' then exit loop;
                        ' otherwise add the name to namesArray
                        If Trim(namesArray(j)(0)) = Trim(Ws.Cells(i, 1)) Then
                            ' add this sheet parameter to array
                            currLength = UBound(namesArray(j))
                            ReDim tempArray(0 To currLength) As Variant
    
                            tempArray = namesArray(j)
                            ReDim Preserve tempArray(0 To currLength + 1)
                            tempArray(currLength + 1) = CStr(Ws.Name)
                            namesArray(j) = tempArray
                            Exit For
                        
                        End If
                        
                        If j = X Then
                            'should have reached end with no match
                            ' add current name to array and grow array by 1 to accommodate next name
                            namesArray(X) = Array(CStr(Ws.Cells(i, 1)), Ws.Name)
                            X = X + 1
                            ReDim Preserve namesArray(0 To X)
                            ' Just to be safe and clean, clear the new, last value in the array
                            namesArray(X) = Array("", "")
                        End If
                    Next j

                Next i
                
                ' X is total number of names now
             
            End If
            
        Next Ws
        
        
        ' Cells object takes row then column!
        Dim cellIndex As Integer
        ActiveSheet.Name = "START"
        namesArray = SortArrayAtoZ(namesArray)
        For k = 0 To UBound(namesArray)
            For l = 0 To UBound(namesArray(k))
                cellIndex = l + 8
                ActiveSheet.Cells(k + 1, cellIndex) = namesArray(k)(l)
            Next l
        Next k
   
End Sub

Sub FilterPastedList_Click()
' FuzzyPercent "Joanna Louise", "Lou Joan" = 0.189 , and if (Joan) is in brackets score is same
' Lou Joan, Joanna Louise is 0.26, Joan, Joanna Louise is 0.35, Josephina, Joanna Louise is 0.16
' For FuzzyPercent surnames: Baked, Baik is 0.4, Hari, Harrhy is 0.45, Archenna, Chenna is 0.25 but not likely i think!
' Jake, Jain is 0.375, Delavega, Vega is 0.25, Vega, Delavega is 0.39!
' --> Put LONGER name as String2
' --> Do 0.2 or above on first part of name, and 0.35 on surname OR 0.5 on whole name

' Resray is Result Array
' C is Compare list array
' O is Original combined list array

' *** Maybe quicker just sort array at end remove duplicates!!!
' instead of below, since not many double matches surely? Then can filter out all the duplicate row numbers at end, use sorting?
' Have array alpha sort function already, can adjust for integer sort, then iterate if next integer same then delete
' Or Could use collection and add row number as the key each time, incrementing Item as integer number by 1
' since Collection.contains method works on the key, or use Dictionary

' O is definitive list to filter, final array comes from O

' **** Outer loop on Compare List C then Inner Loop on Original list O
' MISS NOW: Check if current row number O(Y) already in Resray, if yes skip below
' CHECK MATCH: if C(X) matches O(Y), append O(Y) row number to Resray, just append Y in fact
' MISS NOW: ** OR remove O(Y) from O array instead of checking inclusion, since it only needs to be once in Resray.
' Keep in inner O loop, just for edge cases, checking other names for match

' to iterate each resray row, check trimmed cell to right blank, exit if blank, or add to 2ndary array if not

Dim resArray() As Integer
ReDim Preserve resArray(0 To 0)
resLength = 0

compLen = ActiveSheet.Range("g10000").End(xlUp).Row
attLen = ActiveSheet.Range("h10000").End(xlUp).Row

tempName = ""
' For names in G column, remove (RGN (HCA (RMN right half text if there,
' as this is not part of the name. This makes fuzzy matching easier as names
' are then as similar as possible before the match loop runs.
' Do this here once only, not in any function called from loop
For crunchoff = 2 To compLen


    tempName = Trim(Cells(crunchoff, 7))
    tempName = Split(tempName, "(RGN")(0)
    tempName = Split(tempName, "(RMN")(0)
    tempName = Split(tempName, "(HCA")(0)
    
    ' Format cell here if time, alignment and colours
    Cells(crunchoff, 7) = Trim(tempName)
    tempName = ""
    
Next crunchoff

inactiveNames = "The Inactive Candidates List as follows are not included:" + vbCrLf
inactiveCount = 0

' For a = 2 To compLen
' TRY TESTING with first x candidates in the list to save time
For a = 2 To compLen
    
    ' Split off surname and rest of front name for compLen here
    ' check for all the blue colors here if index matches

    compName = Trim(Cells(a, 7)) ' 7 is column G
    
    compSurname = UBound(Split(compName, " "))
    
    'bestMatch() array holds all true matches for each name in complist
    ' then select best match from this array to add
    
    ' TODO: change type matchData to save match values in the array
    Dim bestMatch() As Integer
    
    ReDim Preserve bestMatch(1, 0 To 0)
    matchCount = 0
    
    ' flag to record whether each name in comp list has a match or not
    matchedFlag = False
    
    For b = 2 To attLen
        
        cI = Cells(a, 7).Font.colorIndex

        If b = 2 Then
        ' if compLen cell is light blue colour, then name is inactive, so don't check
        ' use b=2 condition to only check once in this loop
            If cI = 33 Or cI = 8 Or cI = 20 Or cI = 28 Or cI = 34 Or cI = 41 Or cI = 42 Then
                
                ' Debug.Print ("Cell " + CStr(a) + " i.e. " + Cells(a, 7) + " is light blue, therefore inactive candidate, so we skip it")
                
                inactiveNames = inactiveNames + "Cell " + CStr(a) + " i.e. " + Cells(a, 7) + vbCrLf
                inactiveCount = inactiveCount + 1
                ' TODO: colour the name blue in H bit tricky as we don't have right H row at this point!
                Exit For
            End If
        End If
        

        
        ' TODO: actually within the inner b loop, we only want to match 1 candidate?
        ' add .fuzz1 and .fuzz2 values, save as 2nd item in sub-array, then select higher score as the match
        
        Dim thisData As matchData
        thisData = MatchNHSnames(Cells(a, 7), Cells(b, 8))
        
        ' running this loop multiple times with each name from the comparison list, yes, because outer array same...
        If thisData.matched = True Then
            ' to track simply which candidate, append row number of Column H to the resArray
            
            ' boolean flag this value of a as matched?
            matchedFlag = True

            bestMatch(0, matchCount) = b
            ' so second item in first index of array holds match value to compare
            bestMatch(1, matchCount) = thisData.fuzz1 + thisData.fuzz2
            matchCount = matchCount + 1
            ReDim Preserve bestMatch(1, 0 To matchCount)
            
'            resArray(resLength) = b
'            resLength = resLength + 1
'            ReDim Preserve resArray(0 To resLength)
        End If
    
    Next b

    ' at end of b loop here, decide which is best match by iterating over, resetting to highest each time
    ' remember bestMatch() dies when next a is iterated
    
    ' Only add item number a to the final array if there has been at least one match above
    ' the matchedFlag boolean keeps track of this
    If matchedFlag = True Then
        resArray(resLength) = bestMatch(0, 0)
        
        matchValue = bestMatch(1, 0)
        For c = 0 To matchCount
            If bestMatch(1, c) > matchValue Then
                resArray(resLength) = bestMatch(0, c)
                matchValue = bestMatch(1, c)
'                Debug.Print (matchValue)
'                Debug.Print (resArray(resLength))
            End If
        Next c
        
        ' add next blank space in the resArray for next match name
        resLength = resLength + 1
        ReDim Preserve resArray(0 To resLength)
    End If

Next a

' resArray now holds all the correct row numbers

' But want to remove duplicate names from resArray in case there are any
' Sort by integer value, then remove any consecutively repeated integers
' ** sorting to remove duplicates bit tricky still with an array! should have used dictionary perhaps

' TODO: test this below to delete existing Mail List sheet if there
For Each Sheet In ActiveWorkbook.Worksheets
     If Sheet.Name = "Mail List" Then
     Application.DisplayAlerts = False
     Sheet.Delete
     Application.DisplayAlerts = True
     End If
Next Sheet

' Create new sheet "Mail List" 1 place to the right of DBS sheet
' Populate Column C with final name list
With ThisWorkbook
    .Sheets.Add(After:=.Sheets("START")).Name = "Mail List"
End With
Worksheets("Mail List").Activate
    

' Set column width 30 to accomodate email addresses and wrap text, set height whole sheet
ActiveSheet.Range("A:D").ColumnWidth = 30
ActiveSheet.Cells.WrapText = True
ActiveSheet.Cells.RowHeight = 30
' Set height top row 100 to accomodate Emails button
ActiveSheet.Rows(1).RowHeight = 100
' make top row green colour
ActiveSheet.Rows(1).Interior.colorIndex = 43
ActiveSheet.Cells(1, 4) = "Final Names to Email"


' Create Emails button in the new sheet
  Dim btn As Button
  Application.ScreenUpdating = False
  Dim t As Range

    Set t = ActiveSheet.Range(Cells(1, 1), Cells(1, 2))
    Set btn = ActiveSheet.Buttons.Add(t.Left + 10, t.Top + 10, (t.Width * 0.8), (t.Height * 0.8))
    With btn
      .OnAction = "FilterEmailList_Click"
      .Caption = "Paste List with Names Column A and Emails in Column B, then Click Here!"
      .Name = "Filter_Emails"
    End With

  Application.ScreenUpdating = True



' This populates Column D of the new sheet with the final names
For resCount = 0 To resLength - 1
    Worksheets("Mail List").Cells(resCount + 2, 4) = Worksheets("START").Cells(resArray(resCount), 8)
    ' add attribs to the right, similar code to Create Emails sub
    For colNum = 9 To 20
    ' iterate to right until empty, add the relevant cell attribute from DBS sheet to this sheet
        
        cellString = Trim(Worksheets("START").Cells(resArray(resCount), colNum))
        
        If cellString = "" Then
            Exit For
        End If

        Worksheets("Mail List").Cells(resCount + 2, colNum - 4) = cellString

    Next colNum
    
Next resCount

' print number of inactive candidates to console

' Debug.Print (CStr(inactiveCount) + " inactive names total in list")

' debugging, need resLength - 1 as array begins at 0
'resString = "Candidate row numbers: "
'For resCount = 0 To resLength - 1
'    resString = resString + CStr(resArray(resCount)) + ", "
'Next resCount
'
'Debug.Print (resString)

MsgBox (inactiveNames + "Total Number = " + CStr(inactiveCount))


End Sub
