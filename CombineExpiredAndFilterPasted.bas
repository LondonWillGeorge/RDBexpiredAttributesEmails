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
        
        ' "Mail List" sheet is generated from VBA, but even after manually deleting,
        ' it can still be there even if not visible (bit strange this but using breakpoint confirmed), so explicitly delete here.
        DeleteMailList
 
        For Each Ws In Worksheets
            
            ' sheet index starts at 1, not zero!
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
                inactiveNames = inactiveNames + "Cell " + CStr(a) + " i.e. " + Cells(a, 7) + vbCrLf
                inactiveCount = inactiveCount + 1
                Exit For
            End If
        End If
                
        Dim thisData As matchData
        thisData = MatchNHSnames(Cells(a, 7), Cells(b, 8))
        
        ' running this loop multiple times with each name from the comparison list, yes, because outer array same...
        If thisData.matched = True Then
            ' boolean flag this value of a as matched
            matchedFlag = True
            bestMatch(0, matchCount) = b
            ' so second item in first index of array holds match value to compare
            bestMatch(1, matchCount) = thisData.fuzz1 + thisData.fuzz2
            matchCount = matchCount + 1
            ReDim Preserve bestMatch(1, 0 To matchCount)
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
            End If
        Next c
        ' add a blank item in the resArray for next match name
        resLength = resLength + 1
        ReDim Preserve resArray(0 To resLength)
    End If

Next a

' resArray now holds all the correct row numbers
' But want to remove duplicate names from resArray in case there are any
' Sort by integer value, then remove any consecutively repeated integers
DeleteMailList

' Create new sheet "Mail List" 1 place to the right of START sheet
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

' print number of inactive candidates in Message Box, TODO: keep msg box, also save this to a file
MsgBox (inactiveNames + "Total Number = " + CStr(inactiveCount))


End Sub

Private Function DeleteMailList()
    For Each Sheet In ActiveWorkbook.Worksheets
        If Sheet.Name = "Mail List" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
End Function
