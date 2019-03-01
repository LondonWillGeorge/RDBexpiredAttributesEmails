Attribute VB_Name = "FilterEmailListCreateEmails"
' May need to End 1st word task manually in Task Manager upon running the application

Sub FilterEmailList_Click()

' set column width A to D
ActiveSheet.Range("A:D").ColumnWidth = 30

' Split off the Payroll string from beginning of name list, if this has such numbers/letters
' If the namelist is later 'pure' just names, this part shouldn't matter as it just checks the beginning for
' (example) "PAY" characters which Payroll numbers start with
' ***********
' 6 Feb 2019 Will's comment: actually using dictionary no advantage to using array in the end, because of fuzzy match problem!
' Split Column A names row by row, adding each to Dictionary mailDict as we go
' Column A Dictionary with name as Key, row number as Value
' Iterate over Column D outer loop
' Iterate over whole of Column A inner loop, find best fuzzy match for current Column D row
' Populate Column C row with B Column of the best match row
' If Column A is a dictionary, then can remove Column A row match after each iteration
' Next Column D row

Dim mailDict As Scripting.Dictionary
Set mailDict = New Scripting.Dictionary

emailsLen = ActiveSheet.Range("a10000").End(xlUp).Row

' populate dictionary type with all the names in this column
For ind = 2 To emailsLen
    tempName = ""
    ' If the names begin with payroll number PAY, then space then first name then second name, split off first part
    If Left(Cells(ind, 1), 2) = "PAY" Then

        For a = 1 To UBound(Split(Cells(ind, 1)))
            tempName = tempName + Split(Cells(ind, 1))(a) + " "
        Next a
        tempName = Trim(tempName)
        Cells(ind, 1) = tempName
    End If
    
    mailDict.Add Cells(ind, 1), ind

Next ind

Debug.Print ("OK, here comes the dictionary:")

Dim key As Variant
For Each key In mailDict.Keys
    'check match, remove key
    Debug.Print key, mailDict(key)
Next key

finalsLen = ActiveSheet.Range("d10000").End(xlUp).Row

' iterate over ColumnD
For ind2 = 2 To finalsLen

    matchedFlag = False
    matchCount = 0
    
    Dim bestMatch() As Integer
    ' want string?
    ReDim Preserve bestMatch(1, 0 To 0)

    ' save the item integer value to bestMatch(0, x) not the key which is string wrong type!
    For Each key In mailDict.Keys
    'check best match
        Dim thisData As matchData
        thisData = MatchNHSnames(Cells(ind2, 4), key)
        
        If thisData.matched = True Then

            matchedFlag = True

            bestMatch(0, matchCount) = mailDict(key)
            ' so second item in first index of array holds match value to compare
            bestMatch(1, matchCount) = thisData.fuzz1 + thisData.fuzz2
            matchCount = matchCount + 1
            ReDim Preserve bestMatch(1, 0 To matchCount)

        End If

    Next key
    ' decide best match key
    ' populate Column C row with correct email
    ' remove best match key
    If matchedFlag = True Then
        matchItem = bestMatch(0, 0)
        
        matchValue = bestMatch(1, 0)
        For c = 0 To matchCount
            If bestMatch(1, c) > matchValue Then
                matchItem = bestMatch(0, c)
                matchValue = bestMatch(1, c)
            End If
        Next c
        'Now decided final matchKey
        'row with email is mailDict(matchKey)
        rowNo = matchItem
        Cells(ind2, 3) = Cells(rowNo, 2)
        
    End If

Next ind2

  Dim btn As Button
  Application.ScreenUpdating = False
  Dim t As Range
    ' specify create button in Mail List sheet not the active sheet, in case wrong sheet is active
    Set t = Worksheets("Mail List").Range(Cells(1, 3), Cells(1, 3))
    Set btn = ActiveSheet.Buttons.Add(t.Left + 10, t.Top + 10, (t.Width * 0.8), (t.Height * 0.8))
    With btn
      .OnAction = "CreateEmails_Click"
      .Caption = "Create The Emails"
      .Name = "Create_Emails"
    End With

  Application.ScreenUpdating = True



End Sub


'Email generation base code from: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016
Sub CreateEmails_Click()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range

    Application.ScreenUpdating = False
    Set OutApp = CreateObject("Outlook.Application")
    
    On Error GoTo cleanup

    errorList = ""

    If Cells(2, 3) <> "" Then
    
    Dim messagePath As String: messagePath = ThisWorkbook.path & "\MessageText.docx"
    
    Dim paragArray() As String
    ReDim Preserve paragArray(0 To 0)
    
    Dim objWordMessage As Object: Set objWordMessage = CreateObject("Word.Application")
    ' objWordMessage.Visible = False
    objWordMessage.Application.DisplayAlerts = False
    objWordMessage.Application.ScreenUpdating = False
    
    Dim msgDoc As Object: Set msgDoc = objWordMessage.Documents.Open(messagePath, Visible = False, ReadOnlyRecommended = False)
    ' check path and show mgbox if path has no file

    Set colParagraphs = msgDoc.Paragraphs
    paraCount = 0
    For Each objParagraph In colParagraphs
    
        ' Debug.Print ("parag text is: " + objParagraph.Range.text)
        lineText = Trim(objParagraph.Range.text)
        If Trim(lineText) <> "" Then ' Trimming again here still gives blank items between!
           ' add to the text string array here
           paragArray(paraCount) = lineText
           paraCount = paraCount + 1
           ReDim Preserve paragArray(0 To paraCount)
        End If
    Next
    
    msgDoc.Close
    For Each cell In ActiveSheet.Columns("C").Cells.SpecialCells(xlCellTypeConstants)
        
        ' *** FOR TESTING, STOP AT ROW 4, COMMENT OUT LATER!!!! ****
        If cell.Row > 4 Then
            Exit For
        End If
        
        ' save the row number of the current candidate for use in this loop
        RowNum = cell.Row
        full_name = Cells(RowNum, 4)
        ' Get first name only from name string, so we can address them by first name.
        firstName = Split(Cells(RowNum, 4), " ")(0)
    
        If cell.Value Like "?*@?*.?*" Then
            Set OutMail = OutApp.CreateItem(0)
            On Error Resume Next
            With OutMail
                .To = cell.Value
                .Subject = "Urgent: Please Check your xxxx document" ' say compliance documents?

                .body = "Dear " + firstName + "," + vbCrLf + vbCrLf + "We'd like to thank you again for your valued " + _
                "contribution to (company name). We really want to continue to offer you as many shifts as we can, " + _
                "so would be very grateful if you could check the following documents, which we think are out of date now for you."

                proofs = 0
                refs = 0
                dvla = False

                For colNum = 5 To 21
                ' iterate to right until empty, add the text to .Body
                    If Trim(Cells(RowNum, colNum)) = "" Then
                        Exit For
                    ElseIf colNum = 5 Then
                        .Subject = .Subject + "s"
                    End If

                    ' add text
                    cellValue = Cells(RowNum, colNum)
                    docName = ""

                    ' Select case here would be best I think, set standard text string variable,
                    ' and add this in appropriate place in select case text for each attribute case.

                    Select Case cellValue
                        Case "DBS"
                            docName = "DBS (Disclosure and Barring Service) Enhanced certificate"
                        Case "FTW"
                            docName = "Fitness to Work certificate"
                        Case "Appraisal"
                            docName = "Appraisal document"
                        Case "BLS"
                            docName = "Basic Life Support or Immediate Life Support Training Certificate"
                        Case "NMC"
                            docName = "NMC Pin Check fee expiry document"
                        Case "Manual Handling"
                            docName = "Manual Handling (Moving & Handling) Training Certificate"
                        Case "Proof Address1"
                            proofs = proofs + 1
                        Case "Proof Address2"
                            proofs = proofs + 1
                        Case "Ref1"
                            refs = refs + 1
                        Case "Ref2"
                            refs = refs + 1
                        Case "EU Passport"
                            docName = "Passport or National ID Card"
                        Case "ROW Passport"
                            docName = "Passport"
                        Case "UK Passport"
                            docName = "UK Passport"
                        Case "DVLA"
                            dvla = True
                            docName = "DVLA"
                        Case "Visa"
                            docName = "UK visa or Residence Permit"

                        Case "ID Badge"
                            docName = ""
                            errorList = errorList + full_name + " has an out of date ID Badge on RDB." + vbCrLf

                        Case Else
                            docName = ""
                            errorList = errorList + full_name + " has an error with a document listed as " + cellValue + vbCrLf
                    End Select

                    ' check for permission to work combination invalidities, and generate corresponding messages.

                    ' .Body = .Body + vbCrLf + "    Your " + Cells(rowNum, colNum) + " is either due to expire soon or has expired. Please could you renew this and email us a clear photocopy as soon as possible, ideally within the next week, so we can continue to offer you shifts."
                    If docName <> "" Then
                        .body = .body + vbCrLf + "    Your " + docName + " is either due to expire very soon, or has expired. Please could you renew this and email us a clear photocopy as soon as possible, ideally within the next week, so we can continue to offer you shifts."
                    End If

                Next colNum

                If proofs = 1 Then
                    .body = .body + vbCrLf + "You have one proof of address missing. Please email us one of: a council tax letter, a bank statement, a current UK Driving License (DVLA)."
                ElseIf proofs = 2 Then
                    .body = .body + vbCrLf + "You have 2 proofs of address missing. Please email us two different documents from this list: a council tax letter, a bank statement, a current UK Driving License (DVLA)."
                End If

                If refs = 1 Then
                    .body = .body + vbCrLf + "You have one work reference missing or nearly out of date. We would be very grateful if you could ask any one of your current supervisors or managers to give you a reference and then email it back to us as soon as possible."
                ElseIf refs = 2 Then
                    .body = .body + vbCrLf + "You have both (2) of your work references missing or nearly out of date. We would be very grateful if you could ask any two of your current supervisors or managers to give you a reference and then email it back to us as soon as possible."
                End If

                .body = .body + vbCrLf + " In case you have a question about one of these documents, feel free to just email us in return." + vbCrLf + vbCrLf + "Warm regards and Happy xxxx from the xxxx Compliance Team"

                ' Open Word file object
                ' Open template Word file for the letter
                ' Add the text - ie for now add .Body to this file text in the right place
                ' Save this to a new file name like FilledLetter.docx
                ' Attach FilledLetter.docx
                ' display email still
                ' Debug.Print (".body is " + .body)
                
                Dim btext As String: btext = .body
                Dim endtext As String: endtext = "Eg can insert a message here for everybody in different formatting, like Happy Easter from xxxx! etc"
                
                ' Keep Excel file and template same folder, same level                
                Dim attached As Object: attached = wordLetter(ThisWorkbook.path & "\Template.docx", btext, endtext)
                .Attachments.Add ("C:\PATH\FinishedLetter.docx") '  (attached) doesnt work

                ' We can add files also like this
                '.Attachments.Add ("C:\DBS\JoeBloggs.doc") but will need to be sure files in right path before you do!
                ' Would want error handling anyway on the file path, with message added to error list
                .Display 'change to .Send will send the emails
            End With
            On Error GoTo 0
            Set OutMail = Nothing
            ' Trying to get rid of Locked for Editing message on Word file
            Set attached = Nothing

        Else
           ' MsgBox (cell.Value + " at cell " + CStr(cell.Address) + " is not a proper email address. Check it please!")
           ' TODO: fix error list print to Word file
           errorList = errorList + full_name + " at Row number " + CStr(RowNum) + " in the sheet had a problem with their email address or email doesn't exist, so no email was sent." + vbCrLf + vbCrLf
        End If
    Next cell
    
    End If
    
Debug.Print (errorList)

cleanup:
    Set OutApp = Nothing
    Application.ScreenUpdating = True
    
    
End Sub

' original code
' https://excel-macro.tutorialhorizon.com/vba-excel-edit-and-save-an-existing-word-document/

' Try with ... End With the object
' https://stackoverflow.com/questions/22569182/writing-formatting-word-document-using-excel-vba

' https://docs.microsoft.com/en-us/office/vba/api/word.document

Public Function wordLetter(templateFile As String, bodyText As String, endtext As String) As Object
' uncomment below if you want to debug file path later
' Debug.Print ("template file path coming in to wordletter function as: " + templateFile)

   Dim objWord

   Dim objDoc

   Dim objSelection

   Set objWord = CreateObject("Word.Application") ' NB this is NOT yet set to Nothing at the end, but still works...
   
   objWord.Application.DisplayAlerts = False
   
   objWord.Application.ScreenUpdating = False
   
   Set wordLetter = objWord.Documents.Open(templateFile, Visible = False, ReadOnlyRecommended = False)
   ' ReadonlyRecommended = False stops pop-up dialogue appearing every time this line executes! This is default property for Word docs maybe?
   
   ' Debug.Print ("wordLetter type is " + Str(VarType(wordLetter)))
   ' wordLetter returns as string type (8) here..
   
   ' TODO: Try setting objFont = objWord.Font as Selection may be not most stable according SO poster

   Set objSelection = objWord.Selection
   
   With objSelection.Font
       .Bold = False
       .colorIndex = wdBlack
       .Name = "Verdana"
       .Size = "11"
       
       objSelection.TypeText (bodyText)
       
       ' can change formatting for a uniform end text message
       .Bold = True
       .colorIndex = wdViolet ' not working apparently
       ' doesnt work; object / properties confused I think - .TextColor.ForeColor.RGB = RGB(0, 100, 100)
       .Size = "15"
       
       objSelection.TypeText (endtext)
       
   End With
   
   ' this is necessary to update Word file before saving
   ' Without it, file will be blank!
   objWord.Application.ScreenUpdating = True
   
   ' SaveAs2 needs the FileFormat specified here as well as filepath to save. MS Docs refer FileFormat as optional
   ' but I'm guessing may be because wordLetter is string type object here? Something I don't understand about this fully..
   With wordLetter
       .SaveAs2 Filename:="C:\Users\PATH\FinishedLetter.docx", FileFormat:=wdFormatDocumentDefault ' this is docx format
       ' doesn't work: .SaveAs2 Filename:="C:\Users\PATH\TestLetterSaving.docx"
   End With
   
   ' brings up locked for editing message unless you close it each time,
   ' because it's still open of course
   wordLetter.Close
   
   Application.DisplayAlerts = True

End Function

