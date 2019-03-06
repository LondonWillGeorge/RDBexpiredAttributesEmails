Attribute VB_Name = "FilterEmailListCreateEmails"
' Option Explicit
' May need to End 1st word task manually in Task Manager upon running the application
Dim AttachCount As Integer

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

' Set attachment count, used to name attachment files differently
AttachCount = 0

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

' Formatting Word attachment with different formats in different parts of document
' https://stackoverflow.com/questions/26264715/vba-writing-to-word-changing-font-formatting
' Try convert HTML to plain text from here?
' https://stackoverflow.com/questions/5327512/convert-html-to-plain-text-in-vba

' https://stackoverflow.com/questions/27854534/how-do-i-insert-html-to-word-using-vba
' try getelementbyid or similar on the .htmlbody string
Sub CreateEmails_Click()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range
    
    'As Application declaring as application seems to generate obscure runtime errors?
    Dim objWord As Object ' passed to attachment create function, declaring as object now
    
    ' Set objWord = CreateObject("Word.Application") ' NB this is NOT yet set to Nothing at the end, but still works...
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    If Err.Number > 0 Then Set objWord = CreateObject("Word.Application")
    On Error GoTo 0
   ' Where is 0 here? Don't want to repeat html object creation?

    Application.ScreenUpdating = False
    Set OutApp = CreateObject("Outlook.Application")
    
    On Error GoTo cleanup

    Dim errorList As String: errorList = ""

    If Cells(2, 3) <> "" Then
    
    ' Open Word object and file ONCE for the Message Body, and save the text in several variables, only one format type in each text variable.
    ' Close it after variable text is saved. So we need to save the text variables BEFORE we start looping...
    Dim messagePath As String: messagePath = ThisWorkbook.path & "\MessageText.docx"
    
    Dim paragArray() As String
    ReDim Preserve paragArray(0 To 0)
    
    Dim objWordMessage As Object: Set objWordMessage = CreateObject("Word.Application")
    ' objWordMessage.Visible = False
    objWordMessage.Application.DisplayAlerts = False
    objWordMessage.Application.ScreenUpdating = False
    
'    Dim msgDoc As Object: Set msgDoc = objWordMessage.Documents.Open(messagePath, Visible = False, ReadOnlyRecommended = False)
'    ' check path and show mgbox if path has no file
'
'    Set colParagraphs = msgDoc.Paragraphs
'    paraCount = 0
'    For Each objParagraph In colParagraphs
'
'        ' Debug.Print ("parag text is: " + objParagraph.Range.text)
'        lineText = Trim(objParagraph.Range.text)
'        If Trim(lineText) <> "" Then ' Trimming again here still gives blank items between!
'           ' add to the text string array here
'           paragArray(paraCount) = lineText
'           paraCount = paraCount + 1
'           ReDim Preserve paragArray(0 To paraCount)
'        End If
'    Next
'
'    msgDoc.Close
    
    ' TODO: Find way to put ALL of text in Word files and load in without mangling the HTML, or even load directly from Word and convert into HTML
    ' maybe someone wrote VBA module to do this already?
    Dim introPath As String: introPath = ThisWorkbook.path & "\MessageIntroHTML.docx"
    
    Dim introDoc As Object: Set introDoc = objWordMessage.Documents.Open(introPath, Visible = False, ReadOnlyRecommended = False)
    
    Dim introHTML As String: introHTML = introDoc.Range.text
    
    introDoc.Close
    
    Set introDoc = Nothing
    
'  Footer is so short, leaving it in hard code for now
'    ' Same for the footer for now, TODO: shorten to one set commands, and ideally one document
'    Dim footerPath As String: footerPath = ThisWorkbook.path & "\MessageFooterHTML.docx"
'
'    Dim footerDoc As Object: Set footerDoc = objWordMessage.Documents.Open(footerPath, Visible = False, ReadOnlyRecommended = False)
'
'    Dim footerHTML As String: footerHTML = footerDoc.Range.text
'
'    footerDoc.Close
'
'    Set footerDoc = Nothing
    
    ' The .Quit line closes the Word process in Windows Task Manager - crucial!
    ' Setting object to Nothing without this will still NOT end the Word Task
    objWordMessage.Quit
    
    Set objWordMessage = Nothing
    
    ' index 0 in paragraph array should now be subject line,
    ' each underlined title or paragraph is separate item in array
    
'    For a = 0 To UBound(paragArray)
'        Debug.Print ("no " + Str(a) + " is: " + paragArray(a))
'    Next a
    
    ' on debugging, Word Quit execution hangs before break point execution,
    ' have to manually end task in Task Manager for some reason to get to the break point?
    
    'Start processing the emails here:
    For Each cell In ActiveSheet.Columns("C").Cells.SpecialCells(xlCellTypeConstants)
        
        ' *** FOR TESTING, STOP AT ROW 4, COMMENT OUT LATER!!!! 4 March 10:30 ran with 20 OK ****
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
                .Importance = 2
                .To = cell.Value
                .Subject = "Urgent: Your xxxx Company Compliance Documents are Expiring!" ' paragArray(0)
                ' .body = emailMainText1(firstName, paragArray()) ' returns string with main text before variable text
                
                .BodyFormat = olFormatHTML
                
                ' Ron De Bruin's website again: remember need "" double quotes inside the HTML style tags for them to be read,
                ' somewhat nastily it will not throw error with single " just runs by ignoring it
                
                ' was 14px, is it picking up font size? it does pick up color attribute
                .htmlbody = "<HTML><body style=""font-family: Calibri; font-size: 16px; color: #000; line-height: 1; font-weight: bold;"">" + "<div><p>Dear " + firstName + ",</p></div>"
 
                .htmlbody = .htmlbody + introHTML
                
                proofs = 0
                refs = 0
                training = 0
                dvla = False

                For colNum = 5 To 21
                ' iterate to right until empty, add the text to .htmlbody
'                    If Trim(Cells(RowNum, colNum)) = "" Then
'                        Exit For
'                    ElseIf colNum = 5 Then
'                        .Subject = .Subject + "s"
'                    End If

                    ' add text
                    cellValue = Cells(RowNum, colNum)
                    ' docName = ""

                    ' Select case here would be best I think, set standard text string variable,
                    ' and add this in appropriate place in select case text for each attribute case.

                    ' Being naughty for now, and adding strings as hard code here, TODO: save in JSON file or Word file
                    ' changing <br> tags to <p> now as got single line spacing working in body tag
                    
                    ' 5/3/19 NB The variable paragraphs are not picked up by HTML parser with <P><h3>...</h3>...</P> format despite Outlook not having problems with this format
                    ' So change to <h3>..</h3><p>...</p> seems to be picked up properly then
                    Select Case cellValue
                        Case "DBS"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>DBS</u></h3><p>Your DBS needs updating. If you have moved address within the last 12 months, please may you " + _
                            "provide us with your new full address and the date you moved into this address. If you registered with " + _
                            "the DBS Update Service, please may you provide us with the hard copy of your DBS and the 16 digit disclosure " + _
                            "number, so we can make relevant checks on-line. Please may you email this to xxxx" + "</p></div>"
                            
                        Case "FTW"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>(FTW) - Fitness to Work Certificate</u></h3><p>Every year we need to obtain a new FTW certificate for you. If your " + _
                            "health has changed, please may you inform us at xxxx. If your circumstances haven't changed " + _
                            "in the last 12 months, please do inform us, then we can apply for a new FTW certificate for you. If you work in an xxxx area, " + _
                            "do inform us (yyyy)." + "</p></div>"

                        Case "Appraisal"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>Appraisal</u></h3><p>When you join xxxx, you will be required to have an appraisal within the first six " + _
                            "months of joining us. Thereafter you will be required to have an appraisal annually. You are due an appraisal, so please may " + _
                            "you call the Compliance Team on <span  style=""color: #800080;"">0800 1234 5678</span> and one of our xxxx will conduct an appraisal with you." + "</p></div>"

                        Case "BLS"
                            training = training + 1
                        
                        Case "NMC"
                            ' passing at the moment
                        
                        Case "Manual Handling"
                            training = training + 1

                        Case "Proof Address1"
                            proofs = proofs + 1
                            
                        Case "Proof Address2"
                            proofs = proofs + 1
                            
                        Case "Ref1"
                            refs = refs + 1
                            
                        Case "Ref2"
                            refs = refs + 1
                            
                        Case "EU Passport"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>EU Passport & Right To Work in the UK (Brexit)</u></h3><p>Your Passport is about to expire. It is a " + _
                            "legal requirement that you update this and send us a clear copy of your renewed Passport. Please may you send " + _
                            "this to xxxx. If you're from the EU, from the 30th March 2019, we require a copy " + _
                            "of your Pre-Settled or Settled status; without this, you will not be able to work in the UK. For more information " + _
                            "about how you can obtain your status, please call the Compliance Team on <span  style=""color: #800080;"">0800 1234 5678</span> or email us at " + _
                            "xxxx" + "</p></div>"
                        
                        Case "ROW Passport"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>Non-EU Passport</u></h3><p>Your Passport is about to expire. It is a legal requirement that you update this and " + _
                            "send us a clear copy of your renewed Passport. Please may you send this to xxxx" + "</p></div>"
                        
                        Case "UK Passport"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>UK Passport</u></h3><p>Your Passport is about to expire. It is a legal requirement that you update this and " + _
                            "send us a clear copy of your renewed Passport. Please may you send this to xxxx" + "</p></div>"
                        
                        Case "DVLA"
                            dvla = True
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>DVLA</u></h3><p>We require a copy of your driving licence, please may you email this to us - " + _
                            "xxxx." + "</p></div>"
                        
                        Case "Visa"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>Visa</u></h3><p>Your Visa is about to expire. It is a legal requirement that you update this " + _
                            "and send us a clear copy of your renewed Visa. Please may you send this to xxxx." + "</p></div>"

                        Case "ID Badge"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>ID Badge</u></h3><p>Your ID badge is only valid for one year and your current one is expiring. " + _
                            "You will shortly receive a new valid ID badge in the post (if your address has changed within the " + _
                            "last 12 months, please inform the compliance team asap on <span  style=""color: #800080;"">0800 1234 5678</span> or email them at " + _
                            "xxxx)." + "</p></div>"
                            
                            errorList = errorList + full_name + " has an out of date ID Badge on RDB." + vbCrLf
                            
                        Case "YMCA"
                            .htmlbody = .htmlbody + _
                            "<div><h3><u>YMCA - Prevention & Management of Bad Dancing Certificate</u></h3><p>We " + _
                            "require an up to date YMCA training certificate from you. If you have completed a course elsewhere, " + _
                            "may you email this to xxxx. " + _
                            "If you haven't completed this course within the last year, we will happily book you into a course. Please " + _
                            "may you call our Compliance Team on 0800 1234 5678 so that we can arrange this." + "</p></div>"

                        Case Else
                               
                               errorList = errorList + full_name + " has an error with a document listed as " + cellValue + vbCrLf
                    End Select

                    ' check for permission to work combination invalidities, and generate corresponding messages.

                Next colNum
                
                If training > 0 Then
                    .htmlbody = .htmlbody + _
                    "<div><h3><u>Mandatory Training</u></h3><p>Your mandatory training is about to expire. If you have completed your " + _
                    "training for Moving & Handling, Basic Life Support or any other training elsewhere, please may you " + _
                    "forward these to us at xxxx. Alternatively we will pay and book you into " + _
                    "various on-line courses or practical training courses close to your house, so please get in contact with us " + _
                    "ASAP on 0800 1234 5678 or email us at xxxx" + "</p></div>"
                    
                End If

                If proofs > 0 Then
                    .htmlbody = .htmlbody + _
                    "<div><h3><u>Proof of Address</u></h3><p>We require two proofs of your address, this can be Utility Bills, Bank Statements, Council Tax " + _
                    "Bill, Letter from HMRC / Job Centre or your Driving Licence. Please may you email us a clear copy of two Proofs of Addresses to " + _
                    "xxxx." + "</p></div>"

                End If

                If refs > 0 Then
                    .htmlbody = .htmlbody + _
                    "<div><h3><u>Professional References</u></h3><p>Annually we have to renew your references, therefore we require two professional references " + _
                    "for you. Please may you provide us with the full name of the referee, their position, their place of work, their email " + _
                    "address and their contact telephone number. Please may you email this information to xxxx" + "</p></div>"

                End If
                
                ' Add footer signiature to the message
                .htmlbody = .htmlbody + "<div><p><br><br>Kind regards,<br><br><br>" + _
                "<b><span  style=""color: #672983;"">The Compliance Team</span><br><span  style=""color: #0399A3;"">xxxx Company</span><br><br><br>Tel <span  style=""color: #672983;"">0800 1234 5678</span><br>Email compliance@xxxxcompany.co.uk<br>Web <span  style=""color: #672983;"">www.xxxxcompany.co.uk</span></b><br><br></p></div>"
                
                .htmlbody = .htmlbody + "</HTML></BODY>"
                
                Dim btext As String: btext = .htmlbody
                
                ' make email address xxxx brand green #0399A3 in htmlbody string everywhere here, NB in Intro string from Word file, Excel/Word mangles colour tags and therefore disables them somehow
                ' Doing it here avoids btext passing them into the attachment function, also it's neater code
                .htmlbody = Replace(.htmlbody, "compliance@xxxxcompany.co.uk", "<span  style=""color: #0399A3;"">compliance@xxxxcompany.co.uk</span>")

                ' "Template.docx" keep Excel file and template same folder, same level

                Dim attached As Object: attached = wordLetter(ThisWorkbook.path & "\Template.docx", btext, objWord)
                
                ' Print doesn't work because attached is Nothing at this point Locals window shows, I don't know why yet trying to return it as the wordLetter Word object
                ' Debug.Print ("attached doc path is: " + attached.document.path)
                
                savePath2 = ThisWorkbook.path & "\Attachments\FinishedLetter" & Str(AttachCount) & ".docx"
                
                .Attachments.Add (savePath2)  '  (attached) doesnt work because it's Nothing
                ' & "\Attachments\" & full_name & "-NextStepExpiringComplianceDocs.docx"
                
                ' Trying to get rid of Locked for Editing message on Word file
                Set attached = Nothing
                attached.Close
                
                ' Would want error handling anyway on the file path, with message added to error list
                .Display 'change to .Send will send the emails
            End With
            On Error GoTo 0
            Set OutMail = Nothing

        Else
           ' MsgBox (cell.Value + " at cell " + CStr(cell.Address) + " is not a proper email address. Check it please!")
           ' TODO: fix error list print to Word file
           errorList = errorList + full_name + " at Row number " + CStr(RowNum) + " in the sheet had a problem with their email address or email doesn't exist, so no email was sent." + vbCrLf + vbCrLf
        End If
    Next cell
    
    End If
    
' Debug.Print (errorList)

cleanup:
    Set OutApp = Nothing
    Application.ScreenUpdating = True
    
    
End Sub

' original code
' https://excel-macro.tutorialhorizon.com/vba-excel-edit-and-save-an-existing-word-document/

' Try with ... End With the object
' https://stackoverflow.com/questions/22569182/writing-formatting-word-document-using-excel-vba

' https://docs.microsoft.com/en-us/office/vba/api/word.document

Public Function wordLetter(templateFile As String, bodyText As String, objWord As Object) As Object

    ' For the attachment, try to parse the whole htmlbody string as HTML..
    Dim html As HTMLDocument ' as new maybe not the best as late binding better?
    ' html = create...
    Set html = CreateObject("htmlfile")
    html.body.innerHTML = bodyText
    ' Debug.Print (html.body.innerHTML)

                
'                For Each y In tagas
'                    Debug.Print ("A tag is: " + y.innerText + y.href)
'                Next y

   Dim objDoc

   Dim objSelection
   
'   ' Set objWord = CreateObject("Word.Application") ' NB this is NOT yet set to Nothing at the end, but still works...
'   On Error Resume Next
'   Set objWord = GetObject(, "Word.Application")
'   If Err.Number > 0 Then Set objWord = CreateObject("Word.Application")
'   On Error GoTo 0
'   ' Where is 0 here? Don't want to repeat html object creation?
   
   objWord.Application.DisplayAlerts = False
   
   objWord.Application.ScreenUpdating = False
   
   Set wordLetter = objWord.Documents.Open(templateFile, Visible = False, ReadOnlyRecommended = False)
   ' ReadonlyRecommended = False stops pop-up dialogue appearing every time this line executes! This is default property for Word docs maybe?
   
   ' Debug.Print ("wordLetter type is " + Str(VarType(wordLetter)))
   ' wordLetter returns as string type (8) here..
   
   ' Set whole doc colour to black to be safe... worried without this, local colour set
   ' may affect other parts of document
   wordLetter.Range.Font.textColor.RGB = RGB(0, 0, 0)

    ' Must SET an object, can't just use = !
    ' Need put divs around each paragraph
    Dim tagdivs As Object: Set tagdivs = html.getElementsByTagName("div")
    Dim tagH3s As Object
    Dim tagps As Object
    
    ' Need all paragraphs in Word loaded file and in hard code to have div tags enclosing, then they're picked up properly here
    For Each div In tagdivs
        'reject the footer text by checking if "kind regards" is in it! bit hacky!
        If InStr(div.innerHTML, "Kind regards") = 0 Then
            ' check if h3 is in div or just p, if just p it is in first main paragraph section.
            Set tagH3s = div.getElementsByTagName("h3")
            If tagH3s.Length > 0 Then
                For Each h3 In tagH3s
                    wordLetter.Paragraphs.Add
                    pct = wordLetter.Paragraphs.Count
                    With wordLetter.Paragraphs(pct).Range
                        .text = h3.innerText
                        '.Font.textColor.RGB = RGB(143, 8, 201) ' Purple
                        'brighten text a bit too
                        '.Font.textColor.Brightness = 0.4
                        .Font.Underline = True
                        .Font.Bold = True
                    End With
                    Debug.Print ("got at least one h3 tag which is: " + h3.innerHTML)
                Next h3
            End If
            
            Set tagps = div.getElementsByTagName("p")
            If tagps.Length > 0 Then
                For Each p In tagps
                    wordLetter.Paragraphs.Add
                    pct = wordLetter.Paragraphs.Count
                    With wordLetter.Paragraphs(pct).Range
                        ' check inner text for footer unique substring, if yes call footer process function
                        Debug.Print ("tagp HTML: " + p.innerHTML)
                        .text = p.innerText
                        .Font.Underline = False
                        .Font.Bold = False
                    End With
                    Debug.Print ("got p inner text is: " + p.innerText)
                Next p
            End If
            
            ' Keep so when testing, can see immediate window this is processing to here:
            Debug.Print ("div content is: " + div.innerHTML + vbCrLf)
        End If
    Next div
    
    ' print word footer here individually, otherwise too much of a brain fry with in-paragraph formatting
    ' get the character number of a string with Instr(string, substring) returns 0 if not, but if there,
    ' gives the character number first character is 1, then use this in Word range to format
    ' ActiveDocument.Range(ActiveDocument.Paragraphs(1).Range.Characters(5).Start, _
    ActiveDocument.Paragraphs(1).Range.Characters(10).End).Font.Bold = True
    wordLetter.Paragraphs.Add
    pct = wordLetter.Paragraphs.Count
    With wordLetter.Paragraphs(pct).Range
        .text = vbCrLf & "Kind regards," & vbCrLf & vbCrLf & vbCrLf & "The Compliance Team" & vbCrLf & _
        "xxxx Company" & vbCrLf & "Tel 0800 1234 5678" & vbCrLf & "Email compliance@xxxxcompany.co.uk" _
        & vbCrLf & "Web www.xxxxcompany.co.uk"
        With .Font
            .Bold = True
            'testing
            .textColor.RGB = RGB(143, 8, 201)
        End With
        
        'NB Characters collection doesnt work this way: With .Characters(1, 14)
        ' Apply non-bold and black colour to "Kind regards," only
        For a = 1 To 14
            With .Characters(a)
                .Font.Bold = False
                .Font.textColor.RGB = RGB(0, 0, 0)
            End With
        Next a
        
        Debug.Print ("instring comp team: " + Str(InStr(.text, "The Compliance Team"))) ' gives 18, so starts at 1, same as characters collection
        
'        For b = chStart To chEnd
'            With .Characters(b)
'                .Font.textColor.RGB = RGB(143, 8, 201) ' Purple
'                'brighten text a bit too
'                .Font.textColor.Brightness = 0.4
'            End With
'        Next b
        


    End With
    
    
'    For Each tagp In tagps
'        ' Will be para on own, or have h3 heading tag inside at top
'        ' some have <a> tags around the web address, remove tags and replace with hyperlink in Word doc
'        ' With ActiveDocument.Paragraphs(1).Range end with
'        ' ActiveDocument.Paragraphs.Add - This example adds a new paragraph mark at the end of the active document.
'        ' tagps.get... should also work
'        ' Paragraphs collection object starts at index 1 apparently? https://docs.microsoft.com/en-us/office/vba/api/word.paragraphs
'        wordLetter.Paragraphs.Add ' ActiveDocument.Paragraphs.Add exits here because of error, perhaps there is no active document?
'        pct = wordLetter.Paragraphs.Count
'        With wordLetter.Paragraphs(pct).Range
'            .text = tagp.innerText ' .innerHTML error with .typetext doesn't accept, and .innerText returns "" for <p> tags with <h3> inside them it seems...
'            .Underline = True
'            .Bold = True
'        End With
'
'    Next tagp
    
    ' test print the whole HTML string in the Attachment to plan parsing route
'    wordLetter.Paragraphs.Add
'    last = wordLetter.Paragraphs.Count
'    With wordLetter.Paragraphs(last).Range
'        .text = vbCrLf + vbCrLf + html.body.innerHTML ' error with .typetext doesn't accept, and .innerText returns "" for <p> tags with <h3> inside them it seems...
'    End With

'   Set objSelection = objWord.Selection
'
'   With objSelection.Font
'       .Bold = False
'       .colorIndex = wdBlack
'       .Name = "Verdana"
'       .Size = "11"
'
'       objSelection.typetext (bodyText)
'
'       ' can change formatting for a uniform end text message
'       .Bold = True
'       .colorIndex = wdViolet ' not working apparently
'       ' doesnt work; object / properties confused I think - .TextColor.ForeColor.RGB = RGB(0, 100, 100)
'       .Size = "15"
'
'       objSelection.typetext (endtext)
'
'   End With
   
   ' this is necessary to update Word file before saving
   ' Without it, file will be blank!
   objWord.Application.ScreenUpdating = True
   
   ' DELETE the old file before resaving with SAME NAME
   ' But this still doesn't work it seems remembering and numbering consecutive file names
   ' Error if try save same name 2nd time around, cant find how to overwrite yet
   
   AttachCount = AttachCount + 1
   savePath = ThisWorkbook.path & "\Attachments\Letter" & Str(AttachCount) & ".docx"
   
'   If Dir(savePath) <> "" Then
'      ' First remove readonly attribute, if set
'      SetAttr savePath, vbNormal
'      ' Then delete the file
'      Kill savePath ' this throws error on 2nd time this function called
'   End If
   
   ' SaveAs2 needs the FileFormat specified here as well as filepath to save. MS Docs refer FileFormat as optional
   ' but I'm guessing may be because wordLetter is string type object here? Something I don't understand about this fully..
   With wordLetter
       ' .SaveAs2 Filename:=ThisWorkbook.path & "\Attachments\" & full_name & "-FinishedLetter.docx", FileFormat:=wdFormatDocumentDefault
       .SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault    ' this is docx format
       ' doesn't work: .SaveAs2 Filename:="C:\Users\PATH\TestLetterSaving.docx"
   End With
   
   Debug.Print ("wordLetter just saved to: " + wordLetter.path + "\" + wordLetter.Name)
   
   Set html = Nothing
   
   ' Set nothing in calling function..
   ' Set wordLetter = Nothing
   ' brings up locked for editing message unless you close it each time,
   ' because it's still open of course
   ' wordLetter.Close
   
   Application.DisplayAlerts = True

End Function
