Attribute VB_Name = "FuzzyAndMatchFunctions"
Option Explicit

Type matchData
    matched As Boolean
    fuzz1 As Double
    fuzz2 As Double
End Type

Function SortArrayAtoZ(myArray As Variant)

Dim i As Long
Dim j As Long
Dim Temp

'Sort the Array A-Z
' Will changed UCase indexes from (i) to (i)(0) and j to j 0
For i = LBound(myArray) To UBound(myArray) - 1
    For j = i + 1 To UBound(myArray)
        If UCase(myArray(i)(0)) > UCase(myArray(j)(0)) Then
            Temp = myArray(j)
            myArray(j) = myArray(i)
            myArray(i) = Temp
        End If
    Next j
Next i

SortArrayAtoZ = myArray

End Function

' Brute force function from Stack Overflow, to check if value is already in array or not
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function


Function FuzzyPercent(ByVal String1 As String, _
                      ByVal String2 As String, _
                      Optional Algorithm As Integer = 3, _
                      Optional Normalised As Boolean = False) As Single
'*************************************
'** Return a % match on two strings **
'*************************************
Dim intLen1 As Integer, intLen2 As Integer
Dim intCurLen As Integer
Dim intTo As Integer
Dim intPos As Integer
Dim intPtr As Integer
Dim intScore As Integer
Dim intTotScore As Integer
Dim intStartPos As Integer
Dim strWork As String

'-------------------------------------------------------
'-- If strings havent been normalised, normalise them --
'-------------------------------------------------------
If Normalised = False Then
    String1 = LCase$(Application.Trim(String1))
    String2 = LCase$(Application.Trim(String2))
End If

'----------------------------------------------
'-- Give 100% match if strings exactly equal --
'----------------------------------------------
If String1 = String2 Then
    FuzzyPercent = 1
    Exit Function
End If

intLen1 = Len(String1)
intLen2 = Len(String2)

'----------------------------------------
'-- Give 0% match if string length < 2 --
'----------------------------------------
If intLen1 < 2 Or intLen2 < 2 Then
    FuzzyPercent = 0
    Exit Function
End If

'----------------------------------------
'-- Will's Addition to function here --
' For each fuzzymatch part, ensure the LONGER name part is String2, because of the way fuzzymatch works..
' Put this IN FUZZYMATCH FUNCTION!
'----------------------------------------

'    Dim TempString As String
'    If intLen2 < intLen1 Then
'        TempString = String1
'        String1 = String2
'        String2 = TempString
'    End If


intTotScore = 0                   'initialise total possible score
intScore = 0                      'initialise current score

'--------------------------------------------------------
'-- If Algorithm = 1 or 3, Search for single characters --
'--------------------------------------------------------
'If (Algorithm And 1) <> 0 Then
'    FuzzyAlg1 String1, String2, intScore, intTotScore
'    If intLen1 < intLen2 Then FuzzyAlg1 String2, String1, intScore, intTotScore
'End If

' ** TESTING
If (Algorithm And 1) <> 0 Then
    If intLen1 < intLen2 Then
        FuzzyAlg1 String2, String1, intScore, intTotScore
        Else: FuzzyAlg1 String1, String2, intScore, intTotScore
    End If
End If


'-----------------------------------------------------------
'-- If Algorithm = 2 or 3, Search for pairs, triplets etc. --
'-----------------------------------------------------------
'If (Algorithm And 2) <> 0 Then
'    FuzzyAlg2 String1, String2, intScore, intTotScore
'    If intLen1 < intLen2 Then FuzzyAlg2 String2, String1, intScore, intTotScore
'End If


If (Algorithm And 2) <> 0 Then
    If intLen1 < intLen2 Then
        FuzzyAlg2 String2, String1, intScore, intTotScore
        Else: FuzzyAlg2 String1, String2, intScore, intTotScore
    End If
End If

FuzzyPercent = intScore / intTotScore

End Function
Private Sub FuzzyAlg1(ByVal String1 As String, _
                      ByVal String2 As String, _
                      ByRef Score As Integer, _
                      ByRef TotScore As Integer)
Dim intLen1 As Integer, intPos As Integer, intPtr As Integer, intStartPos As Integer

intLen1 = Len(String1)
TotScore = TotScore + intLen1              'update total possible score
intPos = 0
For intPtr = 1 To intLen1
    intStartPos = intPos + 1
    intPos = InStr(intStartPos, String2, Mid$(String1, intPtr, 1))
    If intPos > 0 Then
        If intPos > intStartPos + 3 Then     'No match if char is > 3 bytes away
            intPos = intStartPos
        Else
            Score = Score + 1          'Update current score
        End If
    Else
        intPos = intStartPos
    End If
Next intPtr
End Sub
Private Sub FuzzyAlg2(ByVal String1 As String, _
                        ByVal String2 As String, _
                        ByRef Score As Integer, _
                        ByRef TotScore As Integer)
Dim intCurLen As Integer, intLen1 As Integer, intTo As Integer, intPtr As Integer, intPos As Integer
Dim strWork As String

intLen1 = Len(String1)
For intCurLen = 2 To intLen1
    strWork = String2                          'Get a copy of String2
    intTo = intLen1 - intCurLen + 1
    TotScore = TotScore + Int(intLen1 / intCurLen)  'Update total possible score
    For intPtr = 1 To intTo Step intCurLen
        intPos = InStr(strWork, Mid$(String1, intPtr, intCurLen))
        If intPos > 0 Then
            Mid$(strWork, intPos, intCurLen) = String$(intCurLen, &H0) 'corrupt found string
            Score = Score + 1     'Update current score
        End If
    Next intPtr
Next intCurLen

End Sub



' Split bracket (RGN (RMN (HCA to right off name list,
' can't use general brackets as some names have eg bracket middle name!then trim trailing spaces
' Split strings into 2 parts - surname and forename(s)

' --> Do FuzzyMatch() >= 0.2 on first part of name, and 0.35 on surname
Function MatchNHSnames(ByVal String1 As String, ByVal String2 As String) As matchData
    
    Dim Words1() As String
    Dim Words2() As String
    Dim Name1(0 To 1) As String
    Dim Name2(0 To 1) As String
    
'    ********* splitting brackets in calling function now...
'    String1 = Split(String1, "(RGN")(0)
'    String1 = Split(String1, "(RMN")(0)
'    String1 = Split(String1, "(HCA")(0)
'    String2 = Split(String2, "(RGN")(0)
'    String2 = Split(String2, "(RMN")(0)
'    String2 = Split(String2, "(HCA")(0)
    ' Now whichever way round strings come in, split off the end where it says (RGN)_Spencer etc.
    String1 = Trim(String1)
    String2 = Trim(String2)
 
    'Split name into surname and forenames now
    Words1 = Split(String1)
    Name1(1) = Words1(UBound(Words1))
    ' split off rest of string as forename(s)
    Name1(0) = ""
    Dim a As Integer
    For a = LBound(Words1) To (UBound(Words1) - 1)
        Name1(0) = Name1(0) + Words1(a) + " "
    Next a
    Name1(0) = Trim(Name1(0))
    
    ' Repeat above for second name
    Words2 = Split(String2)
    Name2(1) = Words2(UBound(Words2))
    Name2(0) = ""
    Dim b As Integer
    For b = LBound(Words2) To (UBound(Words2) - 1)
        Name2(0) = Name2(0) + Words2(b) + " "
    Next b
    Name2(0) = Trim(Name2(0))
    
    ' *** OK splits 2 names correct now surname and forenames
    ' Debug.Print ("Name2(0) is " + Name2(0) + " and Name2(1) is " + Name2(1))
    
    ' Now we fuzzy compare Name1(0) with Name2(0), and fuzzy compare Name1(1) with Name2(1)
    ' Have now inserted string swap code into FuzzyMatch function itself
    ' so longer string is ensured as String2
    
    ' Do 0.2 or above on first part of name, and 0.35 on surname OR 0.5 on whole name
    
    ' ************ TESTING **************
'    Debug.Print ("Karen AP with longer first fuzzy % = " + CStr(FuzzyPercent("Karen Heather Allen-Powlett", "Karen Allen-Powlett")))
'    Debug.Print ("Karen AP with shorter first fuzzy % = " + CStr(FuzzyPercent("Karen Allen-Powlett", "Karen Heather Allen-Powlett")))

    
    ' with swapping names - now don't know which name is original column H name here
    ' BUT we only want the H row integer anyway which is still same in calling function though?
    
    MatchNHSnames.fuzz1 = FuzzyPercent(Name1(0), Name2(0))
    MatchNHSnames.fuzz2 = FuzzyPercent(Name1(1), Name2(1))
    'Debug.Print (CStr(MatchNHSnames.fuzz1))
    
    If MatchNHSnames.fuzz1 >= 0.2 And MatchNHSnames.fuzz2 >= 0.35 Then
        ' Debug.Print ("The name " + Name2(0) + " " + Name2(1) + " is coming as a match.")
        ' MatchNHSnames = True
        MatchNHSnames.matched = True
    Else
        ' MatchNHSnames = False
        MatchNHSnames.matched = False
    End If

    
    
    
End Function
