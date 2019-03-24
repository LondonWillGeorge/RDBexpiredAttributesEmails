
' NB This module requires the Microsoft Scripting Runtime reference to have been added to this project.
' It should have been added already! But in case problem, add again from VBA Editor by clicking Tools menu
' Then References, then choosing Microsoft Scripting Runtime from the list.

' NB Alternative to File System Object is using Dir built-in VBA function doesn't seem work very easily,
' plus someone on SO says it's silly anyway..
Sub PasteColumnG_ListfromFolderNames()

    Dim fso As New FileSystemObject
    ' Get the file path from whatever user typed in Cell E26, trim any spaces front and end of this
    Dim path As String: path = Trim(Sheets("START").Cells(26, 5).Value)
    Dim fldParent As Folder
    Dim fldChild As Folder
    Dim folderName As String
    
    Dim gRow As Integer: gRow = 2
    
    ' Don't need an array here yet: Dim fldArray() As String
    
    ' Clear previous results
    Sheets("START").Range("G:G").Clear
    
    ' If user pastes folder path without the \ at end, then add this on, and also show in the typing box
    If Right(path, 1) <> "\" Then
        path = path & "\"
        Sheets("START").Cells(26, 5).Value = path
    End If
    
    If fso.FolderExists(path) Then
        Set fldParent = fso.GetFolder(path)
        If fldParent.SubFolders.Count > 0 Then
            For Each fldChild In fldParent.SubFolders
            ' Debug.Print fldChild.Name
            Sheets("START").Cells(gRow, 7) = fldChild.Name ' testing with Column H! which is 8
            gRow = gRow + 1
            Next
        Else
            MsgBox ("Sorry, but I can't find any sub-folders in the folder:" + vbCrLf + vbCrLf + path + vbCrLf + vbCrLf + "Try checking this is the right folder in Windows Explorer?")
        End If
        
    Else
        MsgBox ("Sorry, but I can't find any folder at the file path which you typed:" + vbCrLf + vbCrLf + path + vbCrLf + vbCrLf + "Try double-checking, copying and pasting again from Windows Explorer?")
    
    End If

End Sub
