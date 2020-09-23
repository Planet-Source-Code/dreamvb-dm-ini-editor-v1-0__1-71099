Attribute VB_Name = "modFunc"
Option Explicit

Public LastDirLoc As String
Public ButtonPress As Integer
Public mValName As String
Public mValData As String
Public mCurSelection As String

Public Enum TIniEditMode
    INI_EDIT_VALUE = 1
    INI_NEW_VALUE = 2
    INI_NEW_SELECTION = 3
    INI_RENAME_SELECTION = 4
End Enum

Public mEditMode As TIniEditMode

Public Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function GetFilename(lFile As String) As String
Dim spos As Integer
    spos = InStrRev(lFile, "\", Len(lFile), vbBinaryCompare)
    
    If (spos > 0) Then
        GetFilename = Mid$(lFile, spos + 1)
    End If
    
End Function

Public Function GetAbsPath(lPath As String) As String
Dim spos As Integer
    spos = InStrRev(lPath, "\", Len(lPath), vbBinaryCompare)
    
    If (spos > 0) Then
        GetAbsPath = Left$(lPath, spos)
    End If
End Function
