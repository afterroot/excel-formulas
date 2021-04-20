' Functions By Sandip Vaghela
' Version 1.1
' Last Updated 20-04-21

Attribute VB_Name = "utils"
Function WHENBLANK(value, showIfBlank)
    If Len(value) > 0 Then
        WHENBLANK = value
    Else
        WHENBLANK = showIfBlank
    End If
End Function

Function WHENBLANKORZERO(value, showIfBlank)
    If value = 0 Then
        WHENBLANKORZERO = showIfBlank
    ElseIf Len(value) > 0 Then
        WHENBLANKORZERO = value
    Else
        WHENBLANKORZERO = showIfBlank
    End If
End Function

Function CELLCOMMENT(cellRef As range)
    CELLCOMMENT = cellRef.comment.Text
End Function

Function ISFILEEXIST(path)
    If Dir(path) = "" Then
        ISFILEEXIST = False
    Else
        ISFILEEXIST = True
    End If
End Function