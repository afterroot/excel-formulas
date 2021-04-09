' Functions By Sandip Vaghela
' Version 1.0
' Last Updated 09-04-21

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