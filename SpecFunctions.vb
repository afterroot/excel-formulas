' Functions By Sandip Vaghela
' Version 1.1
' Last Updated 09-04-21

' Do not Import this module as chars like '±' is lost while importing.

Public Const plusMinusSign = "±", rangeSign = "-", minKeyword = "min", maxKeyowrd = "max", charSpace = " "

'
' Supported types (Without Units)
' Ex. 36 ± 2, 34 - 38, Min 34
'
' Supported types (With Units)
' Ex. 36 ± 2 %, 34 - 38 %, Min 34 %
'
Function LCL(value, Optional default, Optional isPercent As Boolean, Optional roundTo As Integer) As Variant
    Dim myLeft, myRight, myLength, firstSpace
    posOfPlusMinus = InStr(value, plusMinusSign) ' Gets position of '±'
    posOfRange = InStr(value, rangeSign) ' Gets position of '-'
    posOfMin = InStr(LCase(value), minKeyword) ' Gets position of 'min'
    posOfMax = InStr(LCase(value), maxKeyowrd) ' Gets position of 'max'
    If IsMissing(default) Then
        default = 0
    End If
    If posOfPlusMinus > 0 Then
        myLeft = left(value, posOfPlusMinus - 1) ' Gets all characters to left of '±'
        myLength = InStr(posOfPlusMinus + 2, value, charSpace) - posOfPlusMinus ' Finds length after second charSpace to get tolerance. 0 if not available.
        If myLength < 0 Then
            tolerance = Mid(value, posOfPlusMinus + 1) ' Gets all characters to right of '±'
        Else
            tolerance = Mid(value, posOfPlusMinus + 1, myLength) ' Gets all characters to right of '±' but unit/other text eliminated
        End If
        LCL = CDbl(myLeft - tolerance) ' Minus for LCL
    ElseIf posOfRange > 0 Then
        myLeft = left(value, posOfRange - 1) ' Gets all characters to left of '-'
        LCL = CDbl(myLeft) ' Left of '-' is LCL
    ElseIf posOfMin > 0 Then
        firstSpace = InStr(posOfMin, value, charSpace) ' Finds 'space' after 'Min' keyword
        myLength = InStr(firstSpace + 1, value, charSpace) - firstSpace ' Finds length after second charSpace to get value. 0 if not available.
        If myLength < 0 Then
            myRight = Mid(value, firstSpace + 1) ' Gets all characters to right of 'Min'
        Else
            myRight = Mid(value, firstSpace + 1, myLength) ' Gets all characters to right of 'Min' but unit/other text eliminated
        End If
        LCL = CDbl(myRight) ' Right of 'Min' is LCL
    ElseIf posOfMax > 0 Then
        LCL = default
    ElseIf default > 0 Then
        LCL = default ' if parameter 'default' is supplied and parameter 'value' syntex not matched with '±' or '-' or 'Min'
    Else
        LCL = CVErr(xlErrNA) ' Value #N/A error if parameter 'default' is not supplied and parameter 'value' syntex not matched with '±' or '-' or 'Min'
    End If
    If isPercent Then
        LCL = LCL / 100 ' if parameter 'isPercent' set to TRUE then divide LCL by 100
    End If
    If roundTo > 0 Then
        LCL = Round(LCL, roundTo) ' if parameter 'roundTo' supplied then round LCL by provied number
    End If
End Function

'
' Supported types (Without Units)
' Ex. 36 ± 2, 34 - 38, Max 38
'
' Supported types (With Units)
' Ex. 36 ± 2 %, 34 - 38 %, Max 38 %
'
Function UCL(value, Optional default, Optional isPercent As Boolean, Optional roundTo As Integer) As Variant
    Dim myLeft, myRight, myLength, firstSpace
    posOfPlusMinus = InStr(value, plusMinusSign) ' Gets position of '±'
    posOfRange = InStr(value, rangeSign) ' Gets position of '-'
    posOfMin = InStr(LCase(value), minKeyword) ' Gets position of 'min'
    posOfMax = InStr(LCase(value), maxKeyowrd) ' Gets position of 'max'
    If IsMissing(default) Then
        default = 0
    End If
    If posOfPlusMinus > 0 Then
        myLeft = left(value, posOfPlusMinus - 1) ' Gets all characters to left of '±'
        myLength = InStr(posOfPlusMinus + 2, value, charSpace) - posOfPlusMinus ' Finds length after second charSpace to get tolerance. 0 if not available.
        If myLength < 0 Then
            tolerance = Mid(value, posOfPlusMinus + 1) ' Gets all characters to right of '±'
        Else
            tolerance = Mid(value, posOfPlusMinus + 1, myLength) ' Gets all characters to right of '±' but unit/other text eliminated
        End If
        UCL = CDbl(myLeft) + CDbl(tolerance) ' Plus for UCL
    ElseIf posOfRange > 0 Then
        myLength = InStr(posOfRange + 2, value, charSpace) - posOfRange ' Finds length after second charSpace to get UCL text. 0 if not available.
        If myLength < 0 Then
            myRight = Mid(value, posOfRange + 1) ' Gets all characters to right of '-'
        Else
            myRight = Mid(value, posOfRange + 1, myLength) ' Gets all characters to right of '-' but unit/other text eliminated
        End If
        UCL = CDbl(myRight) ' Right of '-' is UCL
    ElseIf posOfMax > 0 Then
        firstSpace = InStr(posOfMax, value, charSpace) ' Finds 'space' after 'max' keyword
        myLength = InStr(firstSpace + 1, value, charSpace) - firstSpace ' Finds length after second charSpace to get value. 0 if not available.
        If myLength < 0 Then
            myRight = Mid(value, firstSpace + 1) ' Gets all characters to right of 'Min'
        Else
            myRight = Mid(value, firstSpace + 1, myLength) ' Gets all characters to right of 'Min' but unit/other text eliminated
        End If
        UCL = CDbl(myRight) ' Right of 'max' is UCL
    ElseIf posOfMin > 0 Then
        UCL = default
    ElseIf default > 0 Then
        UCL = default ' if parameter 'default' is supplied and parameter 'value' syntex not matched with '±' or '-' or 'max'
    Else
        UCL = CVErr(xlErrNA) ' Value #N/A error if parameter 'default' is not supplied and parameter 'value' syntex not matched with '±' or '-' or 'max'
    End If
    If isPercent Then
        UCL = UCL / 100 ' if parameter 'isPercent' set to TRUE then divide UCL by 100
    End If
    If roundTo > 0 Then
        UCL = Round(UCL, roundTo) ' if parameter 'roundTo' supplied then round UCL by provied number
    End If
End Function

'
' Checks provided value is between spec.
'
Function ISBETWN(valueToCompare As Double, specText, Optional default, Optional isPercent As Boolean, Optional roundTo As Integer)
    myLcl = LCL(specText, default, isPercent, roundTo)
    myUcl = UCL(specText, default, isPercent, roundTo)
    If roundTo > 0 Then
        valueToCompare = Round(valueToCompare, roundTo) ' if parameter 'roundTo' supplied then round value by provied number
    End If
    If myLcl <= valueToCompare And valueToCompare <= myUcl Then
        ISBETWN = True
    Else
        ISBETWN = False
    End If
End Function

' Pending
Function LIMITTEXT(min As Double, maxOrTol As Double, limVariant)
    Select Case limVariant
        Case 0
            LIMITTEXT = Format((min + maxOrTol) / 2, "number") + Space(1) + plusMinusSign + Space(1) + Format(Abs((min - maxOrTol) / 2), "number")
            Case 1
            LIMITTEXT = Str(min) + Space(1) + plusMinusSign + Space(1) + Format(maxOrTol, "number")
        Case 2
            LIMITTEXT = Str(min) + Space(1) + rangeSign + Space(1) + Format(maxOrTol, "number")
    End Select
End Function


Sub test()
    ISBETWN 0.5, "Min 50 %", 100, True
End Sub
