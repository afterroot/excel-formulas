' Functions By Sandip Vaghela
' Version 1.2
' Last Updated 17-05-21

' Do not Import this module as chars like '±' is lost while importing.

Public Const plusMinusSign = "±", rangeSign = "-", minKeyword = "min", maxKeyowrd = "max", charSpace = " "

'
' Supported types (Without Units)
' Ex. 36 ± 2, 34 - 38, Min 34, -36 ± 2, -38 - -34, Min -38
'
' Supported types (With Units)
' Ex. 36 ± 2 %, 34 - 38 %, Min 34 %, -36 ± 2 Unit, -38 - -34 Unit, Min -38 Unit
'
Function LCL(value, Optional default, Optional isPercent As Boolean, Optional roundTo As Integer) As Variant
    Dim myLeft, myRight, myLength, firstSpace
    posOfPlusMinus = InStr(value, plusMinusSign) ' Gets position of '±'
    posOfRange = InStr(2, value, rangeSign) ' Gets position of '-' (Support for negative values, InStr will find '-' from 2nd char)
    posOfMin = InStr(LCase(value), minKeyword) ' Gets position of 'min'
    posOfMax = InStr(LCase(value), maxKeyowrd) ' Gets position of 'max'
    If posOfMin > 0 Or posOfMax > 0 Then
        posOfPlusMinus = 0
        posOfRange = 0
    End If
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
' Ex. 36 ± 2, 34 - 38, Max 38, -36 ± 2, -38 - -34, Max -34
'
' Supported types (With Units)
' Ex. 36 ± 2 %, 34 - 38 %, Max 38 %, -36 ± 2 Unit, -38 - -34 Unit, Max -34 Unit
'
Function UCL(value, Optional default, Optional isPercent As Boolean, Optional roundTo As Integer) As Variant
    Dim myLeft, myRight, myLength, firstSpace
    posOfPlusMinus = InStr(value, plusMinusSign) ' Gets position of '±'
    posOfRange = InStr(2, value, rangeSign) ' Gets position of '-' (Support for negative values, InStr will find '-' from 2nd char)
    posOfMin = InStr(LCase(value), minKeyword) ' Gets position of 'min'
    posOfMax = InStr(LCase(value), maxKeyowrd) ' Gets position of 'max'
    If posOfMin > 0 Or posOfMax > 0 Then
        posOfPlusMinus = 0
        posOfRange = 0
    End If
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

'
' Returns SpecText from min and max values
'
Function LIMITTEXT(min As Double, maxOrTol As Double, limVariant, Optional unit)
    Select Case limVariant
        Case 0 'Variant: PlusMinus
            LIMITTEXT = Trim(((min + maxOrTol) / 2)) + Space(1) + plusMinusSign + Space(1) + Trim(Str(Abs((min - maxOrTol) / 2)))
        Case 1 'Variant: Min
            LIMITTEXT = "Min" + Space(1) + Trim(Str(min))
        Case 2 'Variant: Max
            LIMITTEXT = "Max" + Space(1) + Trim(Str(maxOrTol))
        Case 3 'Variant: Range
            LIMITTEXT = Trim(Str(min)) + Space(1) + rangeSign + Space(1) + Trim(Str(maxOrTol))
    End Select
    If Not IsMissing(unit) Then
        LIMITTEXT = LIMITTEXT + Space(1) + unit 'Append Unit at last if provided
    End If
End Function


' Run tests to verify
Sub test()
    Dim testMin As Double, testMax As Double, plusMinusVariant As String, minVariant As String, _
    maxVariant As String, rangeVariant As String, myResult As String
    
    'Positive Value Test
    testMin = CDbl(2)
    testMax = CDbl(4)
    plusMinusVariant = LIMITTEXT(testMin, testMax, 0) 'PlusMinus Variant
    minVariant = LIMITTEXT(testMin, testMax, 1) 'Min Variant
    maxVariant = LIMITTEXT(testMin, testMax, 2) 'Max Variant
    rangeVariant = LIMITTEXT(testMin, testMax, 3) 'Range Variant
    
    myResult = "Test Result:" + vbNewLine
    myResult = myResult + vbNewLine + "Positive Values: "
    
    'PlusMinus Variant
    myResult = myResult + vbNewLine + "plusMinusVariant: "
    If plusMinusVariant = "3 ± 1" Then
        myResult = myResult + "Pass"
    ElseIf plusMinusVariant <> "3 ± 1" Then
        myResult = myResult + "Fail"
    End If
    
    'Min Variant
    myResult = myResult + vbNewLine + "minVariant: "
    If minVariant = "Min 2" Then
        myResult = myResult + "Pass"
    ElseIf minVariant <> "Min 2" Then
        myResult = myResult + "Fail"
    End If
    
    'Max Variant
    myResult = myResult + vbNewLine + "maxVariant: "
    If maxVariant = "Max 4" Then
        myResult = myResult + "Pass"
    ElseIf maxVariant <> "Max 4" Then
        myResult = myResult + "Fail"
    End If
    
    'Range Variant
    myResult = myResult + vbNewLine + "rangeVariant: "
    If rangeVariant = "2 - 4" Then
        myResult = myResult + "Pass"
    ElseIf rangeVariant <> "2 - 4" Then
        myResult = myResult + "Fail"
    End If
    
    'Negative Value Test
    testMin = CDbl(-4)
    testMax = CDbl(-2)
    plusMinusVariant = LIMITTEXT(testMin, testMax, 0) 'PlusMinus Variant
    minVariant = LIMITTEXT(testMin, testMax, 1) 'Min Variant
    maxVariant = LIMITTEXT(testMin, testMax, 2) 'Max Variant
    rangeVariant = LIMITTEXT(testMin, testMax, 3) 'Range Variant
    myResult = myResult + vbNewLine + vbNewLine + "Negative Values: "
    
    'PlusMinus Variant
    myResult = myResult + vbNewLine + "plusMinusVariant: "
    If plusMinusVariant = "-3 ± 1" Then
        myResult = myResult + "Pass"
    ElseIf plusMinusVariant <> "-3 ± 1" Then
        myResult = myResult + "Fail"
    End If
    
    'Min Variant
    myResult = myResult + vbNewLine + "minVariant: "
    If minVariant = "Min -4" Then
        myResult = myResult + "Pass"
    ElseIf minVariant <> "Min -4" Then
        myResult = myResult + "Fail"
    End If
    
    'Max Variant
    myResult = myResult + vbNewLine + "maxVariant: "
    If maxVariant = "Max -2" Then
        myResult = myResult + "Pass"
    ElseIf maxVariant <> "Max -2" Then
        myResult = myResult + "Fail"
    End If
    
    'Range Variant
    myResult = myResult + vbNewLine + "rangeVariant: "
    If rangeVariant = "-4 - -2" Then
        myResult = myResult + "Pass"
    ElseIf rangeVariant <> "-4 - -2" Then
        myResult = myResult + "Fail"
    End If
    
    'Show Dialog at last
    MsgBox (myResult)
End Sub
