Attribute VB_Name = "Module1"
Public Function numWor(a As String) As String
If IsNumeric(Val(numWor)) = False Then
 numWor = ""
Exit Function
End If
Dim conNum(6)
conNum(1) = "crore "
conNum(2) = "lakh "
conNum(3) = "thousand "
conNum(4) = "hundred "
conNum(5) = ""
Dim f As String
Dim s As String
a = Format(a, "000000000.00")
d = 1
For i = 1 To 5
If i <> 4 Then
    If Val(Mid(a, d, 2)) >= 20 And Val(Mid(a, d, 2)) <= 99 Then
    f = Mid(a, d, 1) & "0"
    s = Mid(a, d + 1, 1)
    numWor = numWor & convert(f) & convert(s) & conNum(i)
    End If
    If Val(Mid(a, d, 2)) >= 1 And Val(Mid(a, d, 2)) < 20 Then
        numWor = numWor & convert(Val(Mid(a, d, 2))) & conNum(i)
    End If
    d = d + 2
Else
    If Val(Mid(a, d, 1)) >= 1 And Val(Mid(a, d, 1)) < 10 Then
        numWor = numWor & convert(Val(Mid(a, d, 1))) & conNum(i)
        If Val(Mid(a, d + 1, 2)) > 0 Then numWor = numWor & "and "
    End If
    d = d + 1
End If
Next

If Val(Mid(a, 1, 9)) > 0 Then numWor = numWor & "Rupees "
If Val(Right(a, 2)) > 0 Then
            If Val(Mid(a, 1, 9)) > 0 Then numWor = numWor & "And "
            If Val(Right(a, 2)) >= 20 And Val(Right(a, 2)) <= 99 Then
                f = Mid(a, 11, 1) & "0"
                s = Mid(a, 12, 1)
                numWor = numWor & convert(f) & convert(s)
            End If
            
            If Val(Right(a, 2)) >= 1 And Val(Right(a, 2)) < 20 Then
                numWor = numWor & convert(Val(Right(a, 2)))

            End If
            numWor = numWor & "Paise "
End If

numWor = StrConv(numWor, vbProperCase)
End Function

Public Function convert(a As String) As String
Dim n
n = Val(a)
Select Case n
Case 0: convert = ""
Case 1: convert = "one "
Case 2: convert = "two "
Case 3: convert = "three "
Case 4: convert = "four "
Case 5: convert = "five "
Case 6: convert = "six "
Case 7: convert = "seven "
Case 8: convert = "eight "
Case 9: convert = "nine "
Case 10: convert = "ten "
Case 11: convert = "eleven "
Case 12: convert = "twelve "
Case 13: convert = "thirteen "
Case 14: convert = "fouteen "
Case 15: convert = "fifteen "
Case 16: convert = "sixteen "
Case 17: convert = "seventeen "
Case 18: convert = "eighteen "
Case 19: convert = "nineteen "
Case 20: convert = "twenty "
Case 30: convert = "thirty "
Case 40: convert = "fourty "
Case 50: convert = "fifty "
Case 60: convert = "sixty "
Case 70: convert = "seventy "
Case 80: convert = "eighty "
Case 90: convert = "ninety "
End Select
End Function
