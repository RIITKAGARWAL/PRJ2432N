Attribute VB_Name = "Module2"
Option Explicit

Public Function ConvertNumberToText(ByVal number As Long) As String
    Dim units() As String
    Dim tens() As String
    
    units = Split("Zero One Two Three Four Five Six Seven Eight Nine Ten Eleven Twelve Thirteen Fourteen Fifteen Sixteen Seventeen Eighteen Nineteen")
    tens = Split("Zero Ten Twenty Thirty Forty Fifty Sixty Seventy Eighty Ninety")
    
    If number = 0 Then
        ConvertNumberToText = "Zero"
        Exit Function
    End If
    
    If number < 20 Then
        ConvertNumberToText = units(number)
        Exit Function
    End If
    
    If number < 100 Then
        ConvertNumberToText = tens(number \ 10)
        If number Mod 10 <> 0 Then
            ConvertNumberToText = ConvertNumberToText & " " & units(number Mod 10)
        End If
        Exit Function
    End If
    
    If number < 1000 Then
        ConvertNumberToText = units(number \ 100) & " Hundred"
        If number Mod 100 <> 0 Then
            ConvertNumberToText = ConvertNumberToText & " and " & ConvertNumberToText(number Mod 100)
        End If
        Exit Function
    End If
    
    If number < 100000 Then
        ConvertNumberToText = ConvertNumberToText(number \ 1000) & " Thousand"
        If number Mod 1000 <> 0 Then
            ConvertNumberToText = ConvertNumberToText & " " & ConvertNumberToText(number Mod 1000)
        End If
        Exit Function
    End If
    
    If number < 10000000 Then
        ConvertNumberToText = ConvertNumberToText(number \ 100000) & " Lakh"
        If number Mod 100000 <> 0 Then
            ConvertNumberToText = ConvertNumberToText & " " & ConvertNumberToText(number Mod 100000)
        End If
        Exit Function
    End If
    
    ConvertNumberToText = "Number out of range"
End Function


