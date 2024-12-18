Attribute VB_Name = "Module2"



Public Sub isValidInput(ctrl As TextBox, packet As String, keyascii As Integer, strlen As Integer)
        Select Case packet
            Case "digit"
                If Not ((keyascii >= 48 And keyascii <= 57) Or keyascii = vbKeyBack Or keyascii = vbKeyReturn) Then
                    keyascii = 0 ' Simulate a backspace
                    response = MsgBox("Numeric Digits (0,1,2,3...9) only", vbExclamation)
                    'ctrl.Text = Left(ctrl.Text, Len(ctrl.Text) - 1)
                End If
            
            Case "decimal"
                If Not ((keyascii >= 48 And keyascii <= 57) Or keyascii = vbKeyBack Or keyascii = vbKeyReturn Or keyascii = 46) Then
                           keyascii = 0 ' Simulate a backspace
                           response = MsgBox("decimal format (0.00,1.01,...9.99) only", vbExclamation)
                ElseIf (keyascii = 46) Then
                ' Ensure only one decimal point is allowed
                    If InStr(ctrl.Text, ".") > 0 Then
                        keyascii = 0 'Simulate a backspace
                    End If
                End If
 
            Case "alphabet"
                If Not ((keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 132) Or (keyascii = vbKeySpace) Or (keyascii = vbKeyBack) Or (keyascii = vbKeyReturn)) Then
                    keyascii = 0 ' Simulate a backspace
                    response = MsgBox("Alphabets (a,b,c,...,z) (A,B,C,...Z) only", vbExclamation)
                End If
            Case "length"
                If (Len(ctrl.Text) > strlen) And (keyascii <> vbKeyBack) Then
                    keyascii = 0 ' Simulate a backspace
                    response = MsgBox("Length Exceeds!!!", vbExclamation)
                End If
            Case "email"
                'on lost focus event of textbox
                Dim emailPattern As String
                emailPattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$"
                If Not ctrl.Text Like emailPattern Then
                response = MsgBox("Please enter a valid email address.", vbExclamation)
                End If
            Case "special"
            
            Case "empty"
                If Trim(ctrl.Text) = "" Then
                    ctrl.SetFocus
                    response = MsgBox("The Field is Mandatory", vbExclamation)
                End If
        End Select
    End Sub

Public Sub enterKeyPress(t1 As Control, t2 As Control, keyascii As Integer)
If keyascii = vbKeyReturn Then
t2.SetFocus
End If
End Sub

Public Sub cmdOnOff(response As Boolean)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is CommandButton Then
         ctrl.Enabled = response
        End If
    Next ctrl
End Sub
Public Sub TextBoxes(packet As String, state As Boolean)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            Select Case packet
                Case "blank"
                    ctrl.Text = ""
                Case "locked"
                    ctrl.Locked = state
            End Select
        End If
    Next ctrl
End Sub

Public Sub ErrHndlCode()
'Informing User about the Runtime Error
    Dim response As Integer
    response = MsgBox("Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.Description & Chr(13) & Chr(10) & _
        "Error Source: " & Err.Source, vbCritical, "AIC_SOLUTIONS ___ Something Unexpected has Occured! ")
    
    'clear the Error
    Err.Clear
End Sub
