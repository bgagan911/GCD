Attribute VB_Name = "Control_Mod"
Public Sub ClearTextBoxes(frmClearMe As Form)
'**************************************
'Clears all the TEXT Boxes on the form
'**************************************

 Dim txt As Control

'clear the text boxes
 For Each txt In frmClearMe
  If TypeOf txt Is TextBox Then txt.Text = "00"
 Next

End Sub

Public Sub eofbof(rs As ADODB.Recordset)
    'Monitors the Errors for BOF / EOF
        
        If rs.EOF = True Then
            rs.MoveLast
        Else
        If rs.BOF = True Then
            rs.MoveFirst
        End If
        End If
    
End Sub

Public Function mMsgBox(str As String)
    'My Message Box Function for Error Handling
    
    MsgBox str, vbCritical, "Error Handler"
    
'    & Chr(10) & Chr(10) & "- More Information -" & Chr(10) & _
'    Chr(10) & "Error Number :" & Err.Number & Chr(10) & Chr(10) & _
'    "Error Description: " & Err.Description

End Function

Public Sub errHandler()
    
    With Form1
    'Check if the Values are Numeric or Not...?
    If Not IsNumeric(.txtdeg_lat1.Text) Or Not IsNumeric(.txtDeg_lat2.Text) Or _
    Not IsNumeric(.txtdeg_long1.Text) Or Not IsNumeric(.txtDeg_long2.Text) Or _
    Not IsNumeric(.txtmin_lat1.Text) Or Not IsNumeric(.txtMin_lat2.Text) Or _
    Not IsNumeric(.txtmin_long1.Text) Or Not IsNumeric(.txtMin_long2.Text) Or _
    Not IsNumeric(.txtsec_lat1.Text) Or Not IsNumeric(.txtSec_lat2.Text) Or _
    Not IsNumeric(.txtsec_long1.Text) Or Not IsNumeric(.txtSec_long2.Text) Then
        
        mMsgBox "Only Numbers Allowed."
        ClearTextBoxes Form1
    End If
    End With
    
    'Small Error Handler - Checks for Input-ted values only
    With Form1
        If (.txtdeg_lat1.Text Or .txtDeg_lat2.Text Or .txtdeg_long1.Text Or _
        .txtDeg_long2.Text) > 99 Then
            mMsgBox "Please Enter the Degree Part Only..." & Chr(10) & _
            Chr(10) & "Only 2-Digits Allowed"
            ClearTextBoxes Form1
        End If
        If (.txtmin_lat1.Text Or .txtMin_lat2.Text Or .txtmin_long1.Text Or _
        .txtMin_long2.Text) > 99 Then
            mMsgBox "Please Enter the Minute Part Only..." & Chr(10) & _
            Chr(10) & "Only 2-Digits Allowed"
            ClearTextBoxes Form1
        End If
        If (.txtsec_lat1.Text Or .txtSec_lat2.Text Or .txtsec_long1.Text Or _
        .txtSec_long2.Text) > 99 Then
            mMsgBox "Please Enter the Second Part Only..." & Chr(10) & _
            Chr(10) & "Only 2-Digits Allowed"
            ClearTextBoxes Form1
        End If
    End With
    
    
    
End Sub
