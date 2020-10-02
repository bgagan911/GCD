Attribute VB_Name = "Main_Module"
Global choice As Integer

Public Sub cmbInit()
    'Initialisez combo boxes
    Form1.cmbDirect_lat1.AddItem ("N")
    Form1.cmbDirect_lat1.AddItem ("S")
    
    Form1.cmbDirect_lat2.AddItem ("N")
    Form1.cmbDirect_lat2.AddItem ("S")
    
    Form1.cmbDirect_long1.AddItem ("E")
    Form1.cmbDirect_long1.AddItem ("W")
    
    Form1.cmbDirect_long2.AddItem ("E")
    Form1.cmbDirect_long2.AddItem ("W")
End Sub

Public Sub pureDegree()
    'Converts the Degree-Min-Sec to pure Degree form
   
    Dim temp As Single
    Dim temp_Lat2 As Single
    Dim temp_Long1 As Single
    Dim temp_Long2 As Single
    
    'Convert Latitude 1
    temp = Form1.txtdeg_lat1
    temp = temp + Form1.txtmin_lat1 / 60
    temp = temp + Form1.txtsec_lat1 / 3600
    'Convert Latitude 2
    temp_Lat2 = Form1.txtDeg_lat2
    temp_Lat2 = temp_Lat2 + Form1.txtMin_lat2 / 60
    temp_Lat2 = temp_Lat2 + Form1.txtSec_lat2 / 3600
    'Convert Longitude 1
    temp_Long1 = Form1.txtdeg_long1
    temp_Long1 = temp_Long1 + Form1.txtmin_long1 / 60
    temp_Long1 = temp_Long1 + Form1.txtsec_long1 / 3600
    'Convert Longitude 2
    temp_Long2 = Form1.txtDeg_long2
    temp_Long2 = temp_Long2 + Form1.txtMin_long2 / 60
    temp_Long2 = temp_Long2 + Form1.txtSec_long2 / 3600
    
    Form1.Text1 = temp              'Latitude 1
    Form1.Text2 = temp_Lat2         'Latitude 2
    Form1.Text3 = temp_Long1        'Longitude 1
    Form1.Text4 = temp_Long2        'Longitude 2

End Sub

Public Sub deg2rad()
    'Converts the Pure-Degree Data to Radians Notation
    
    Form1.Text1_RAD = Form1.Text1 * (22 / 7) / 180
    Form1.Text2_RAD = Form1.Text2 * (22 / 7) / 180
    Form1.Text3_RAD = Form1.Text3 * (22 / 7) / 180
    Form1.Text4_RAD = Form1.Text4 * (22 / 7) / 180
End Sub

Public Sub GCD_(choice As Integer)
    'Calculates the distance between the 2 points
    'Requires the data to be in Radians
    
    Dim temp As Double
    Dim temp1(10) As Double
 
    temp = Sin((Form1.Text1_RAD - Form1.Text2_RAD) / 2)
    temp = temp * temp
    
    temp1(1) = Cos(Form1.Text1_RAD)
    temp1(2) = Cos(Form1.Text2_RAD)
    temp1(3) = temp1(1) * temp1(2)
    
    temp1(4) = Sin((Form1.Text3_RAD - Form1.Text4_RAD) / 2)
    temp1(5) = temp1(4) * temp1(4)
    temp1(6) = temp1(3) * temp1(5)
    
    temp1(7) = temp + temp1(6)
    temp1(8) = Sqr(temp1(7))
    
    'Equivalent of ArcSine
    temp1(9) = ArcSin(temp1(8))
    temp = 2 * temp1(9)

If choice = 1 Then
    'Distance in Radians
    Form1.txtResult = temp
Else
'**********************************************************
If choice = 2 Then
    
    temp = temp * 6364.1454545
    'Distance in Kms
    Form1.txtResult = temp
    
End If
End If

End Sub

'Public Sub gcd2kms()
'    Form1.txtResult = Form1.txtResult * 6364.1454545
'End Sub

Function ArcSin(X As Double) As Double
    ArcSin = Atn(X / Sqr(-X * X + 1))
End Function
