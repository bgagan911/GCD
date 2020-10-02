Attribute VB_Name = "dbCon"
'Transmit Informations between forms...
    
    'Make a Connection Variable...
    Global cn As New ADODB.Connection
    'Make a recordset variable...
    Global rs As New ADODB.Recordset
    'Make a connection string Variable...
    Global str As String
    
    'Information Exchange Variable...
    Global var(10)
    'Parameter pass variable for Point 1 / 2
    Global pass As Integer

Public Sub cp_vars()
    'Copies values from Form2->DB....

        With rs
        'Copy Latitude Information
            var(0) = .Fields(3)
            var(1) = .Fields(4)
            var(2) = .Fields(5)
            If .Fields(6) = "n" Then
                var(3) = 1
            Else: var(3) = 0
            End If
            
        'Copy Longitude Information
            var(4) = .Fields(7)
            var(5) = .Fields(8)
            var(6) = .Fields(9)
            If .Fields(10) = "e" Then var(7) = 1
        End With
        
        ' Copy the Button Caption / State Name
        ' Defined there itself
        
End Sub

Public Sub paste_vars(pass)
    'Paste the Values in Form1...
    
    If pass = 1 Then
    '==============================   Point 1  ===========
        With Form1
            
        'Fill Selected Name
            .Command1.Caption = var(8)
            
        'Fill Latitude Information
            .txtdeg_lat1 = var(0)
            .txtmin_lat1 = var(1)
            .txtsec_lat1 = var(2)
            If var(3) = 1 Then
                .cmbDirect_lat1.ListIndex = 0
            Else: .cmbDirect_lat1.ListIndex = 1
            End If
        
        'Fill Longitude Information
            .txtdeg_long1 = var(4)
            .txtmin_long1 = var(5)
            .txtsec_long1 = var(6)
            If var(7) = 1 Then
                .cmbDirect_long1.ListIndex = 0
            Else: .cmbDirect_long1.ListIndex = 1
            End If

        End With
    Else
    '==============================   Point 2  ===========
    If pass = 2 Then
    With Form1
            
        'Fill Selected Name
            .Command2.Caption = var(8)
            
        'Fill Latitude Information
            .txtDeg_lat2 = var(0)
            .txtMin_lat2 = var(1)
            .txtSec_lat2 = var(2)
            If var(3) = 1 Then
                .cmbDirect_lat2.ListIndex = 0
            Else: .cmbDirect_lat2.ListIndex = 1
            End If
        
        'Fill Longitude Information
            .txtDeg_long2 = var(4)
            .txtMin_long2 = var(5)
            .txtSec_long2 = var(6)
            If var(7) = 1 Then
                .cmbDirect_long2.ListIndex = 0
            Else: .cmbDirect_long2.ListIndex = 1
            End If

        End With
    End If
    End If
End Sub


