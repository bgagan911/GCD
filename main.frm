VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..:: Great Circle Distance Calculator ::.."
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   9960
      TabIndex        =   62
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   495
      Left            =   7320
      TabIndex        =   61
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   495
      Left            =   8640
      TabIndex        =   60
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "System &Info"
      Height          =   495
      Left            =   6000
      TabIndex        =   59
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Radians Notation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   5850
      TabIndex        =   41
      Top             =   3225
      Width           =   5490
      Begin VB.TextBox Text1_RAD 
         Height          =   375
         Left            =   960
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   765
         Width           =   2025
      End
      Begin VB.TextBox Text2_RAD 
         Height          =   375
         Left            =   960
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   1530
         Width           =   2025
      End
      Begin VB.TextBox Text3_RAD 
         Height          =   375
         Left            =   3240
         TabIndex        =   47
         Text            =   "Text3"
         Top             =   765
         Width           =   2025
      End
      Begin VB.TextBox Text4_RAD 
         Height          =   375
         Left            =   3240
         TabIndex        =   46
         Text            =   "Text4"
         Top             =   1530
         Width           =   2025
      End
      Begin VB.Label Label24 
         Caption         =   "Latitude : 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   57
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Longitude : 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   56
         Top             =   390
         Width           =   1500
      End
      Begin VB.Label Label22 
         Caption         =   "Latitude : 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   975
         TabIndex        =   55
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Longitude : 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   54
         Top             =   1215
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pure Degree Notation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   120
      TabIndex        =   40
      Top             =   3225
      Width           =   5670
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1050
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   765
         Width           =   2025
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1065
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   1530
         Width           =   2025
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3330
         TabIndex        =   43
         Text            =   "Text3"
         Top             =   765
         Width           =   2025
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3330
         TabIndex        =   42
         Text            =   "Text4"
         Top             =   1530
         Width           =   2025
      End
      Begin VB.Label Label20 
         Caption         =   "Longitude : 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3345
         TabIndex        =   53
         Top             =   1215
         Width           =   1485
      End
      Begin VB.Label Label19 
         Caption         =   "Latitude : 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   52
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Longitude : 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3345
         TabIndex        =   51
         Top             =   390
         Width           =   1500
      End
      Begin VB.Label Label17 
         Caption         =   "Latitude : 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1065
         TabIndex        =   50
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   105
      TabIndex        =   36
      Top             =   5370
      Width           =   5670
      Begin VB.CommandButton cmdCheck 
         Caption         =   "&Check"
         Height          =   495
         Left            =   4050
         TabIndex        =   58
         Top             =   1035
         Width           =   1215
      End
      Begin VB.TextBox txtResult 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1170
         TabIndex        =   39
         Top             =   1170
         Width           =   2505
      End
      Begin VB.OptionButton Opt_Kms 
         Caption         =   "Distance in Kms"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1125
         TabIndex        =   38
         Top             =   615
         Width           =   2760
      End
      Begin VB.OptionButton Opt_RAD 
         Caption         =   "Distance in Radians"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1125
         TabIndex        =   37
         Top             =   150
         Width           =   3375
      End
   End
   Begin VB.TextBox txtDeg_lat2 
      Height          =   360
      Left            =   7485
      TabIndex        =   9
      Text            =   "27"
      Top             =   720
      Width           =   930
   End
   Begin VB.TextBox txtMin_lat2 
      Height          =   360
      Left            =   7485
      TabIndex        =   10
      Text            =   "10"
      Top             =   1200
      Width           =   930
   End
   Begin VB.TextBox txtSec_lat2 
      Height          =   360
      Left            =   7485
      TabIndex        =   11
      Text            =   "00"
      Top             =   1680
      Width           =   930
   End
   Begin VB.TextBox txtDeg_long2 
      Height          =   360
      Left            =   10365
      TabIndex        =   13
      Text            =   "77"
      Top             =   720
      Width           =   930
   End
   Begin VB.TextBox txtMin_long2 
      Height          =   360
      Left            =   10365
      TabIndex        =   14
      Text            =   "58"
      Top             =   1200
      Width           =   930
   End
   Begin VB.TextBox txtSec_long2 
      Height          =   360
      Left            =   10365
      TabIndex        =   15
      Text            =   "00"
      Top             =   1680
      Width           =   930
   End
   Begin VB.ComboBox cmbDirect_lat2 
      Height          =   315
      ItemData        =   "main.frx":5E62
      Left            =   7470
      List            =   "main.frx":5E64
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Width           =   960
   End
   Begin VB.ComboBox cmbDirect_long2 
      Height          =   315
      ItemData        =   "main.frx":5E66
      Left            =   10365
      List            =   "main.frx":5E68
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2160
      Width           =   960
   End
   Begin VB.CommandButton cmdPoint2 
      Caption         =   "Select from &Database"
      Height          =   435
      Left            =   7725
      TabIndex        =   17
      Top             =   2595
      Width           =   1860
   End
   Begin VB.CommandButton cmdPoint1 
      Caption         =   "&Select from Database"
      Height          =   435
      Left            =   1920
      TabIndex        =   8
      Top             =   2595
      Width           =   1860
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Point - 2 Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5865
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   105
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Point - 1 Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   105
      Width           =   5535
   End
   Begin VB.ComboBox cmbDirect_long1 
      Height          =   315
      ItemData        =   "main.frx":5E6A
      Left            =   4560
      List            =   "main.frx":5E6C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2160
      Width           =   960
   End
   Begin VB.ComboBox cmbDirect_lat1 
      Height          =   315
      ItemData        =   "main.frx":5E6E
      Left            =   1665
      List            =   "main.frx":5E70
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   960
   End
   Begin VB.TextBox txtsec_long1 
      Height          =   360
      Left            =   4560
      TabIndex        =   6
      Text            =   "00"
      Top             =   1680
      Width           =   930
   End
   Begin VB.TextBox txtmin_long1 
      Height          =   360
      Left            =   4545
      TabIndex        =   5
      Text            =   "07"
      Top             =   1200
      Width           =   930
   End
   Begin VB.TextBox txtdeg_long1 
      Height          =   360
      Left            =   4560
      TabIndex        =   4
      Text            =   "77"
      Top             =   720
      Width           =   930
   End
   Begin VB.TextBox txtsec_lat1 
      Height          =   360
      Left            =   1680
      TabIndex        =   2
      Text            =   "00"
      Top             =   1680
      Width           =   930
   End
   Begin VB.TextBox txtmin_lat1 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Text            =   "34"
      Top             =   1200
      Width           =   930
   End
   Begin VB.TextBox txtdeg_lat1 
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Text            =   "28"
      Top             =   720
      Width           =   930
   End
   Begin VB.Line Line6 
      X1              =   5985
      X2              =   11205
      Y1              =   5775
      Y2              =   5775
   End
   Begin VB.Line Line5 
      X1              =   5985
      X2              =   11220
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line4 
      X1              =   165
      X2              =   11310
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   5760
      X2              =   5760
      Y1              =   105
      Y2              =   3045
   End
   Begin VB.Label Label16 
      Caption         =   "Degree Latitude: "
      Height          =   255
      Left            =   5925
      TabIndex        =   35
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label15 
      Caption         =   "Minute Latitude: "
      Height          =   255
      Left            =   5925
      TabIndex        =   34
      Top             =   1260
      Width           =   1350
   End
   Begin VB.Label Label14 
      Caption         =   "Second Latitude: "
      Height          =   255
      Left            =   5940
      TabIndex        =   33
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Direction: "
      Height          =   255
      Left            =   5925
      TabIndex        =   32
      Top             =   2175
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Direction: "
      Height          =   255
      Left            =   8805
      TabIndex        =   31
      Top             =   2175
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Second Longitude: "
      Height          =   255
      Left            =   8805
      TabIndex        =   30
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Minute Longitude: "
      Height          =   255
      Left            =   8805
      TabIndex        =   29
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Degree Longitude: "
      Height          =   255
      Left            =   8805
      TabIndex        =   28
      Top             =   720
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   8565
      X2              =   8565
      Y1              =   735
      Y2              =   2445
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   2760
      Y1              =   735
      Y2              =   2445
   End
   Begin VB.Label Label8 
      Caption         =   "Degree Longitude: "
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Minute Longitude: "
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Second Longitude: "
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Direction: "
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   2175
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Direction: "
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2175
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Second Latitude: "
      Height          =   255
      Left            =   135
      TabIndex        =   20
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Minute Latitude: "
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1260
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Degree Latitude: "
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
    Form3.Show
End Sub

Private Sub cmdCheck_Click()
    errHandler
    pureDegree
    deg2rad
      
    If Opt_RAD.Value = True Then
        GCD_ (1)
    Else
    If Opt_Kms.Value = True Then
        GCD_ (2)
    End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
MsgBox "Will be available Soon...", vbInformation, "GCD - Help"
End Sub

Private Sub cmdInfo_Click()
    Call StartSysInfo
End Sub

Private Sub cmdPoint1_Click()
    pass = 1
    Form2.Show
End Sub

Private Sub cmdPoint2_Click()
    pass = 2
    Form2.Show
End Sub

Private Sub cmdReset_Click()
    ClearTextBoxes Form1
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtdeg_lat1.SetFocus
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtDeg_lat2.SetFocus
End Sub


Private Sub Form_Load()
    Opt_RAD.Value = True
    Opt_Kms.Value = False
    cmbInit
End Sub

Private Sub Opt_Kms_Click()
    pureDegree
    deg2rad
    GCD_ (2)
End Sub

Private Sub Opt_RAD_Click()
    pureDegree
    deg2rad
    GCD_ (1)
End Sub
