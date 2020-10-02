VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   30
      Picture         =   "about.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   15
      Width           =   2745
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "bgagan@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3255
      TabIndex        =   3
      Top             =   2040
      Width           =   2235
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gagan Bhalla"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      TabIndex        =   2
      Top             =   1440
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By -"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3360
      TabIndex        =   1
      Top             =   525
      Width           =   1875
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub
