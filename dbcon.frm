VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Select Place ::"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dbcon.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text9"
      Top             =   750
      Width           =   1005
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2625
      Width           =   5295
   End
   Begin VB.ComboBox cmb1 
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   195
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1905
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1305
      Width           =   1380
   End
   Begin VB.Label Label4 
      Caption         =   "Owner:"
      Height          =   330
      Left            =   135
      TabIndex        =   8
      Top             =   810
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   5430
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Longitude Information"
      Height          =   360
      Left            =   345
      TabIndex        =   7
      Top             =   1935
      Width           =   2250
   End
   Begin VB.Label Label2 
      Caption         =   "Latitude Information"
      Height          =   360
      Left            =   345
      TabIndex        =   6
      Top             =   1320
      Width           =   2070
   End
   Begin VB.Label Label1 
      Caption         =   "Select a Place: "
      Height          =   450
      Left            =   135
      TabIndex        =   5
      Top             =   240
      Width           =   1845
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    App.Path & "\gcd.mdb;Persist Security Info=False"
    
    'Check Connection State *********
    If cn.State = True Then
        Exit Sub
    Else
    If cn.State = False Then
        cn.Open (str)
    End If
    End If
    '*********************************
    
    rs.Open "SELECT name FROM place ORDER BY" & _
    " name", cn, adOpenDynamic, adLockBatchOptimistic
    
   ' Load the ComboBox.
    rs.MoveFirst
    Do While Not rs.EOF
        cmb1.AddItem rs!Name
        rs.MoveNext
    Loop
    
    rs.Close
    
    'Select the first choice.
     cmb1.ListIndex = 0
    
End Sub

Private Sub cmb1_Click()
      
      'Open-up the recordset
        rs.Open "select * from place where name='" & cmb1.Text & "'"
     
      'Fill Owner Information
        Text9.Text = rs.Fields(2)
     
      'Fill Latitude Information
        Text1.Text = rs.Fields(3) & "-" & rs.Fields(4) & _
        "-" & rs.Fields(5) & " " & UCase(rs.Fields(6))
    
      'Fill Longitude Information
        Text2.Text = rs.Fields(7) & "-" & rs.Fields(8) & _
        "-" & rs.Fields(9) & " " & UCase(rs.Fields(10))
      
      'Copy all the required values
        cp_vars
        var(8) = cmb1.Text
        
      'Change Command Caption
        cmdSelect.Caption = "&Select " & cmb1.Text
    
      'Close Recordset
        rs.Close
End Sub

Private Sub cmdSelect_Click()
    
    paste_vars (pass)
    Form1.Show
    Unload Me

End Sub
