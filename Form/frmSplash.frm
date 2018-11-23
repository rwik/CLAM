VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawMode        =   8  'Xor Pen
      DrawStyle       =   2  'Dot
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   360
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   1335
      ScaleWidth      =   2175
      TabIndex        =   19
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go"
      Height          =   255
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1320
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1080
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   3480
      Width           =   7575
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER PASSWORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "anagement"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "utomation and"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ibrary"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ollege"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aavik Amit Rwik Vinay "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "visit :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "http:\\www.spt.sit.nic.in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Height          =   1575
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Height          =   1695
      Left            =   2880
      TabIndex        =   14
      Top             =   1800
      Width           =   2895
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

 If Not Text1.Text = varUserPassword Then
    MsgBox "Invalid Password.Try again.", vbExclamation, "CLAM"
Else
    Me.Visible = False
    frmMain.Show
    Unload Me
End If
 
End Sub

Private Sub Form_Load()
'Me.MousePointer = vbHourglass
Call UserPassword

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Me.MousePointer = vbDefault
frmMain.Show

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click


End If

 
End Sub

Private Sub Timer2_Timer()
Label5.Visible = True


Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label6.Visible = True


Timer4.Enabled = True

Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Label8.Visible = True


Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Label9(0).Visible = True
Timer5.Enabled = False
End Sub
