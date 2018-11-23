VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "About CLAM"
   ClientHeight    =   5175
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   5295
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3571.877
   ScaleMode       =   0  'User
   ScaleWidth      =   4972.279
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000005&
      Height          =   1620
      ItemData        =   "frmAbout.frx":0000
      Left            =   240
      List            =   "frmAbout.frx":0019
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Go &Back"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4958.193
      Y1              =   3561.524
      Y2              =   3561.524
   End
   Begin VB.Line Line3 
      X1              =   4958.193
      X2              =   4958.193
      Y1              =   0
      Y2              =   3727.176
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3561.524
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4958.193
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Height          =   2895
      Left            =   4920
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   3735
      Left            =   4560
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Height          =   2295
      Left            =   4320
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "http://www.spt.sit.co.in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmAbout.frx":0181
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   $"frmAbout.frx":048B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label LblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "CLAM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'About Form
'for Showing Credits


Private Sub cmdExit_Click()

    Unload Me
    Unload frmMain

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub cmdSysInfo_Click()

    ShellEx Label4.Caption

End Sub

Private Sub Form_Load()

    Me.Icon = frmMain.ImgList32.ListImages(5).Picture
    Image1.Picture = frmMain.Icon
    
End Sub

Private Sub Label4_Click()

    ShellEx Label4.Caption
    
End Sub

