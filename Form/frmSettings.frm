VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Save and close"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "1"
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "30"
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Click here for advanced settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "What is the ammount of fines per day enforced if a book is not returned on time?"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "What is the maximum number of days a book can be kept before the fines are generataed?"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form will record Maximum number of days a book can be borrowed
'before the fine is placed and the amount of fine imposed ber day the
'book is late in the registry.

Private Sub Command1_Click()

    On Error GoTo eeee

    If Text1.Text = "" Or IsNumeric(Text1.Text) = False Or Text1.Text < 0 Or Text2.Text = "" Or IsNumeric(Text2.Text) = False Or Text2.Text < 0 Then
        GoTo eeee
        Exit Sub
    Else
        SaveSetting App.Title, "Settings", "Fine Amount", CStr(CCur(Text2.Text))
        SaveSetting App.Title, "Settings", "Max Days", CStr(CCur(Text1.Text))
        Unload Me
    End If

Exit Sub

eeee:

    MsgBox "You have entered an invalid charecter or no charecters at all in the textboxes" & vbNewLine & "therefore you cannot save the settings" & vbNewLine & "You can enter only numeric data in the boxes", vbExclamation

End Sub

Private Sub Form_Load()

    Text2.Text = GetSetting(App.Title, "Settings", "Fine Amount", "2")
    Text1.Text = GetSetting(App.Title, "Settings", "Max Days", "14")

End Sub

Private Sub Label3_Click()
frmDel.Show vbModal

End Sub
