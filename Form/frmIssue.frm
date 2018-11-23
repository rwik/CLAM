VERSION 5.00
Begin VB.Form frmIssue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Books Issue Form"
   ClientHeight    =   5175
   ClientLeft      =   4860
   ClientTop       =   3855
   ClientWidth     =   5790
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
   ScaleHeight     =   5175
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton cmdBook 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Browse for record"
      Top             =   2280
      Width           =   315
   End
   Begin VB.CommandButton cmdCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Browse record."
      Top             =   1320
      Width           =   315
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00F4FEFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00F4FEFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "&Issue Book"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Date to be returned:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Date Issued:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Book Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Student Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmIssue.frx":0000
      Height          =   855
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Book ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Student Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   975
      Y2              =   975
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Books Issue Form
'This form is used to issue books that are present in the library.
'and recorded in the database via the issue form.


Private Sub cmdBook_Click()

    With frmSelectDg
        .CommandText = "Select * From tblBooks where Borrowed=False"
        .DataGrid1.Caption = "Members Table"
        .Show vbModal
        If .OKPressed Then
            Text5.Text = .rRS1
            Text2.Text = .rRS2
        End If
    End With
      cmdIssue.SetFocus
      
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdCode_Click()

Dim A As String, b As String, c As String

    With frmSelectDg
        .CommandText = "Select * From tblMembers where id not like 't%'"
        .DataGrid1.Caption = "Members Table"
        .Show vbModal
        If .OKPressed Then
            Text4.Text = .rRS1
            A = .rRS9
            'b = .rRS3
            'c = .rRS4
            Text1.Text = A '& " " & b & " " & c
        End If
    End With

End Sub

Private Sub cmdIssue_Click()

'Record that the book was taken in two places. In tblTrans, and in
'tblBooks which will set the Borrowed Boolean to True.

Dim RS As ADODB.RecordSet

    If Text4.Text = "" Then Text4.SetFocus: Exit Sub
    If Text5.Text = "" Then Text5.SetFocus: Exit Sub
    On Error GoTo hell
    CN.BeginTrans
    Set RS = New ADODB.RecordSet
    With RS
        .Open "Select * from tblTrans", CN, adOpenDynamic, adLockOptimistic
        .AddNew
        .Fields(0) = Text5.Text
        .Fields(1) = Text4.Text
        .Fields(2) = Date
        .Update
        .Close

        .Open "Select [Borrowed] from tblBooks where [Accesion Number]='" & Text5.Text & "'", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields(0) = True
        .Update
        .Close
        Set RS = Nothing
    End With
    CN.CommitTrans
    If MsgBox("The book " & Text5.Text & " has been issued to " & Text4.Text & vbNewLine & "Do you want to create a new issue instance?", vbInformation + vbYesNo) = vbYes Then
        cmdReset_Click
    Else
        Unload Me
    End If

Exit Sub

hell:
    Handler Err
    CN.RollbackTrans

End Sub

Private Sub cmdReset_Click()

    Text1.Text = ""
    Text2.Text = ""
    Text5.Text = ""
    Text4.Text = ""
    Text3.Text = FormatDateTime$(Date, vbLongDate)
    Text6.Text = FormatDateTime$(Date + frmReturn.MaxDays, vbLongDate)

End Sub

Private Sub Form_Load()

    cmdReset_Click
    With frmMain
        cmdCode.Picture = .ImgList16.ListImages(1).Picture
        Me.Icon = .ImgList32.ListImages(7).Picture
    End With
    cmdBook.Picture = cmdCode.Picture
    Image1.Picture = Me.Icon

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

    cmdCode_Click

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

    cmdBook_Click

End Sub
