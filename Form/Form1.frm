VERSION 5.00
Begin VB.Form frmDpissu 
   Caption         =   "Issue Deposit"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000E&
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
      TabIndex        =   15
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "&Deposit"
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
      Left            =   120
      TabIndex        =   8
      Top             =   3720
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
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
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
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
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
      TabIndex        =   5
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000E&
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
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
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
      TabIndex        =   2
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton cmdCode 
      Height          =   315
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Browse record."
      Top             =   1320
      Width           =   315
   End
   Begin VB.CommandButton cmdBook 
      Height          =   315
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Browse for record"
      Top             =   2040
      Width           =   315
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      Caption         =   "Deposit Amount:"
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
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Member ID:"
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
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Accession_Number:"
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
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fill the fields below and click Deposit to issue a book to an existing member."
      Height          =   855
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
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
      Top             =   1680
      Width           =   1575
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
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "frmDpissu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBook_Click()
With frmSelectDg
        .CommandText = "Select * From tblBooks where Borrowed=False"
        .DataGrid1.Caption = "Members Table"
        .Show vbModal
        If .OKPressed Then
            Text5.Text = .rRS1
                    End If
    End With
End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdCode_Click()
Dim A As String, b As String, c As String

    With frmSelectDg
        .CommandText = "Select * From tblMembers"
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
'Record that the book was taken as Depost. In Accesion Register, and in
'tblDepo .Borrowed Boolean is set to True.

Dim RS As ADODB.RecordSet

    If Text4.Text = "" Then Text4.SetFocus: Exit Sub
    If Text5.Text = "" Then Text5.SetFocus: Exit Sub
    If Text2.Text = "" Then Text5.SetFocus: Exit Sub
    On Error GoTo hell
    CN.BeginTrans
    Set RS = New ADODB.RecordSet
    With RS
        .Open "Select * from tblDepo", CN, adOpenDynamic, adLockOptimistic
        .AddNew
        .Fields(0) = Text4.Text
        .Fields(1) = Text1.Text
        .Fields(2) = Text5.Text
        .Fields(3) = Date
        .Fields(4) = Text2.Text
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
    Text2.Text = ""
    Text1.Text = ""
    Text5.Text = ""
    Text4.Text = ""
    Text3.Text = FormatDateTime$(Date, vbLongDate)
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
