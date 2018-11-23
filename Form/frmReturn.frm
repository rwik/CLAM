VERSION 5.00
Begin VB.Form frmReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books Return Form"
   ClientHeight    =   5190
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5400
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
   ScaleHeight     =   5190
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFines 
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
      Left            =   1800
      TabIndex        =   17
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4920
      Picture         =   "frmReturn.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Use calculator"
      Top             =   2400
      Width           =   315
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1680
      Width           =   3375
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1320
      Width           =   3015
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Browse record."
      Top             =   1320
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   4935
      Begin VB.Label Label4 
         Caption         =   "Date borrowed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Days late in returning the book:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Total amount of fine accumulated:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblDate 
         Caption         =   "Select a book first"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblLate 
         Caption         =   "Select a book first"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblFines 
         Caption         =   "Select a book first"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "R&eturn Book"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   4560
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
      Left            =   1920
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Date Returned:"
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
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Fines collected:"
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
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   2535
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
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   5280
      Y1              =   1080
      Y2              =   1080
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
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReturn.frx":038A
      Height          =   615
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   4335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   5160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   5160
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Books Return Form
'This form is used to return books that have been taken from the library
'and recorded in the database via the issue form.

Public MaxDays As Integer
Public FineAmnt As Currency

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdReset_Click()

    lblLate.Caption = "Select a book first"
    lblFines.Caption = "Select a book first"
    lblDate.Caption = "Select a book first"
    txtFines.Text = ""
    txtFines.Locked = True
    Text1.Text = ""
    Text4.Text = ""
    Text2.Text = FormatDateTime$(Date, vbLongDate)

End Sub

Private Sub cmdReturn_Click()

Dim RS As ADODB.RecordSet

'The return information is recorded
'in two places. One in the Book Table where the book Borrowed is set to
'False, and in the Transaction Table where the amount payed and book
'returned is stored with the date the book is returned.

    If Text4.Text = "" Then Text4.SetFocus
    On Error GoTo hello
    Set RS = New ADODB.RecordSet
    With RS
        CN.BeginTrans       'Begin a new transaction
        .Open "Select [Borrowed] from tblBooks where [Accesion Number]='" & Text4.Text & "'", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields(0) = False
        .Update
        .Close

        .Open "Select [Fines],[Returned] From tblTrans where [Accesion Number]='" & Text4.Text & "'" & "And [Returned] = False", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields("Fines") = CCur(txtFines.Text)
        .Fields("Returned") = True
        .Update
        .Close
        CN.CommitTrans      'If no error was raised then record info
    End With
    Set RS = Nothing

    'Show MsgBox if another book needs returning
    If MsgBox("The book " & Text4.Text & " has been returned from " & Text1.Text & vbNewLine & vbNewLine & "Do you want to create a new return book instance?", vbInformation + vbYesNo) = vbYes Then
        cmdReset_Click
    Else
        Unload Me
    End If

Exit Sub

hello:
    Handler Err

    On Error Resume Next    'If an error was raised then rollback
        CN.RollbackTrans

End Sub

Private Sub cmdCode_Click()

Dim RS As ADODB.RecordSet, i As Integer

'The first part of this event procedure will open the frmSelectDg form
'and expect an input from the user.
    On Error Resume Next
        With frmSelectDg
            'First show the box
            .CommandText = "SELECT tblTrans.[Accesion Number], tblTrans.[ID], tblBooks.[Title], tblMembers.[Name] AS Borrower, tblTrans.[Date_Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Accesion Number] = tblTrans.[Accesion Number]) ON tblMembers.[ID] = tblTrans.[ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Accesion Number];"
            .DataGrid1.Caption = "Members Table"
            .Show vbModal

            'display the data
            If .OKPressed Then
                Text4.Text = .rRS1
                Text1.Text = .rRS9
                txtFines.Locked = False
            Else
                'If the user did not enter anything then skip the second
                'part of the procedure to skip errors that may arise because
                'there will be no data (in text4 and text1) and as such
                'null errors or record not found errors.
                Exit Sub
            End If
        End With

        'The second part will calculate the number of days a book was taken out
        'of the library and print it in the txtFines text box.

        Set RS = New ADODB.RecordSet
        RS.Open "Select * from tblTrans Where [Accesion Number] ='" & Text4.Text & "'", CN, adOpenDynamic, adLockOptimistic
        lblDate.Caption = CDate(RS(2))

        'Store the difference of the current date and the date returned
        'in a variable. It the variable is negative it means that the
        'book returned is within the time limit and Fines=i*FineAmnt
        'must be 0. So transform i into 0
        i = Date - CDate(lblDate.Caption)
        If i < 0 Then i = 0
        If MaxDays < i Then lblLate.Caption = i - MaxDays Else lblLate.Caption = "0"

        'Print fines due in a label and a text box
        lblFines.Caption = CStr(FormatCurrency$(FineAmnt * lblLate))

        'Also, use an editable text box so the correct amount a member
        'is payed is recorded. Sometimes the member may pay money not
        'exactly as required
        txtFines.Text = lblFines.Caption
        
        Set RS = Nothing

        'So, librarian  just select a book id through
        'a GUI friendly interface and everything will be done by the system

End Sub

Private Sub Command4_Click()

    On Error GoTo hell
    Shell "calc.exe", vbNormalFocus

Exit Sub

hell:
    MsgBox "The operating system cannot find the system calculator." & vbNewLine & "Please check whether it is properly installed or not", vbCritical, "File not found"

End Sub

Private Sub Form_Load()

    Me.Icon = frmMain.ImgList32.ListImages(8).Picture
    Image1.Picture = Me.Icon
    cmdReset_Click
    cmdCode.Picture = frmMain.ImgList16.ListImages(1).Picture

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

    cmdCode_Click

End Sub
