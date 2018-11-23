VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   3120
   ClientLeft      =   4665
   ClientTop       =   3675
   ClientWidth     =   5190
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSearch.frx":0CCA
      Left            =   120
      List            =   "frmSearch.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1935
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   2655
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1215
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Look in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearch.frx":0CCE
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSearch.frx":1598
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   165
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Search for:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   975
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   4920
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   4920
      Y1              =   2415
      Y2              =   2415
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Search Form
'This reusable form can be called from anywhere to search a recordset
'according to given criteria and then move the absolute position of the
'record. Includes error trapping to search in different data type
'formats.


Public SourceRs            As ADODB.RecordSet
Private AlreadyFilled      As Boolean
Private AlreadySearched    As Boolean
Private CurrPos            As Long
Private oldpos             As Long

Private Sub Combo1_KeyPress(Keyascii As Integer)

    Keyascii = 0

End Sub

Private Sub Command1_Click()

    On Error GoTo Err
    If Text1.Text = "" Then Text1.SetFocus: Exit Sub
    If Combo1.Text = "" Then Combo1.SetFocus: Exit Sub
    With SourceRs
        If AlreadySearched = False Then
            oldpos = .AbsolutePosition
            .MoveFirst
            .Find "[" & Combo1.Text & "] like *" & Text1.Text & "*"
            CurrPos = .AbsolutePosition
            If .EOF Then
                MsgBox "Could not find '" & Text1.Text & "' in '" & Combo1.Text & "'.", vbExclamation
                .AbsolutePosition = oldpos
            Else
                AlreadySearched = True
                Command1.Caption = "Search Next"
            End If
        Else
            oldpos = .AbsolutePosition
            .MoveNext
            .Find "[" & Combo1.Text & "] like *" & Text1.Text & "*"
            CurrPos = .AbsolutePosition
            If .EOF Then MsgBox "Search completed.", vbInformation: AlreadySearched = False: .AbsolutePosition = oldpos
        End If
    End With

Exit Sub

Err:
    If Err.Number = -2147217881 Then Search_Number: Resume Next
    If Err.Number = 3265 Then MsgBox "Please select a valid section from the list", vbExclamation: HighLight Text1: Exit Sub
    Handler Err

End Sub

Private Sub Search_Number()

'For Number data type

    On Error GoTo Err
    SourceRs.Find "[" & Combo1.Text & "] like " & Text1.Text & ""

Exit Sub

Err:
    Search_DateTime

End Sub

Private Sub Search_DateTime()

'For Date/Time data type

    On Error GoTo Err
    SourceRs.Find "[" & Combo1.Text & "] like #" & Text1.Text & "#"

Exit Sub

Err:
    MsgBox "Please enter an appropriate value that correspand" & vbCrLf & "where to find it (ex.Search for 10/23/1985 and Look in Date).", vbExclamation

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    FillCombo Combo1, SourceRs, False
    Me.Icon = Image1.Picture
    Combo1.ListIndex = 0

End Sub

Private Sub Text1_Change()

    AlreadySearched = False

End Sub
