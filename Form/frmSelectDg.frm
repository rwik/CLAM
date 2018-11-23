VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSelectDg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please select a record from the table"
   ClientHeight    =   5820
   ClientLeft      =   2640
   ClientTop       =   2745
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelect 
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      Picture         =   "frmSelectDg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   6120
      Picture         =   "frmSelectDg.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   615
   End
   Begin VB.Frame fraNavigation 
      Height          =   800
      Left            =   7560
      TabIndex        =   8
      Top             =   4800
      Width           =   1725
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   3
         Left            =   1290
         Picture         =   "frmSelectDg.frx":1434
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   2
         Left            =   870
         Picture         =   "frmSelectDg.frx":17BE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   1
         Left            =   450
         Picture         =   "frmSelectDg.frx":1B48
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   0
         Left            =   30
         Picture         =   "frmSelectDg.frx":1ED2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   400
      End
      Begin VB.TextBox txtcount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "  of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblmax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   990
         TabIndex        =   14
         Top             =   480
         Width           =   700
      End
   End
   Begin VB.CommandButton cmdOperations 
      Height          =   615
      Index           =   2
      Left            =   5400
      Picture         =   "frmSelectDg.frx":225C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdOperations 
      Height          =   615
      Index           =   0
      Left            =   4680
      Picture         =   "frmSelectDg.frx":2B26
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdOperations 
      Height          =   615
      Index           =   1
      Left            =   3960
      Picture         =   "frmSelectDg.frx":37F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   6840
      Picture         =   "frmSelectDg.frx":40BA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4380
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7726
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Table Name"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   5
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   5520
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   9360
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   5520
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   4680
      Y2              =   4680
   End
End
Attribute VB_Name = "frmSelectDg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The selection dialog box
'This form can be called from anywhere in the application with the
'correct prameters and shows a selection box for recordsets. This
'can be used to return record values.
'This form is called from frmIssue and frmReturn to select members and
'books.

Public CommandText As String, OKPressed As Boolean
Public rRS1 As String, rRS2 As String, rRS3 As String, rRS4 As String, rRS9 As String

Private RS As ADODB.RecordSet

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdNavigate_Click(Index As Integer)

    Navigate Index, RS
    DisplayRecords
    

End Sub

Private Sub cmdRefresh_Click()

    With RS
        .Filter = adFilterNone
        .Requery
    End With

End Sub

Private Sub cmdSelect_Click()

    On Error Resume Next
        With RS
            If .RecordCount < 1 Then MsgBox "No record to select!" & vbNewLine & "Please add records to the library first to select data from them.", vbExclamation, "No data Selected": Exit Sub
            rRS1 = .Fields(0)
            rRS2 = .Fields(2)
           rRS3 = .Fields(3)
            rRS4 = .Fields(4)
            rRS9 = .Fields(1)
        End With
        CommandText = ""
        OKPressed = True
        Unload Me

End Sub

Private Sub DataGrid1_DblClick()

    cmdSelect_Click

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Or KeyCode = 40 Then DisplayRecords

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    DisplayRecords

End Sub

Private Sub Form_Load()

'CommandText = "SELECT tblTrans.[Accesion Number], tblTrans.[ID], tblBooks.[Title], [ Name] AS Borrower, tblTrans.[Date_Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Accesion Number] = tblTrans.[Accesion Number]) ON tblMembers.[ID] = tblTrans.[ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Accesion Number];"
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open CommandText, CN, adOpenDynamic, adLockOptimistic
    
    DisplayRecords
    Me.Icon = cmdSelect.Picture
    Set DataGrid1.DataSource = RS
    OKPressed = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set RS = Nothing

End Sub

Private Sub cmdOperations_Click(Index As Integer)

Dim obj As Form

    If Index = 1 Then Set obj = frmSearch
    If Index = 0 Then Set obj = frmFilter
    If Index = 2 Then Set obj = frmSort

    With obj
        Set .SourceRs = RS
        .Show vbModal
    End With
    Set obj = Nothing

End Sub

Private Sub DisplayRecords()

'-Display the current and total number of record

    On Error GoTo hell
    With RS
        If .RecordCount < 1 Then
            txtcount.Text = 0
        Else
            txtcount.Text = .AbsolutePosition
        End If
        lblmax.Caption = .RecordCount
    End With

Exit Sub

hell:
    Handler Err

End Sub
