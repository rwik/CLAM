VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDeposit 
   Caption         =   "Deposit record"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   3360
      Picture         =   "frmDeposit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton cmdOperations 
      Height          =   615
      Index           =   0
      Left            =   1560
      Picture         =   "frmDeposit.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox txtfrom 
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   17
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   16
      Top             =   7920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Grid View"
      TabPicture(0)   =   "frmDeposit.frx":1594
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reports"
      TabPicture(1)   =   "frmDeposit.frx":15B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdReport(1)"
      Tab(1).Control(1)=   "cmdReport(0)"
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(3)=   "Label21"
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(5)=   "Label23"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   1
         Left            =   -74040
         Picture         =   "frmDeposit.frx":15CC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   0
         Left            =   -74040
         Picture         =   "frmDeposit.frx":2296
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6900
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   12171
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Deposit Details"
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
      Begin VB.Label Label22 
         Caption         =   "Create a custom record based on your search criteria."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -73200
         TabIndex        =   11
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label Label21 
         Caption         =   "Create Custom Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   10
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label20 
         Caption         =   "Create a complete report of all Deposit History."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73200
         TabIndex        =   9
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label Label23 
         Caption         =   "Create Complete Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73200
         TabIndex        =   8
         Top             =   1080
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdRtrn 
      Height          =   615
      Index           =   0
      Left            =   14400
      Picture         =   "frmDeposit.frx":2F60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   615
   End
   Begin VB.CommandButton cmddpst 
      Height          =   615
      Index           =   1
      Left            =   13560
      Picture         =   "frmDeposit.frx":33A2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   23
      Top             =   7920
      Width           =   435
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   22
      Top             =   7920
      Width           =   210
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   7800
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   9000
      Width           =   615
   End
   Begin VB.Shape Shape3 
      Height          =   1695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   4935
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   15
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   14
      Top             =   240
      Width           =   210
   End
   Begin VB.Label Label8 
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14400
      TabIndex        =   3
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   2
      Top             =   8760
      Width           =   615
   End
End
Attribute VB_Name = "frmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RS As ADODB.RecordSet


Private Sub cmddpst_Click(Index As Integer)
frmDpissu.Show vbModal
End Sub



Private Sub cmdOperations_Click(Index As Integer)

If RS.State = adStateOpen Then
RS.Close
End If
If txtfrom.Text = "" Or txtTo.Text = "" Then
MsgBox "Please Enter valid dates ", vbCritical, "Attendance"
Exit Sub
Else
If Not IsDate(txtfrom.Text) Then
MsgBox "Please enter a valid date ", vbCritical, "Attendance"
Exit Sub
Else
If Not IsDate(txtTo.Text) Then
MsgBox "Please enter a valid date in it", vbCritical, "Attendance"
Exit Sub
Else
RS.Open "select * from tblDepo where Date between #" & CDate(txtfrom.Text) & "# and #" & CDate(txtTo.Text) & "# ", CN
If RS.RecordCount > 0 Then
RS.MoveFirst
Set DataGrid1.DataSource = RS
Else
MsgBox "Could not found fine between" & "  " & txtfrom.Text & " and " & txtTo.Text, vbCritical, "Attendance"
End If
End If
End If
End If
End Sub



Private Sub cmdRefresh_Click()
Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
RS.Open "SELECT * FROM tblDepo  ", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
    'DisplayRecords
End Sub

Private Sub cmdReport_Click(Index As Integer)

 If Index = 0 Then cmdRefresh_Click
    Set drDeposit.DataSource = RS
    drDeposit.Show

End Sub



Private Sub cmdRtrn_Click(Index As Integer)
frmDprtrn.Show vbModal

End Sub

Private Sub Form_Load()
'Loads a form and initializes all variables
    On Error GoTo hell
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblDepo", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
    'DisplayRecords
    With frmMain.ImgList32
        cmdReport(0).Picture = .ListImages(6).Picture
        cmdReport(1).Picture = .ListImages(6).Picture
    End With

Exit Sub

hell:
    Handler Err
    Resume Next
End Sub


Private Sub Form_Resize()
'Resizes a form according to the screen size, resolution or form resize
    On Error Resume Next
        SSTab1.Height = Me.Height - 2500
        SSTab1.Width = Me.Width - 400

       

        DataGrid1.Width = SSTab1.Width - 280
        DataGrid1.Height = SSTab1.Height - 580
       
       

End Sub




