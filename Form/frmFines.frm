VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFines 
   Caption         =   "Fines "
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9540
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   8160
      Picture         =   "frmFines.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8520
      Width           =   615
   End
   Begin VB.CommandButton cmdOperations 
      Height          =   615
      Index           =   0
      Left            =   6600
      Picture         =   "frmFines.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox txtfrom 
      Height          =   285
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   9
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   10
      Top             =   8040
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Grid View"
      TabPicture(0)   =   "frmFines.frx":1594
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reports"
      TabPicture(1)   =   "frmFines.frx":15B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdReport(1)"
      Tab(1).Control(1)=   "cmdReport(0)"
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(3)=   "Label21"
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(5)=   "Label23"
      Tab(1).ControlCount=   6
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4860
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   8573
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
         Caption         =   "Member Details"
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
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   1
         Left            =   -74640
         Picture         =   "frmFines.frx":15CC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   0
         Left            =   -74640
         Picture         =   "frmFines.frx":2296
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   $"frmFines.frx":2F60
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
         Left            =   -73800
         TabIndex        =   7
         Top             =   2160
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
         Left            =   -73800
         TabIndex        =   6
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label Label20 
         Caption         =   "Create a complete report on all the books that are in the library. The Grid View will show the complete inventory."
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
         Left            =   -73800
         TabIndex        =   5
         Top             =   960
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
         Left            =   -73800
         TabIndex        =   4
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
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
      Left            =   5760
      TabIndex        =   13
      Top             =   8040
      Width           =   435
   End
   Begin VB.Label Label2 
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
      Left            =   7680
      TabIndex        =   8
      Top             =   8040
      Width           =   210
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   7920
      Width           =   4695
   End
End
Attribute VB_Name = "frmFines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.RecordSet

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
RS.Open "select * from tblTrans where Date_Borrowed between #" & CDate(txtfrom.Text) & "# and #" & CDate(txtTo.Text) & "# and fines>0", CN
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
RS.Open "SELECT * FROM tblTrans where fines>0 ", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
    DisplayRecords
End Sub

Private Sub cmdReport_Click(Index As Integer)
 If Index = 0 Then cmdRefresh_Click
    Set drFine.DataSource = RS
    drFine.Show

End Sub


Private Sub Form_Load()
On Error GoTo err1
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblTrans where fines>0 ", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
    DisplayRecords
    With frmMain.ImgList32
        cmdReport(0).Picture = .ListImages(6).Picture
        cmdReport(1).Picture = .ListImages(6).Picture
    End With

Exit Sub

err1:
    Handler Err
    Resume Next

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    DisplayRecords

End Sub

Private Sub DisplayRecords()

'-Display the current and total number of record

Dim i As Integer

    On Error Resume Next

        With RS
            If .RecordCount < 1 Then
                txtcount.Text = 0
            Else
                txtcount.Text = .AbsolutePosition
            End If
            lblmax.Caption = .RecordCount

            'For i = 0 To 6
           '     txtDisp(i).Text = .Fields(i)
            'Next i '

        End With

End Sub
