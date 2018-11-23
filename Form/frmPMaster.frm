VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMembers 
   Caption         =   "Member Details Master Form"
   ClientHeight    =   9465
   ClientLeft      =   1545
   ClientTop       =   1170
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraNavigation 
      Height          =   800
      Left            =   12000
      TabIndex        =   17
      Top             =   7920
      Width           =   1725
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
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   0
         Left            =   30
         Picture         =   "frmPMaster.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   1
         Left            =   450
         Picture         =   "frmPMaster.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   2
         Left            =   870
         Picture         =   "frmPMaster.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   3
         Left            =   1290
         Picture         =   "frmPMaster.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   400
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
         TabIndex        =   24
         Top             =   480
         Width           =   700
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
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3480
      ScaleHeight     =   975
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   7920
      Width           =   6855
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   0
         Left            =   960
         Picture         =   "frmPMaster.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Height          =   615
         Left            =   5280
         Picture         =   "frmPMaster.frx":1AF2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   4560
         Picture         =   "frmPMaster.frx":27BC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   3840
         Picture         =   "frmPMaster.frx":3486
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAMod 
         Default         =   -1  'True
         Height          =   615
         Index           =   1
         Left            =   240
         Picture         =   "frmPMaster.frx":4150
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   0
         Left            =   1680
         Picture         =   "frmPMaster.frx":4E1A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   1
         Left            =   2400
         Picture         =   "frmPMaster.frx":56E4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   2
         Left            =   3120
         Picture         =   "frmPMaster.frx":63AE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
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
         Left            =   5280
         TabIndex        =   16
         Top             =   720
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
         Left            =   4560
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
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
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modify"
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
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Add New"
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
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
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
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   720
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
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   720
         Width           =   615
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
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Grid View"
      TabPicture(0)   =   "frmPMaster.frx":6C78
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Form View"
      TabPicture(1)   =   "frmPMaster.frx":6C94
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Reports"
      TabPicture(2)   =   "frmPMaster.frx":6CB0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label22"
      Tab(2).Control(1)=   "Label21"
      Tab(2).Control(2)=   "Label20"
      Tab(2).Control(3)=   "Label23"
      Tab(2).Control(4)=   "cmdReport(1)"
      Tab(2).Control(5)=   "cmdReport(0)"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   0
         Left            =   -74640
         Picture         =   "frmPMaster.frx":6CCC
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   1
         Left            =   -74640
         Picture         =   "frmPMaster.frx":7996
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1800
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Member Details"
         Height          =   6375
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   11775
         Begin CLAM.Photo picBox 
            Height          =   4335
            Left            =   4440
            TabIndex        =   48
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7646
         End
         Begin VB.CommandButton cmdRetrive 
            Caption         =   "&Retrive"
            Height          =   375
            Left            =   5040
            TabIndex        =   47
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label19 
            Caption         =   "Picture:"
            Height          =   255
            Left            =   4080
            TabIndex        =   40
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "Roll:"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Branch:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Semester:"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "First Name:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Student Code:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   1215
         End
         Begin VB.Line lnBorder 
            BorderColor     =   &H80000014&
            Index           =   1
            X1              =   120
            X2              =   9600
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line lnBorder 
            BorderColor     =   &H80000010&
            BorderWidth     =   2
            Index           =   0
            X1              =   120
            X2              =   9600
            Y1              =   2040
            Y2              =   2040
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6900
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   9915
         _ExtentX        =   17489
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
         TabIndex        =   46
         Top             =   600
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
         TabIndex        =   45
         Top             =   840
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
         TabIndex        =   44
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label22 
         Caption         =   $"frmPMaster.frx":8660
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
         TabIndex        =   43
         Top             =   2040
         Width           =   4095
      End
   End
   Begin VB.Label Label24 
      Caption         =   "  Card "
      Height          =   255
      Left            =   9600
      TabIndex        =   49
      Top             =   8640
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   15000
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   480
   End
   Begin VB.Label Label9 
      Caption         =   "Member Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Information of all the members of the Library are stored here."
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
      Height          =   615
      Left            =   720
      TabIndex        =   25
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   15000
      Y1              =   7800
      Y2              =   7800
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Members Master Form
'this form manupulates the Members Table, except the Adding/Editing which
'are done by Members Add/Edit Form

Private RS As ADODB.RecordSet

Private Sub cmdOperations_Click(Index As Integer)
'Shows the Search/Sort/Filter form by creating a new instance and destroying it once done
Dim obj As Form

    If Index = 0 Then Set obj = frmSearch
    If Index = 1 Then Set obj = frmFilter
    If Index = 2 Then Set obj = frmSort

    With obj
        Set .SourceRs = RS
        .Show vbModal
    End With
    Set obj = Nothing

End Sub

Private Sub cmdReport_Click(Index As Integer)
'Creates dynamic reports
    If Index = 0 Then cmdRefresh_Click
    'Set drBookList.DataSource = RS
       Set drMembers.DataSource = RS
       
    drMembers.Show

End Sub

Private Sub cmdRetrive_Click()
'Retrives a picture to be shown in the Photo box with the use of Photo Access User Control
Dim tmpRS As New ADODB.RecordSet

    With tmpRS
        .Open "SELECT [Picture] FROM tblMembers WHERE [ID]='" & txtDisp(0).Text & "'", CN, adOpenForwardOnly, adLockOptimistic
        If Len(RS!Picture) > 0 Then
            picBox.LoadPhoto RS!Picture
        Else
            Set picBox.Picture = LoadPicture()
        End If
        .Close
    End With
    Set tmpRS = Nothing

End Sub

Private Sub cmdTiket_Click()
frmCard.Show

End Sub

Private Sub Form_Load()
'Loads a form and initializes all variables
    On Error GoTo err1
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblMembers", CN, adOpenDynamic, adLockOptimistic
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

Private Sub Form_Resize()
'Resizes a form according to the screen size, resolution or form resize
    On Error Resume Next
        SSTab1.Height = Me.Height - 2500
        SSTab1.Width = Me.Width - 400

        Line1.X1 = SSTab1.Left
        Line1.X2 = SSTab1.Left + SSTab1.Width
        Line1.Y1 = SSTab1.Top + SSTab1.Height + 400
        Line1.Y2 = Line1.Y1

        DataGrid1.Width = SSTab1.Width - 280
        DataGrid1.Height = SSTab1.Height - 580
        Frame1.Height = DataGrid1.Height - 100
        Frame1.Width = DataGrid1.Width - 200

        lnBorder(0).X1 = Frame1.Left
        lnBorder(0).X2 = Frame1.Width - Frame1.Left - 180
        lnBorder(0).Y1 = txtDisp(3).Height + txtDisp(3).Top + 180
        lnBorder(0).Y2 = lnBorder(0).Y1

        lnBorder(2).X1 = lnBorder(1).X1
        lnBorder(2).X2 = lnBorder(1).X2
        lnBorder(2).Y1 = txtDisp(6).Height + txtDisp(6).Top + 180
        lnBorder(2).Y2 = lnBorder(2).Y1

        LineMove Line2, Line1
        LineMove lnBorder(1), lnBorder(0)
        LineMove lnBorder(3), lnBorder(2)

        picBox.Left = txtDisp(0).Left + txtDisp(0).Width + 200
        picBox.Top = txtDisp(0).Top
        picBox.Height = Frame1.Height - cmdRetrive.Height - Frame1.Top
        picBox.Width = Frame1.Width - picBox.Left - picBox.Width + txtDisp(0).Width

        cmdRetrive.Left = picBox.Left + picBox.Width - cmdRetrive.Width

        pic.Top = Line1.Y1 + 200
        Label9.Top = pic.Top
        Label11.Top = Label9.Top + Label9.Height
        Label19.Left = picBox.Left
        Label19.Top = picBox.Top - Label19.Height
        Image1.Top = pic.Top
        fraNavigation.Top = pic.Top
        fraNavigation.Left = Line1.X2 - fraNavigation.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Destroys variables to free memory
    Set RS = Nothing
    Set frmMembers = Nothing

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38 Or KeyCode = 40 Then DisplayRecords

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

            For i = 0 To 6
                txtDisp(i).Text = .Fields(i)
            Next i

        End With

End Sub

Private Sub cmdDelete_Click()
'Deletes a record
    On Error GoTo hell
    With RS
        '-Check if there is no record
        If .RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
        '-Confirm deletion of record

Dim ans As Integer, pos As Integer
        ans = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Deletion")
        Screen.MousePointer = vbHourglass
        If ans = vbYes Then
            '-Delete the record
            pos = .AbsolutePosition
            CN.BeginTrans
            .Delete
            .Requery
            CN.CommitTrans
            If pos > .RecordCount Then
                If Not .EOF Or .BOF Then .MoveFirst
            Else
                .AbsolutePosition = pos
            End If
            MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
        End If
        Screen.MousePointer = vbDefault
    End With

Exit Sub

hell:
    Handler Err
    CN.RollbackTrans

End Sub

Private Sub cmdNavigate_Click(Index As Integer)
'
    Navigate Index, RS
    DisplayRecords

End Sub

Private Sub cmdRefresh_Click()

    With RS
        .Filter = adFilterNone
        .Requery
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdAMod_Click(Index As Integer)

    On Error Resume Next
        With frmMembersAE
            .AddState = Index
            .OldID = RS.Fields(0)
            If Index = 0 Then
                .txtId.Text = RS(0)
                .txtName.Text = RS(1)
                .cmbClass.Text = RS(2)
                .cmbSection.Text = RS(3)
                .txtRoll.Text = RS(4)
                '.cmbSection = RS(5)
                .txtTik = RS(5)
                .cmbCat.Text = RS(9)
                
            End If
            .Show vbModal
        End With

        cmdRefresh_Click
        DisplayRecords

End Sub

