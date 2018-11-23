VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBooks 
   Caption         =   "Books Master File"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   10350
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12938
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Grid View"
      TabPicture(0)   =   "frmBooks.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Form View"
      TabPicture(1)   =   "frmBooks.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Reports"
      TabPicture(2)   =   "frmBooks.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdReport2(1)"
      Tab(2).Control(1)=   "cmdReport(0)"
      Tab(2).Control(2)=   "Label22"
      Tab(2).Control(3)=   "Label21"
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(5)=   "Label19"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdReport2 
         Height          =   615
         Index           =   1
         Left            =   -74640
         Picture         =   "frmBooks.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdReport 
         Height          =   615
         Index           =   0
         Left            =   -74640
         Picture         =   "frmBooks.frx":0D1E
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Book Information"
         Height          =   6615
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   9615
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   11
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   4440
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   10
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   4080
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   9
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   3720
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   8
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   3360
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   3000
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   2640
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   2280
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1920
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1560
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1200
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   840
            Width           =   6495
         End
         Begin VB.TextBox txtDisp 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   480
            Width           =   6495
         End
         Begin VB.Label Label27 
            Caption         =   "Source,Bill No :"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label Label26 
            Caption         =   "Price:"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "pages:"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Year:"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Category:"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "Publisher :"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Vol no:"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Ed:"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Author:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "Title:"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Date:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Accession No:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6900
         Left            =   -74880
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
         Caption         =   "Book Details"
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
         Caption         =   "Create a complete book report containing Accession NO,Category,Year,Pages,Price,Source,Bill no. etc."
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
         TabIndex        =   47
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label Label21 
         Caption         =   "Create Book Report 2"
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
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label Label20 
         Caption         =   "Create a complete book report containing Accesion NO,Date,Title,Author,Ed,Vol no,Publisher Name."
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
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label19 
         Caption         =   "Create Book Report 1"
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
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   3480
      ScaleHeight     =   975
      ScaleWidth      =   6855
      TabIndex        =   8
      Top             =   7920
      Width           =   6855
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   2
         Left            =   3120
         Picture         =   "frmBooks.frx":19E8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   1
         Left            =   2400
         Picture         =   "frmBooks.frx":22B2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdOperations 
         Height          =   615
         Index           =   0
         Left            =   1680
         Picture         =   "frmBooks.frx":2F7C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   1
         Left            =   240
         Picture         =   "frmBooks.frx":3846
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   3840
         Picture         =   "frmBooks.frx":4510
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdRefresh 
         Height          =   615
         Left            =   4560
         Picture         =   "frmBooks.frx":51DA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   5280
         Picture         =   "frmBooks.frx":5EA4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAMod 
         Height          =   615
         Index           =   0
         Left            =   960
         Picture         =   "frmBooks.frx":6B6E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   720
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
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraNavigation 
      Height          =   800
      Left            =   12000
      TabIndex        =   0
      Top             =   7920
      Width           =   1725
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   3
         Left            =   1290
         Picture         =   "frmBooks.frx":7838
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   2
         Left            =   870
         Picture         =   "frmBooks.frx":7BC2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   1
         Left            =   450
         Picture         =   "frmBooks.frx":7F4C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   400
      End
      Begin VB.CommandButton cmdNavigate 
         Height          =   265
         Index           =   0
         Left            =   30
         Picture         =   "frmBooks.frx":82D6
         Style           =   1  'Graphical
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   480
         Width           =   700
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Information of all the books in the library are stored in this table."
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
      TabIndex        =   26
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Book Details"
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
      TabIndex        =   25
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   15000
      Y1              =   7800
      Y2              =   7800
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
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Book Table Master Form

Private RS As ADODB.RecordSet


Private Sub cmdReport2_Click(Index As Integer)
Set drBookList2.DataSource = RS
    drBookList2.Show
  
End Sub

Private Sub Form_Load()

'Create recordset and refresh. Link Report icons to ImageList

    On Error GoTo ero1
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblBooks", CN, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = RS
   
    
    
    DisplayRecords
    With frmMain.ImgList32
        cmdReport(0).Picture = .ListImages(6).Picture
        cmdReport2(1).Picture = .ListImages(6).Picture
    End With

Exit Sub

ero1:
    Handler Err
    Resume Next

End Sub

Private Sub Form_Resize()


    On Error Resume Next
        SSTab1.Height = Me.Height - 2500
        SSTab1.Width = Me.Width - 400

        Line2.X1 = SSTab1.Left
        Line2.X2 = SSTab1.Left + SSTab1.Width
        Line2.Y1 = SSTab1.Top + SSTab1.Height + 400
        Line2.Y2 = Line2.Y1
        Line2.ZOrder vbBringToFront

        DataGrid1.Width = SSTab1.Width - 280
        DataGrid1.Height = SSTab1.Height - 580
        Frame1.Height = DataGrid1.Height - 100
        Frame1.Width = DataGrid1.Width - 200

        'Line3.X1 = Frame1.Left
        'Line3.X2 = Frame1.Width - Frame1.Left - 180
        'Line3.Y1 = txtDisp(6).Height + txtDisp(6).Top + 1000
        'Line3.Y2 = Line3.Y1

        'LineMove Line4, Line3
        'LineMove Line1, Line2

        pic.Top = Line1.Y1 + 200
        Label9.Top = pic.Top
        Label11.Top = Label9.Top + Label9.Height
        Image1.Top = pic.Top
        fraNavigation.Top = pic.Top
        fraNavigation.Left = Line1.X2 - fraNavigation.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)


    Set RS = Nothing
    Set frmBooks = Nothing

End Sub

Private Sub cmdOperations_Click(Index As Integer)



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

'Create dynamic reports

    'If Index = 0 Then cmdRefresh_Click
    Set drBookList.DataSource = RS
    drBookList.Show
    
    
    


End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

'Allow keyboard navigation to display recordnumber

    If KeyCode = 38 Or KeyCode = 40 Then DisplayRecords

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

'Allow mouse navigation to display recordnumber

    DisplayRecords

End Sub

Private Sub DisplayRecords()

'Display the current and total number of record

Dim i As Integer

    On Error Resume Next
        With RS
            If .RecordCount < 1 Then
                txtcount.Text = 0
            Else
                txtcount.Text = .AbsolutePosition
            End If
            lblmax.Caption = .RecordCount

            For i = 0 To 11
                txtDisp(i).Text = .Fields(i)
            Next i
        End With
        txtDisp(10).Text = FormatCurrency$(txtDisp(5).Text)

End Sub

Private Sub cmdDelete_Click()

'Deletes a record, undeletable if a book is borrowed.

Dim ans As Integer, pos As Integer

    On Error GoTo hell
    With RS
        'Check if there is no record
        If .RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
        'Check whether book is borrowed
        If .Fields("Borrowed") = True Then MsgBox "You cannot delete this book record because it is borrowed by someone" & vbNewLine & "The book must be returned to the library before its record can be deleted.", vbInformation, "Book Borrowed"
        'Confirm deletion of record
        ans = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Deletion")
        Screen.MousePointer = vbHourglass
        If ans = vbYes Then
            'Delete the record
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
    On Error Resume Next
        Handler Err
        CN.RollbackTrans

End Sub

Private Sub cmdNavigate_Click(Index As Integer)

'Navigate a recordset through command buttons

    Navigate Index, RS
    DisplayRecords

End Sub

Private Sub cmdRefresh_Click()

'Refresh the recordset

    With RS
        .Filter = adFilterNone
        .Requery
    End With
    DisplayRecords

End Sub

Private Sub cmdClose_Click()

'Close the form

    Unload Me

End Sub

Private Sub cmdAMod_Click(Index As Integer)

'Open the add/edit form. Display current record values in form if modifying.

    On Error Resume Next
        With frmBooksAE
            .AddState = Index
            .OldID = RS.Fields(0)
            If Index = 0 Then
                .msdID.Text = RS.Fields(0)
                .txtTitle.Text = RS.Fields(2)
                .txtAuthor.Text = RS.Fields(3)
                .txtPublisher.Text = RS.Fields(6)
                .cmbCategory.Text = RS.Fields(7)
                .txtPrice.Text = RS.Fields(10)
                .txtbil.Text = RS.Fields(11)
                .txted.Text = RS.Fields(4)
                .txtvol.Text = RS.Fields(5)
                .txtyr.Text = RS.Fields(8)
                .txtpage.Text = RS.Fields(9)
                .txtdate.Text = RS.Fields(1)
                
                
                
                
                
                
            End If
            .Show vbModal
        End With
        cmdRefresh_Click
        DisplayRecords

End Sub


