VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMembersAE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Staff"
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
      Left            =   3120
      TabIndex        =   24
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox cmbCat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPeoplesAE.frx":0000
      Left            =   1680
      List            =   "frmPeoplesAE.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1920
      Width           =   3855
   End
   Begin VB.CommandButton cmdPicShow 
      Height          =   375
      Left            =   4440
      Picture         =   "frmPeoplesAE.frx":0028
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "View Picture"
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton cmdPicInsert 
      Height          =   375
      Left            =   4440
      Picture         =   "frmPeoplesAE.frx":05B2
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Browse for a picture to store in the database..."
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdPicSave 
      Height          =   375
      Left            =   1680
      Picture         =   "frmPeoplesAE.frx":0B3C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Save Picture in a file..."
      Top             =   5880
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".jpg"
      DialogTitle     =   "Save File from database as..."
      Filter          =   "Picture Files (*.jpg,*.bmp,*.wmf,*.emf)|*.jpg;*.bmp;*.wmf;*.emf|All files (*.*)|*.*"
   End
   Begin VB.TextBox txtId 
      BackColor       =   &H80000004&
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtRoll 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
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
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox cmbClass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPeoplesAE.frx":0EC6
      Left            =   1680
      List            =   "frmPeoplesAE.frx":0EDF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2760
      Width           =   3855
   End
   Begin VB.ComboBox cmbSection 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPeoplesAE.frx":0EFF
      Left            =   1680
      List            =   "frmPeoplesAE.frx":0F0F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtTik 
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
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtName 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton cmdAddSave 
      Caption         =   "Update"
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
      Left            =   360
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   6720
      Width           =   1095
   End
   Begin CLAM.Photo Photo1 
      Height          =   2055
      Left            =   1680
      TabIndex        =   21
      Top             =   4200
      Width           =   3135
      _extentx        =   5530
      _extenty        =   3625
   End
   Begin VB.Label Label4 
      Caption         =   "Category:"
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
      TabIndex        =   22
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   17
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Label10 
      Caption         =   "Roll:"
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
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Branch:"
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
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Sem:"
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
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Picture:"
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
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Ticket No:"
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
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
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
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
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
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5640
      Y1              =   960
      Y2              =   960
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
      Caption         =   $"frmPeoplesAE.frx":0F43
      Height          =   855
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5640
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5640
      Y1              =   6480
      Y2              =   6480
   End
End
Attribute VB_Name = "frmMembersAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Members Add/Edit Form
'With this form it is possible to add/modify the members table.
'
'AddState is the bool value that decides whether the form will update
'or modify a record. When AddState = True then Add
'                    When AddState = False then Modify

Private RS As ADODB.RecordSet
Public OldID As String, AddState As Boolean

Private Sub Command1_Click()
txtRoll.Text = "00"
End Sub

Private Sub Form_Load()

    On Error GoTo Err
    Set RS = New ADODB.RecordSet
    If AddState Then
        Image1.Picture = frmMembers.cmdAMod(1).Picture
        RS.Open "SELECT * FROM tblMembers", CN, adOpenStatic, adLockOptimistic
        Me.Caption = "Add Record"
    Else 'NOT AddState...
        Image1.Picture = frmMembers.cmdAMod(0).Picture
        Me.Caption = "Modify Record"
        cmdAddSave.Caption = "Save"
        RS.Open "SELECT * FROM tblMembers WHERE [ID] = '" & OldID & "'", CN, adOpenStatic, adLockOptimistic
        If Len(RS!Picture) > 0 Then
            Photo1.LoadPhoto RS!Picture
        End If
    End If

Exit Sub

Err:
    If Err.Number = 94 Or Err.Number = 3265 Then
        Resume Next 'If a null value is encountered
    Else
        Handler Err 'Unexpected error
    End If

End Sub

Private Sub cmdAddSave_Click()

'Add or save data in the recordset according to AddState

    On Error GoTo e1

    If txtId.Text = "" Then txtId.SetFocus: Exit Sub
    If txtName.Text = "" Then txtName.SetFocus: Exit Sub
    'If cmbClass.Text = "" Then cmbClass.SetFocus: Exit Sub
    'If cmbSection.Text = "" Then cmbSection.SetFocus: Exit Sub
    'If txtRoll.Text = "" Then txtRoll.SetFocus: Exit Sub
    If cmbCat.Text = " " Then cmbCat.SetFocus: Exit Sub
    
    If IsNumeric(txtRoll.Text) <> True Then MsgBox "Roll Numbers must be numeric and between 1 and 99", vbExclamation, "Type Mismatch": HighLight txtRoll: Exit Sub
    txtRoll.Text = Int(txtRoll.Text)
    If txtRoll.Text < -1 Or txtRoll.Text > 99 Then MsgBox "Roll numbers must be between 1 and 99", vbExclamation, "Type Mismatch": HighLight txtRoll: Exit Sub
    

    If AddState Then
        If RecordExists("tblMembers", "ID", txtId.Text, txtId) = True Then Exit Sub
    Else 'NOT AddState...
        If txtId.Text <> OldID Then
            If RecordExists("tblMembers", "ID", txtId.Text, txtId) = True Then Exit Sub
          
        End If
    End If

    CN.BeginTrans
    With RS
        If AddState Then RS.AddNew
        .Fields(0) = txtId.Text
        .Fields(1) = txtName.Text
        .Fields(2) = cmbClass.Text
        .Fields(3) = cmbSection.Text
        .Fields(4) = txtRoll.Text
        .Fields(5) = txtTik.Text
        '.Fields(6) = txtRoll.Text
        .Fields(9) = cmbCat.Text
        
        
        Photo1.SavePhoto .Fields("Picture")
        RS.Update
    End With
    CN.CommitTrans

    If AddState Then
        FindRecord RS, RS.Fields(0).Name, True, txtId.Text, 0
        MsgBox "New record has been successfully added", vbInformation

        If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            cmdReset_Click
        Else
            Unload Me
        End If

    Else 'NOT AddState...
        FindRecord RS, RS.Fields(0).Name, True, txtId.Text, 0

        MsgBox "Changes in record has been successfully saved", vbInformation
        Unload Me
    End If

Exit Sub

e1:
    On Error Resume Next
        'CN.RollbackTrans
        
        Handler Err

End Sub





Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdPicInsert_Click()

'Open photo from disk

    Photo1.OpenPhotoFile

End Sub

Private Sub cmdPicSave_Click()

'Save photo to disk

    On Error GoTo hell
    cmdlg.ShowSave
    If cmdlg.Filename <> "" Then
        SavePicture Photo1.Picture, cmdlg.Filename
    End If
hell:

End Sub

Private Sub cmdPicShow_Click()

'Open photo from a temp file

    On Error Resume Next
        Kill "tmp.jpg"
        SavePicture Photo1.Picture, "tmp.jpg"
        ShellEx "tmp.jpg"

End Sub

Private Sub cmdReset_Click()

    txtId.Text = ""
    txtName.Text = ""
    txtRoll.Text = ""
    txtTik.Text = ""

    cmbClass.ListIndex = 0
    cmbSection.ListIndex = 0
    cmbCat.ListIndex = 0
    

End Sub



