VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBooksAE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbil 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   3360
      Width           =   3855
   End
   Begin VB.TextBox txtdate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   28
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox txtyr 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   26
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox txtpage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtvol 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   22
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox txted 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   20
      Top             =   1200
      Width           =   3615
   End
   Begin MSMask.MaskEdBox txtPrice 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   3000
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmbCategory 
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
      ItemData        =   "frmBooksAE.frx":0000
      Left            =   2280
      List            =   "frmBooksAE.frx":001F
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   3855
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
      Left            =   7560
      TabIndex        =   8
      Top             =   3840
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
      Left            =   8880
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
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
      Left            =   10200
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtAuthor 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtPublisher 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
   End
   Begin MSMask.MaskEdBox msdID 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   11280
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label14 
      Caption         =   "Date"
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
      Left            =   6720
      TabIndex        =   27
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Year"
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
      Left            =   6720
      TabIndex        =   25
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Pages"
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
      Left            =   6720
      TabIndex        =   23
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Vol NO"
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
      Left            =   6720
      TabIndex        =   21
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "ED:"
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
      Left            =   6720
      TabIndex        =   19
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Source,Bill No & date:"
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
      TabIndex        =   18
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Price:"
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
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBooksAE.frx":008E
      Height          =   855
      Left            =   6960
      TabIndex        =   16
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   480
      Stretch         =   -1  'True
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Accesion Number:"
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
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Title:"
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
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Author:"
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
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Publisher Name:"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
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
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Top             =   1200
      Width           =   105
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
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   11280
      Y1              =   975
      Y2              =   960
   End
End
Attribute VB_Name = "frmBooksAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AddState As Boolean, OldID As String
Private RS As ADODB.RecordSet

Private Sub Form_Load()


    On Error GoTo Err
    Set RS = New ADODB.RecordSet
    If AddState Then
        Image1.Picture = frmBooks.cmdAMod(1).Picture
        RS.Open "SELECT * FROM tblBooks", CN, adOpenStatic, adLockOptimistic
        Me.Caption = "Add Record"
    Else
        Image1.Picture = frmBooks.cmdAMod(0).Picture
        Me.Caption = "Modify Record"
        cmdAddSave.Caption = "Save"
        RS.Open "SELECT * FROM tblBooks WHERE [Accesion Number] =  '" & OldID & "'", CN, adOpenStatic, adLockOptimistic
        
        
        
    End If

Exit Sub

Err:
  If Err.Number = 94 Or Err.Number = 3265 Then
      Resume Next
 Else
        Handler Err
 End If

End Sub

Private Sub cmdAddSave_Click()



    On Error GoTo err1
    
    'my addition
  If txtdate.Text = "" Then txtdate.Text = "00/00/0000"
       ' If txtTitle.Text = "" Then .Fields(2) = " " Else .Fields(2) = txtTitle.Text
        If txtAuthor.Text = "" Then txtAuthor.Text = "N/A"
        If txted.Text = "" Then txted.Text = "N/A"
        If txtvol.Text = "" Then txtvol.Text = "N/A"
        If txtPublisher.Text = "" Then txtPublisher.Text = "N/A"
        If txtyr.Text = "" Then txtyr.Text = "N/A"
        If txtpage.Text = "" Then txtpage.Text = "N/A"
        If txtbil.Text = "" Then txtbil.Text = "N/A"
        If txtPrice.Text = "" Then txtPrice.Text = "0"
        If cmbCategory.ListIndex = 0 Then cmbCategory.Text = "N/A"
        
        
    'end
    
    
    If msdID.Text = "" Then msdID.SetFocus: Exit Sub
    If txtTitle.Text = "" Then txtTitle.SetFocus: Exit Sub
    'If Len(msdID.Text) <> 10 Then MsgBox "All Book ID must be 10 charecters long", vbExclamation: HighLight msdID: Exit Sub
     msdID.Text = UCase$(msdID.Text)
    'If IsNumeric(Right$(msdID.Text, 9)) = False Then MsgBox "Book ID must start with B followed by 9 digits", vbExclamation: HighLight msdID: Exit Sub

    If AddState Then
        If RecordExists("tblBooks", "Accesion Number", msdID.Text, msdID) = True Then Exit Sub
    Else
        If msdID.Text <> OldID Then
            If RecordExists("tblBooks", "Accesion Number", msdID.Text, msdID) = True Then Exit Sub
        End If
    End If

    CN.BeginTrans
    With RS
        If AddState = True Then RS.AddNew
        .Fields(0) = msdID.Text
        .Fields(1) = txtdate.Text
        .Fields(2) = txtTitle.Text
        .Fields(3) = txtAuthor.Text
        .Fields(4) = txted.Text
        .Fields(5) = txtvol.Text
        .Fields(6) = txtPublisher.Text
        .Fields(7) = cmbCategory.Text
        .Fields(8) = txtyr.Text
        .Fields(9) = txtpage.Text
        .Fields(10) = CCur(txtPrice.Text)
        .Fields(11) = txtbil.Text
        
        If txtdate.Text = "" Then .Fields(1) = "N/A" Else .Fields(1) = txtdate.Text
        If txtTitle.Text = "" Then .Fields(2) = " " Else .Fields(2) = txtTitle.Text
        If txtAuthor.Text = "" Then .Fields(3) = "N/A" Else .Fields(3) = txtAuthor.Text
        If txted.Text = "" Then .Fields(4) = "N/A" Else .Fields(4) = txted.Text
        If txtvol.Text = "" Then .Fields(5) = "N/A" Else .Fields(5) = txtvol.Text
        If txtPublisher.Text = "" Then .Fields(6) = "N/A" Else .Fields(6) = txtPublisher.Text
        If txtyr.Text = "" Then .Fields(8) = "N/A" Else .Fields(8) = txtyr.Text
        If txtpage.Text = "" Then .Fields(9) = "N/A" Else .Fields(9) = txtpage.Text
        If txtbil.Text = "" Then .Fields(11) = "N/A" Else .Fields(11) = txtbil.Text



If txtPrice.Text = "" Then txtPrice.Text = "0"
        
        
        
        
        RS.Update
    End With
    CN.CommitTrans

    If AddState Then
        
        FindRecord RS, RS.Fields(0).Name, True, msdID.Text, 0
        MsgBox "New record has been successfully added", vbInformation

        'If MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            'Unload Me
           ' frmBooksAE.Show vbModal
            'cmdReset_Click
            
       ' Else
            'Unload Me
        'End If

    Else
        FindRecord RS, RS.Fields(0).Name, True, msdID.Text, 0

        MsgBox "Changes in record has been successfully saved", vbInformation
        Unload Me
    End If

Exit Sub

err1:
    On Error Resume Next
        Handler Err
        CN.RollbackTrans

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdReset_Click()


    msdID.Text = ""
    txtTitle.Text = ""
    txtAuthor.Text = ""
    txtPublisher.Text = ""
    txtPrice.Text = ""
    txtdate.Text = ""
    txtyr.Text = ""
    txtpage.Text = ""
    txtvol.Text = ""
    txted.Text = ""
    txtbil.Text = ""
    
    cmbCategory.ListIndex = 3

End Sub

