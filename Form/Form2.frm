VERSION 5.00
Begin VB.Form frmDprtrn 
   Caption         =   "Return Deposit"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4770
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Height          =   315
      Left            =   5520
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Use calculator"
      Top             =   3240
      Width           =   315
   End
   Begin VB.TextBox Text3 
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   3735
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
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
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
      Left            =   4200
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "R&eturn Deposit"
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
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdCode 
      Height          =   315
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Browse record."
      Top             =   1320
      Width           =   315
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
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
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   3375
   End
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label8 
      Caption         =   "Accession Number:"
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
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":038A
      Height          =   735
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Handling Charges"
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
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label7 
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
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmDprtrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdCode_Click()
On Error Resume Next
        With frmSelectDg
         .CommandText = "SELECT * from tbldepo where Returned=false"
         .DataGrid1.Caption = "Deposit Table"
            .Show vbModal
            
            'display the data
            If .OKPressed Then
                Text4.Text = .rRS1
                Text3.Text = .rRS2
                Text1.Text = .rRS9
                Text2.Text = .rRS4
                
                Else
                'If the user did not enter anything then skip the second
                'part of the procedure to skip errors that may arise because
                'there will be no data (in text4 and text1) and as such
                'null errors or record not found errors.
                Exit Sub
            End If
        End With
        


End Sub

Private Sub cmdReset_Click()

    Text1.Text = ""
    Text4.Text = ""
    txtFines.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
End Sub

Private Sub cmdReturn_Click()
Dim RS As ADODB.RecordSet
      If Text4.Text = "" Then Text4.SetFocus
    On Error GoTo hell
    Set RS = New ADODB.RecordSet
    With RS
        CN.BeginTrans       'Begin a new transaction
        .Open "Select [Borrowed] from tblBooks where [Accesion Number]='" & Text3.Text & "'", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields(0) = False
        .Update
        .Close
        
        
        .Open "Select [Handling_Charge],[Returned] From tblDepo where [Accession_Number]='" & Text3.Text & "'" & "And [Returned] = False", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields("Handling_Charge") = txtFines.Text
        .Fields("Returned") = True
        .Update
        .Close
        CN.CommitTrans      'If no error was raised then record info
    End With
    Set RS = Nothing
    
    'Show MsgBox if another book needs returning
    If MsgBox("The book " & Text3.Text & " has been returned from " & Text1.Text & vbNewLine & vbNewLine & "Do you want to create a new return book instance?", vbInformation + vbYesNo) = vbYes Then
        cmdReset_Click
    Else
        Unload Me
    End If
 Exit Sub

hell:
    Handler Err

    On Error Resume Next    'If an error was raised then rollback
        CN.RollbackTrans
   
    
    
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

