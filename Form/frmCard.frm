VERSION 5.00
Begin VB.Form frmCard 
   Caption         =   "Issue Cards"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdId 
      Caption         =   "ISSUE IDENTITY CARD"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   4575
   End
   Begin VB.CommandButton cmdLib 
      Caption         =   "ISSUE LIBRARY CARD"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdCode 
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Browse record."
      Top             =   1200
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCard.frx":0000
      Height          =   855
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   4920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Student Code:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCode_Click()


    With frmSelectDg
        .CommandText = "Select * From tblMembers"
        .DataGrid1.Caption = "Members Table"
        .Show vbModal
        If .OKPressed Then
            txtCode.Text = .rRS1
            
        End If
    End With

End Sub


Private Sub cmdId_Click()
Dim RS As ADODB.RecordSet
On Error GoTo hell
    CN.BeginTrans
    Set RS = New ADODB.RecordSet
    With RS
     .Open "Select [Id card] from tblMembers where [ID]='" & txtCode.Text & "'", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields(0) = True
        .Update
        .Close
        Set RS = Nothing
    End With
    CN.CommitTrans
        
    If MsgBox("The Identity card has been issued to " & txtCode.Text & vbNewLine & "Do you want to create a new issue instance?", vbInformation + vbYesNo) = vbYes Then
        txtCode.Text = " "
        
        
    Else
        Unload Me
    End If
    Exit Sub
        
hell:
    Handler Err
    CN.RollbackTrans
 
End Sub

Private Sub cmdLib_Click()
Dim RS As ADODB.RecordSet
On Error GoTo hell
    CN.BeginTrans
    Set RS = New ADODB.RecordSet
    With RS
     .Open "Select [Library card] from tblMembers where [ID]='" & txtCode.Text & "'", CN, adOpenDynamic, adLockOptimistic
        .MoveFirst
        .Fields(0) = True
        .Update
        .Close
        Set RS = Nothing
    End With
    CN.CommitTrans
        
    If MsgBox("The Library card has been issued to " & txtCode.Text & vbNewLine & "Do you want to create a new issue instance?", vbInformation + vbYesNo) = vbYes Then
        txtCode.Text = " "
        
        
    Else
        Unload Me
    End If
    Exit Sub
        
hell:
    Handler Err
    CN.RollbackTrans
 
End Sub

Private Sub Form_Load()

    With frmMain
        cmdCode.Picture = .ImgList16.ListImages(1).Picture
        Me.Icon = .ImgList32.ListImages(7).Picture
    End With
       Image1.Picture = Me.Icon
End Sub
