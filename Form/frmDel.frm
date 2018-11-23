VERSION 5.00
Begin VB.Form frmDel 
   Caption         =   "Delete Records"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Erase Deposit history from Deposit Table"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3840
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Erase Transaction history from Fines Table"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "The following settings can parmanently change tha datbase by deleting records.Once deleted data can't be re acquired."
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDel.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Caption         =   "WARNING: This will clear all Deposit records "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "CLEAR DEPOSIT HISTORY FOR ALL RETURNED BOOKS                                                       "
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "WARNING: This will clear all fine records "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "CLEAR TRANSACTION HISTORY FOR ALL RETURNED BOOKS                                                       "
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "frmDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.RecordSet


Private Sub Command1_Click()
If MsgBox("Do you Realy want to Modify Transaction table ?", vbInformation + vbYesNo) = vbYes Then
     'On Error GoTo hell
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "DELETE  * FROM tblTrans where returned=True ", CN, adOpenDynamic, adLockOptimistic
End If
'hell:
'Handler


'Resume Next

End Sub

Private Sub Command2_Click()
If MsgBox("Do you Realy want to Modify Deposit table ?", vbInformation + vbYesNo) = vbYes Then
     'On Error GoTo hell
    Set RS = New ADODB.RecordSet
    RS.CursorLocation = adUseClient
    RS.Open "DELETE  * FROM tblDepo where returned=True ", CN, adOpenDynamic, adLockOptimistic
End If
'hell:
'Resume Next

End Sub

Private Sub Command3_Click()
frmPass.Show vbModal

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Destroys variables to free memory
    Set RS = Nothing
    Set frmMembers = Nothing

End Sub
