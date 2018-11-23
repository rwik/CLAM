VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0E0FF&
   Caption         =   "CLAM-College Library Automation & Management"
   ClientHeight    =   7035
   ClientLeft      =   3435
   ClientTop       =   3270
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0E42
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Issue Book"
            Object.ToolTipText     =   "Issue Book"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Return Book"
            Object.ToolTipText     =   "Return Book"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Book Records"
            Object.ToolTipText     =   "Book Records"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Member Records"
            Object.ToolTipText     =   "Member Records"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Fines"
            Object.ToolTipText     =   "Fine log"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deposit Window"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Reports"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Settings"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "About"
            Object.ToolTipText     =   "About"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Currently Logged in user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2805
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList16 
      Left            =   8520
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13F4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList32 
      Left            =   8520
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":142E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":177C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":184A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1917C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AB30
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuFines 
         Caption         =   "Fines"
      End
      Begin VB.Menu mnuCard 
         Caption         =   "Issue cards"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuIs 
      Caption         =   "Issue book"
      Begin VB.Menu mnuIssue 
         Caption         =   "&Issue STUDENT"
      End
      Begin VB.Menu mnuSI 
         Caption         =   "Staff Issue"
      End
   End
   Begin VB.Menu mnuRt 
      Caption         =   "Return Book"
      Begin VB.Menu mnuReturn 
         Caption         =   "&Return STUDENT"
      End
      Begin VB.Menu mnuSR 
         Caption         =   "Staff Return"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Begin VB.Menu mnuBookRec 
         Caption         =   "&Book Record"
      End
      Begin VB.Menu mnuMembers 
         Caption         =   "&Members Record"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "R&eport"
      Begin VB.Menu mnuBookRep 
         Caption         =   "&Book Report1"
      End
      Begin VB.Menu mnuBookRep2 
         Caption         =   "Book &Report2"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Student Report"
      End
      Begin VB.Menu mnuUnreturnedBooks 
         Caption         =   "&Unreturned Books"
      End
      Begin VB.Menu mnuFineR 
         Caption         =   "&Fine Report"
      End
      Begin VB.Menu mnuDepoR 
         Caption         =   "&Deposit Report"
      End
   End
   Begin VB.Menu mnuDepo 
      Caption         =   "&Deposit"
      Begin VB.Menu mnuDepoBk 
         Caption         =   "Deposit Book"
      End
      Begin VB.Menu mnuDepoRtrn 
         Caption         =   "Return Deposit"
      End
      Begin VB.Menu mnuDpRcrd 
         Caption         =   "Deposit Record"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertically"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub MDIForm_Load()

    Me.Show
    Set CN = New ADODB.Connection
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MasterFile.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2006"
    If CN.State <> adStateOpen Then MsgBox "Could not establish a connection with the database" & vbNewLine & "The database should exist in ApplicationPath\MasterFile.mdb", vbExclamation, "Database not found!": Unload Me
    frmReturn.FineAmnt = CCur(GetSetting(App.Title, "Settings", "Fine Amount", "2"))
    frmReturn.MaxDays = CInt(GetSetting(App.Title, "Settings", "Max Days", "14"))

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim Form As Form

    For Each Form In Forms
        Unload Form
        Set Form = Nothing
    Next Form
    Set CN = Nothing

End Sub

Private Sub MDIForm_Initialize()
    InitCommonControls
End Sub



Private Sub mnuAbout_Click()

    frmAbout.Show vbModal

End Sub

'Private Sub mnuArrangeIcons_Click()

  '  frmMain.Arrange vbArrangeIcons

'End Sub

Private Sub mnuBookRec_Click()

    With frmBooks
        .Show
        .SetFocus
    End With

End Sub

Private Sub mnuBookRep_Click()

Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblBooks Order by [Accesion Number]", CN, adOpenStatic, adLockReadOnly
    Set drBookList.DataSource = RS
    drBookList.Show
    Set RS = Nothing

End Sub

Private Sub mnuBookRep2_Click()
Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblBooks Order by [Accesion Number]", CN, adOpenStatic, adLockReadOnly
    Set drBookList2.DataSource = RS
    drBookList2.Show
    Set RS = Nothing

End Sub

Private Sub mnuCard_Click()
frmCard.Show vbModal

End Sub

Private Sub mnuCascade_Click()

    frmMain.Arrange vbCascade

End Sub

Private Sub mnuDepoBk_Click()
frmDpissu.Show vbModal

End Sub

Private Sub mnuDepoR_Click()
Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblDepo Order by [Member_ID]", CN, adOpenStatic, adLockReadOnly
    Set drDeposit.DataSource = RS
    drDeposit.Show
    Set RS = Nothing

End Sub


Private Sub mnuDepoRtrn_Click()
frmDprtrn.Show vbModal

End Sub

Private Sub mnuDpRcrd_Click()
With frmDeposit
     .Show
     .SetFocus
  End With
End Sub

Private Sub mnuFineR_Click()
Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblTrans where fines>0 Order by [ID] ", CN, adOpenStatic, adLockReadOnly
    Set drFine.DataSource = RS
    drFine.Show
    Set RS = Nothing

End Sub


Private Sub mnuFines_Click()
  With frmFines
     .Show
     .SetFocus
  End With

End Sub

Private Sub mnuIssue_Click()

    frmIssue.Show vbModal

End Sub

Private Sub mnuMembers_Click()

    With frmMembers
        .Show
        .SetFocus
    End With

End Sub

Private Sub mnuReport_Click()

Dim RS As New ADODB.RecordSet

    RS.Open "SELECT * FROM tblMembers Order by [ID]", CN, adOpenStatic, adLockReadOnly
    Set drMembers.DataSource = RS
    drMembers.Show
    Set RS = Nothing

End Sub

Private Sub mnuReturn_Click()

    frmReturn.Show vbModal

End Sub

Private Sub mnuSettings_Click()

    frmSettings.Show vbModal

End Sub

Private Sub mnuSI_Click()
frmStaffI.Show vbModal
End Sub

Private Sub mnuSR_Click()
frmStaffR.Show vbModal
End Sub

Private Sub mnuTileHorizontal_Click()

    frmMain.Arrange vbTileHorizontal

End Sub

Private Sub mnuTileVertical_Click()

    frmMain.Arrange vbTileVertical

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuUnreturnedBooks_Click()

Dim RS As New ADODB.RecordSet

    RS.Open "SELECT tblTrans.[Accesion Number], tblTrans.[ID], tblBooks.Title, [Name] AS Borrower, tblTrans.[Date_Borrowed] FROM tblMembers INNER JOIN (tblBooks INNER JOIN tblTrans ON tblBooks.[Accesion Number] = tblTrans.[Accesion Number]) ON tblMembers.[ID] = tblTrans.[ID] Where (((tblTrans.Returned) = False)) ORDER BY tblTrans.[Accesion Number];", CN, adOpenStatic, adLockReadOnly
    Set drTransUn.DataSource = RS
    drTransUn.Show
    Set RS = Nothing

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1: PopupMenu mnuIs, , Toolbar1.Buttons(1).Left, Toolbar1.Top + Toolbar1.Height
    Case 2: PopupMenu mnuRt, , Toolbar1.Buttons(2).Left, Toolbar1.Top + Toolbar1.Height
    Case 4: mnuBookRec_Click
    Case 5: mnuMembers_Click
    Case 6: mnuFines_Click
    Case 7: mnuDpRcrd_Click
    Case 8: PopupMenu mnuReports, , Toolbar1.Buttons(7).Left, Toolbar1.Top + Toolbar1.Height
    Case 10: mnuSettings_Click
    Case 11: mnuAbout_Click
    End Select

End Sub
