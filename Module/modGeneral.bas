Attribute VB_Name = "modGeneral"
Option Explicit
Global pWords As String
Dim rsUserPassword As New ADODB.RecordSet
Global varUserPassword As String
Dim Conn As New ADODB.Connection
Dim rsPassword As New ADODB.RecordSet
Dim rsPassword1 As New ADODB.RecordSet



Public CN As ADODB.Connection
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Main()

'Checks for previous version of CLAM


    On Error Resume Next
        If App.PrevInstance Then
            MsgBox "An instance of " & App.Title & " is already running!" & vbNewLine & "You cannot run two instances of this application at the same time", vbCritical, "Application already running"
            End
        Else
            frmMain.Show
        End If

End Sub

Public Sub Handler(Error As ErrObject)

'Shows msgbox for unhandled errors only when error has truly occured,
'i.e. err<>0

    If Error.Number <> 0 Then
        MsgBox "Error Number: " & Error.Number & vbNewLine & Error.Description, vbExclamation, "Unexpected Error"
    End If

End Sub

Public Sub CenterObj(ByRef ChildObj As Variant, ByVal ParentObj As Variant)

'This procedure centers an object over another object
'works with phot control application


    ChildObj.Move (ParentObj.Width - ChildObj.Width) / 2 + ChildObj.Left, (ParentObj.Height - ChildObj.Height) / 2 + ParentObj.Top

End Sub

Public Sub HighLight(ByRef sObj As Object)

'Procedure highlights text in a textbox

    With sObj
        .SelStart = 0
        .SelLength = Len(sObj.Text)
    End With

End Sub

Public Sub FillCombo(ByRef sCombo As ComboBox, ByVal sRS As ADODB.RecordSet, Sort As Boolean)

'This procedure fills a combo box with field name from a given recordset
'used in the combo boxes for Searching/Filtering/Sorting records

Dim X As Long

    With sCombo
        For X = 0 To sRS.Fields.Count - 1
            If sRS.Fields.Item(X).Name = "Picture" Then GoTo nexus
            If Sort Then
                .AddItem "[" & sRS.Fields.Item(X).Name & "] Asc"
                .AddItem "[" & sRS.Fields.Item(X).Name & "] Desc"
            Else 'NOT SORT...
                .AddItem sRS.Fields.Item(X).Name
            End If
nexus:
        Next X
    End With

End Sub

Public Sub FindRecord(ByRef sRS As ADODB.RecordSet, ByVal sField As String, ByVal isString As Boolean, ByVal sStr As String, ByVal sNum As Long)

'This procedure finds a record in the selected recordset
'and sets its absolute position with the found record.

    On Local Error Resume Next
        With sRS
            .Filter = adFilterNone
            .Requery
            .MoveFirst
            If isString Then
                .Find sField & " = '" & sStr & "'"
            Else
                .Find sField & " = " & sNum
            End If
        End With

End Sub

Public Function RecordExists(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, ByRef sEntryField As Object) As Boolean

Dim RS As New ADODB.RecordSet

    RS.Open "Select * From " & sTable & " Where [" & sField & "] = '" & sStr & "'", CN, adOpenStatic, adLockReadOnly
    If RS.RecordCount < 1 Then
        RecordExists = False
    Else
        MsgBox "The adding of new entry cannot be done because " & sStr & " already" & vbCrLf & "exists in the recordset. Please check and change it." & vbCrLf & vbCrLf & "Note: Duplication of entries is not allowed in this application.", vbExclamation
        HighLight sEntryField
        RecordExists = True
    End If
    Set RS = Nothing

End Function

Public Sub Navigate(Index As Integer, RecordSet As ADODB.RecordSet)

    On Local Error Resume Next
        With RecordSet
            Select Case Index
            Case 0
                If Not .RecordCount <= 1 Then
                    .MoveFirst
                End If
            Case 3
                If Not .RecordCount <= 1 Then
                    .MoveLast
                End If
            Case 2
                If Not .AbsolutePosition >= .RecordCount Or .RecordCount <= 1 Then
                    .MoveNext
                End If
            Case 1
                If Not .AbsolutePosition <= 1 Then
                    .MovePrevious
                End If
            End Select
        End With

End Sub

Public Sub LineMove(MoveLine As Line, FixedLine As Line)

'Sub used to align one line over the other

    MoveLine.X1 = FixedLine.X1
    MoveLine.X2 = FixedLine.X2
    MoveLine.Y1 = FixedLine.Y1
    MoveLine.Y2 = FixedLine.Y2

End Sub

Public Sub ShellEx(PathName As String)
'Sub used to open a non-excutable file
    If ShellExecute(&O0, "Open", PathName, vbNullString, vbNullString, 1) < 33 Then
        Handler Err
    End If

End Sub

Public Sub Pword()
Call Connect
rsPassword.Open "Select * From SECURITY_PASSWORD ", Conn, adOpenStatic, adLockOptimistic
    pWords = (rsPassword.Fields(0))
Set rsPassword = Nothing
Set Conn = Nothing
End Sub
Public Sub updatePword()
Call Connect
rsPassword1.Open "Select * From SECURITY_PASSWORD ", Conn, adOpenStatic, adLockOptimistic
    rsPassword1.Fields(0) = (frmPass.Text2.Text)
    rsPassword1.Update
    MsgBox "Changes has been successfully save.", vbInformation, "Library System"
    Unload frmPass
Set rsPassword1 = Nothing
Set Conn = Nothing
End Sub
Public Sub UserPassword()
Call Connect
rsUserPassword.Open "Select * From SECURITY_PASSWORD ", Conn, adOpenStatic, adLockOptimistic
    varUserPassword = (rsUserPassword.Fields(0))
Set rsUserPassword = Nothing
Set Conn = Nothing
End Sub

Public Sub Connect()
On Error Resume Next
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MasterFile.mdb;Persist Security Info=False;Jet OLEDB:Database Password=2006"
End Sub
