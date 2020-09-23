VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Events
Event RequestComplete(ByVal RequestID As Long, ByVal RequestType As enumRequest, ByVal Success As Boolean, ByVal Message As String, ByVal Args As Variant)

'Variables
Private DB As Database
Private Requests As Collection
Private Timer As Long
Private Ctr As Long

'Enums
Public Enum enumRequest
    rqLogon = 1
    rqLogoff = 2
    rqAddUser = 3
    rqWaitRequest = 4
End Enum

'Declarations
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Class_Initialize()
    'Open the database
    Set DB = OpenDatabase(App.Path & "\ActiveX DB.mdb")
    'Clear the inuse flags
    DB.Execute "UPDATE Users SET InUse = FALSE"
    'Initialize the requests collection
    Set Requests = New Collection
    'Start the timer
    Timer = SetTimer(0, 0, 500, AddressOf DBCallback)
End Sub

Private Sub Class_Terminate()
    'Stop the timer
    KillTimer 0, Timer
    'Destroy the requests collection
    Set Requests = Nothing
    'Close the database
    DB.Close
End Sub

Public Function Request(ByVal RequestType As enumRequest, ParamArray Args() As Variant) As Long
    'Don't actually perform the request here!
    'Instead, save it in a collection for future use.
    'The timer process will periodically check the queue, and execute any
    'requests it finds.
    
    'Add a new request to the queue, and assign it an id.
    Dim NewRequest As clsRequest
    'Assign a new ID to the request
    Ctr = Ctr + 1
    'Create a new request object, and fill in the info
    Set NewRequest = New clsRequest
    NewRequest.ID = Ctr
    NewRequest.Action = RequestType
    NewRequest.Args = Args
    'Add the new request to the collection
    Requests.Add NewRequest
    'Return the ID of this request.
    Request = Ctr
End Function

Friend Sub ProcessRequests()
    Dim Request As clsRequest
    'Called by the timer to process any waiting requests.
    Do While Requests.Count > 0
        'Grab next request
        Set Request = Requests(1)
        'Remove it from the list
        Requests.Remove 1
        'Process it
        Select Case Request.Action
            Case rqLogon: Logon Request
            Case rqLogoff: Logoff Request
            Case rqAddUser: AddUser Request
            Case rqWaitRequest: WaitRequest Request
        End Select
    Loop
End Sub

Private Sub Display(ByVal s As String)
    'Display text on the console
    
    Dim ShortText As String
    
    'Ensure it ends in vbcrlf
    If Len(s) > 1 Then
        If Right$(s, 2) <> vbCrLf Then
            s = s & vbCrLf
        End If
    Else
        s = s & vbCrLf
    End If
    
    With frmMain.txtOutput
        'Shorten the textbox if necessary to keep it under 15,000 characters.
        If Len(.Text) + Len(s) > 15000 Then
            ShortText = Right$(.Text, Len(.Text) - InStrRev(.Text, vbCrLf, 5000) - 1)
            .Text = ShortText
        End If
        'Add the text to the textbox
        .SelStart = Len(.Text)
        .SelText = s
        .SelStart = Len(.Text)
        .Refresh
    End With
End Sub

'--------------------------------------------------------------------------------
' DB processing goes here.
'--------------------------------------------------------------------------------
Private Sub WaitRequest(ByVal Request As clsRequest)
    Dim WaitTime As Long, EndTime As Long
    
    If VerifyArgumentCount(Request, 1) = False Then
        Exit Sub
    End If
    
    WaitTime = Request.Args(0)
    EndTime = timeGetTime + (WaitTime * 1000)
    Do While timeGetTime < EndTime
        DoEvents
        Sleep 1
    Loop
    
    RaiseEvent RequestComplete(Request.ID, Request.Action, True, "Waited " & WaitTime & " seconds.", "")
End Sub

Private Sub Logon(ByVal Request As clsRequest)
    Dim SQL As String, rsTemp As Recordset
    Dim Name As String, Password As String
    
    If VerifyArgumentCount(Request, 2) = False Then
        Exit Sub
    End If

    Name = Request.Args(0)
    Password = Request.Args(1)
    
    SQL = "SELECT * FROM Users WHERE LCASE(NAME) = '" & LCase(Name) & "'"
    Set rsTemp = DB.OpenRecordset(SQL)
    If rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        If LCase(rsTemp!Password) = LCase(Password) Then
            If rsTemp!InUse = False Then
                rsTemp.Edit
                rsTemp!InUse = True
                rsTemp!LastLogon = Now
                rsTemp.Update
                RaiseEvent RequestComplete(Request.ID, Request.Action, True, "Logged on name " & Name & " password " & Password, rsTemp!Description)
            Else
                RaiseEvent RequestComplete(Request.ID, Request.Action, False, "Logonid " & Name & " in use!", "")
            End If
        Else
            RaiseEvent RequestComplete(Request.ID, Request.Action, False, "Invalid password supplied for " & Name & "!", "")
        End If
    ElseIf rsTemp.RecordCount = 0 Then
        RaiseEvent RequestComplete(Request.ID, Request.Action, False, "No such user found!", "")
    Else
        RaiseEvent RequestComplete(Request.ID, Request.Action, False, "Multiple users for " & Name & " found!", "")
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
End Sub

Private Sub Logoff(ByVal Request As clsRequest)

    Dim SQL As String, rsTemp As Recordset
    Dim Name As String, Password As String, Description As String
    
    If VerifyArgumentCount(Request, 3) = False Then
        Exit Sub
    End If

    Name = Request.Args(0)
    Password = Request.Args(1)
    Description = Request.Args(2)
    
    SQL = "SELECT * FROM Users WHERE LCASE(Name) = '" & LCase(Name) & _
    "' AND LCASE(Password) = '" & LCase(Password) & "'"
    Set rsTemp = DB.OpenRecordset(SQL)
    If rsTemp.RecordCount = 1 Then
        rsTemp.MoveFirst
        rsTemp.Edit
        rsTemp!Description = Description
        rsTemp!InUse = False
        rsTemp.Update
        RaiseEvent RequestComplete(Request.ID, Request.Action, True, "Logged off " & Name & " password " & Password, "")
    ElseIf rsTemp.RecordCount = 0 Then
        RaiseEvent RequestComplete(Request.ID, Request.Action, False, "No such user " & Name & " found!", "")
    Else
        RaiseEvent RequestComplete(Request.ID, Request.Action, False, "Multiple users for " & Name & " found!", "")
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
End Sub

Private Sub AddUser(ByVal Request As clsRequest)
    Dim SQL As String, rsTemp As Recordset
    Dim Name As String, Password As String, Description As String
    
    If VerifyArgumentCount(Request, 3) = False Then
        Exit Sub
    End If

    Name = Request.Args(0)
    Password = Request.Args(1)
    Description = Request.Args(2)
    
    SQL = "SELECT * FROM Users WHERE LCASE(NAME) = '" & LCase(Name) & "'"
    Set rsTemp = DB.OpenRecordset(SQL)
    If rsTemp.RecordCount > 0 Then
        RaiseEvent RequestComplete(Request.ID, Request.Action, False, "User " & Name & " already exists!", "")
    Else
        rsTemp.AddNew
        rsTemp!Name = Name
        rsTemp!Password = Password
        rsTemp!Description = Description
        rsTemp.Update
        RaiseEvent RequestComplete(Request.ID, Request.Action, True, "Added user " & Name & " password " & Password, "")
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
End Sub

Private Function VerifyArgumentCount(ByRef Request As clsRequest, ByVal NumArgumentsRequired As Long) As Boolean
    If UBound(Request.Args, 1) + 1 < NumArgumentsRequired Then
        RaiseEvent RequestComplete(Request.ID, Request.Action, False, "Incorrect number of arguments! Requires " & NumArgumentsRequired, "")
        VerifyArgumentCount = False
    Else
        VerifyArgumentCount = True
    End If
End Function

