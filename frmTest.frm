VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "Bogus Request"
      Height          =   495
      Index           =   5
      Left            =   7575
      TabIndex        =   6
      Top             =   4815
      Width           =   1035
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Long request"
      Height          =   495
      Index           =   4
      Left            =   5883
      TabIndex        =   5
      Top             =   4815
      Width           =   1395
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Short request"
      Height          =   495
      Index           =   3
      Left            =   4191
      TabIndex        =   4
      Top             =   4815
      Width           =   1395
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Add User"
      Height          =   495
      Index           =   2
      Left            =   2859
      TabIndex        =   3
      Top             =   4815
      Width           =   1035
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Logoff"
      Height          =   495
      Index           =   1
      Left            =   1527
      TabIndex        =   2
      Top             =   4815
      Width           =   1035
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Logon"
      Height          =   495
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   4815
      Width           =   1035
   End
   Begin VB.TextBox txtOutput 
      Height          =   4530
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   8520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents DB As ActiveXDB.clsDB
Attribute DB.VB_VarHelpID = -1
Private cctDB As cctDB

Private Sub cmdAction_Click(Index As Integer)
    If Index = 0 Then
        DB.Request rqLogon, "Cindy", "Password"
    ElseIf Index = 1 Then
        DB.Request rqLogoff, "Cindy", "Password", "Modified Description Here"
    ElseIf Index = 2 Then
        DB.Request rqAddUser, "Cindy", "Password", "Cindy's Description Here"
    ElseIf Index = 3 Then
        DB.Request rqWaitRequest, 2
    ElseIf Index = 4 Then
        DB.Request rqWaitRequest, 10
    ElseIf Index = 5 Then
        DB.Request rqAddUser, "Cindy", "Password"
    End If
End Sub

Private Sub DB_RequestComplete(ByVal RequestID As Long, ByVal RequestType As ActiveXDB.enumRequest, ByVal Success As Boolean, ByVal Message As String, ByVal Args As Variant)
    Display RequestID & " - " & _
    Choose(RequestType, "Logon", "Logoff", "AddUser", "WaitRequest") & _
    IIf(Success = True, " Success:", " Failure:") & " " & Message
End Sub

Private Sub Form_Load()
    Set cctDB = New cctDB
    Set DB = cctDB.PublicDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set DB = Nothing
    Set cctDB = Nothing
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
    
    With Form1.txtOutput
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

