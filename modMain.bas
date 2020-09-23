Attribute VB_Name = "modMain"
Option Explicit

Public DB As clsDB
Public QuitGame As Boolean

Public Sub Main()
    Set DB = New clsDB
    frmMain.Show
End Sub

Public Sub DBCallback()
    If QuitGame = True Then
        ShutDown
    Else
        DB.ProcessRequests
    End If
End Sub

Public Sub ShutDown()
    Set DB = Nothing
    Unload frmMain
    Set frmMain = Nothing
End Sub
