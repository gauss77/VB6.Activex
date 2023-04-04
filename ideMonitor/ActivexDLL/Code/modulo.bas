Attribute VB_Name = "modulo"
Option Explicit

Private mMonitor As Object

Public Property Get MonitorEXE() As Object
    If mMonitor Is Nothing Then
        Set mMonitor = GetObject("", "ideObjectMonitor.CMonitor")
    End If
    Set MonitorEXE = mMonitor
End Property

Public Sub Destroy()
    MonitorEXE.Destroy
    Set mMonitor = Nothing
End Sub
