VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMonitor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub IniciouObjeto(AppEXEName As String, TypeNameObjeto As String)
    Call modulo.MonitorEXE.IniciouObjeto(AppEXEName & "." & TypeNameObjeto)
End Sub

Public Sub TerminouObjeto(AppEXEName As String, TypeNameObjeto As String)
    Call modulo.MonitorEXE.TerminouObjeto(AppEXEName & "." & TypeNameObjeto)
End Sub

Public Sub ShowMonitor()
    Call modulo.MonitorEXE.ShowMonitor
End Sub

Public Sub Destroy()
    Call modulo.Destroy
End Sub
