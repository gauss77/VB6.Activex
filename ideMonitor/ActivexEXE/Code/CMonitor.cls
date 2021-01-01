VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub IniciouObjeto(ByVal psObjeto As String)
    Call modulo.AddObjeto(UCase$(psObjeto))
End Sub

Public Sub TerminouObjeto(ByVal psObjeto As String)
    Call modulo.RemoveObjeto(UCase$(psObjeto))
End Sub

Public Sub ShowMonitor()
    modulo.ShowMonitor
End Sub

Public Sub Destroy()
    On Error Resume Next
    Unload frmObjetos
    Set frmObjetos = Nothing
End Sub
