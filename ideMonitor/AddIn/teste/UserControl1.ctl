VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.UserControl1.UserControl_Initialize")

End Sub

Private Sub UserControl_Terminate()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.UserControl1.UserControl_Terminate")

End Sub
