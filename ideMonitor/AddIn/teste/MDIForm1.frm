VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Initialize()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.MDIForm1.MDIForm_Initialize")

End Sub

Private Sub MDIForm_Terminate()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.MDIForm1.MDIForm_Terminate")

End Sub
