VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Get TesteProperty() As Variant
    Call AddInMonitor.Gravar.Log("Processo", "Project1.Form2.Property Get TesteProperty")

End Property

Private Sub Form_Initialize()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.Form2.Form_Initialize")

End Sub

Private Sub Form_Terminate()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.Form2.Form_Terminate")

End Sub
