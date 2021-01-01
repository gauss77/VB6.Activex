VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event EventTeste()

Private Sub MetodoTeste()
    Call AddInMonitor.Gravar.Log("Processo", "Project1.Form1.MetodoTeste")

End Sub



Public Function FuntionTeste()
    Call AddInMonitor.Gravar.Log("Processo", "Project1.Form1.FuntionTeste")

End Function

Public Property Get TesteProperty() As Variant

End Property

Public Property Let TesteProperty(ByVal vNewValue As Variant)
    Call AddInMonitor.Gravar.Log("Processo", "Project1.Form1.Property Let TesteProperty")

End Property

Private Sub Form_Initialize()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.Form1.Form_Initialize")

End Sub

Private Sub Form_Terminate()
    Call AddInMonitor.Gravar.Log("Objeto", "Project1.Form1.Form_Terminate")

End Sub
