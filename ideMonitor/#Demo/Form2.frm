VERSION 5.00
Begin VB.Form frmForm2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Initialize()
  ideMonitor.IniciouObjeto App.EXEName, Me.Name
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmForm2 = Nothing
End Sub

Private Sub Form_Terminate()
  ideMonitor.TerminouObjeto App.EXEName, Me.Name
End Sub

