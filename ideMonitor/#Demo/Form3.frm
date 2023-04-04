VERSION 5.00
Begin VB.Form frmForm3 
   Caption         =   "Form3"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   LinkTopic       =   "Form3"
   ScaleHeight     =   3045
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
  ideMonitor.IniciouObjeto App.EXEName, Me.Name
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmForm3 = Nothing
End Sub

Private Sub Form_Terminate()
  ideMonitor.TerminouObjeto App.EXEName, Me.Name
End Sub

