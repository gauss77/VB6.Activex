VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exemplor de Uso do Monitor de Objetos...."
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Destruir Monitor"
      Height          =   435
      Left            =   2985
      TabIndex        =   4
      Top             =   600
      Width           =   1845
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Instanciar Form3"
      Height          =   435
      Left            =   3930
      TabIndex        =   3
      Top             =   90
      Width           =   1830
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Form2"
      Height          =   435
      Left            =   2025
      TabIndex        =   2
      Top             =   90
      Width           =   1830
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Monitor"
      Height          =   435
      Left            =   1050
      TabIndex        =   1
      Top             =   600
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Form1"
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   1830
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  frmForm1.Show
End Sub

Private Sub Command2_Click()
  frmForm2.Show
End Sub

Private Sub Command3_Click()
  ideMonitor.ShowMonitor
End Sub

Private Sub Command4_Click()
  Dim oF As Form
  
  Set oF = New frmForm3
  oF.Show
  Set oF = Nothing
End Sub

Private Sub Command5_Click()
  ideMonitor.Destroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
