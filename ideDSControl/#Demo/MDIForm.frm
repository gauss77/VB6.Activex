VERSION 5.00
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   0
      ScaleHeight     =   5370
      ScaleWidth      =   1860
      TabIndex        =   0
      Top             =   0
      Width           =   1890
      Begin VB.CommandButton Command1 
         Caption         =   "XDSMaster"
         Height          =   435
         Left            =   75
         TabIndex        =   5
         Top             =   1800
         Width           =   1650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "XDSSimple"
         Height          =   435
         Left            =   75
         TabIndex        =   4
         Top             =   2280
         Width           =   1650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Carregar Tabela"
         Height          =   435
         Left            =   75
         TabIndex        =   3
         Top             =   2925
         Width           =   1650
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   " [Sair] "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   930
         TabIndex        =   2
         Top             =   4710
         Width           =   465
      End
      Begin VB.Image imgsair 
         Height          =   480
         Left            =   540
         Picture         =   "MDIForm.frx":0000
         Top             =   4425
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Para criar mais de uma instância de sua aplicação (simulando um ambiente multiusuário) , clique duas vezes no icone."
         ForeColor       =   &H00808080&
         Height          =   1335
         Left            =   75
         TabIndex        =   1
         Top             =   90
         Width           =   1680
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  'Abre uma instância dodo formulario de dados

  Dim frmNovo As frmMaster
  Static intFormContador As Integer
  
  On Error GoTo trata_erros
  
  intFormContador = intFormContador + 1
  
  Set frmNovo = New frmMaster
  Load frmNovo
  frmNovo.Caption = "Formulários de Dados Multiusuário, Instancia => #" & intFormContador
  frmNovo.Show
  
  Exit Sub
trata_erros:
  MsgBox Err.Description, vbInformation, " Erros "

End Sub


Private Sub Command2_Click()
'Abre uma instância dodo formulario de dados

'    Dim frmNovo As frmStandart
'    Static intFormContador As Integer
'
'    On Error GoTo trata_erros
'
'    intFormContador = intFormContador + 1
'
'    Set frmNovo = New frmStandart
'    Load frmNovo
'    frmNovo.Caption = "Formulários de Dados Multiusuário, Instancia => #" & intFormContador
'    frmNovo.Show
'
'    Exit Sub
'trata_erros:
'     MsgBox Err.Description, vbInformation, " Erros "

End Sub

Private Sub Command3_Click()
'Abre uma instância dodo formulario de dados

    Dim frmNovo As frmSimple
    Static intFormContador As Integer
    
    On Error GoTo trata_erros
    
    intFormContador = intFormContador + 1
    
    Set frmNovo = New frmSimple
    Load frmNovo
    frmNovo.Caption = "Formulários de Dados Multiusuário, Instancia => #" & intFormContador
    frmNovo.Show
    
    Exit Sub
trata_erros:
     MsgBox Err.Description, vbInformation, " Erros "

End Sub

Private Sub Command4_Click()
  Dim XDS As New CDSControl
  
  Const SQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
  
  XDS.Conectar "SELECT * FROM TABELA", SQL & App.Path & "\Teste.mdb;"
  
  Command4.Enabled = False
  Dim I As Integer
  For I = 10 To 2000
    XDS.AddNew Array("TEXTO", "DATA", "CEP", "CGC", "CPF", "FONE", "HORA", "NUMERICO"), _
               Array("TEXTO_" & I, "01/01/2003", "35100-000", "00.000.000/0000-00", "038.065.976-00", "(33)8801-7394", "04:11:00", I)
    XDS.Update
  Next
  Command4.Enabled = True

End Sub

Private Sub imgsair_Click()
  Unload Me
End Sub

Private Sub Label2_Click()
  Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim I As Integer
    
    On Error GoTo trata_erros
    
    While Forms.Count > 1
        I = 0
        While Forms(I).Caption = Me.Caption
             I = I + 1
        Wend
        Unload Forms(I)
    Wend
    
    Unload Me
    End

    Exit Sub
trata_erros:
    MsgBox Err.Description
End Sub
