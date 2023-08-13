VERSION 5.00
Object = "{C6FEE5AC-DF5F-47A6-BE77-6DCE10AA8AB9}#4.2#0"; "ideDSControl.ocx"
Begin VB.Form frmMaster 
   BackColor       =   &H00C8D0D4&
   Caption         =   "frmMaster"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   7890
   WindowState     =   2  'Maximized
   Begin Insignia_DSControl.ideDSControl ideDSControl1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1429
      CaptionColor    =   -2147483630
      BackColor       =   13160660
      ButtonColor     =   13160660
      ButtonColorDesab=   9936289
      ButtonsExtras   =   0
      ButtonType      =   7
      Modelo          =   0
      Operacao        =   0
      Permissoes      =   0
   End
   Begin VB.TextBox Text1 
      DataField       =   "MEMO"
      DataSource      =   "XDSMaster1"
      Height          =   765
      Index           =   10
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmMaster.frx":0000
      Top             =   3135
      Width           =   2760
   End
   Begin VB.TextBox Text1 
      DataField       =   "NUMERICO"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   9
      Left            =   390
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2655
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "FONE"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   8
      Left            =   3255
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2100
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "HORA"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   7
      Left            =   1815
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2100
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "CGC"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   6
      Left            =   375
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2100
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "CPF"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   5
      Left            =   3255
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1620
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "CEP"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   4
      Left            =   1815
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1620
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "DATA"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   3
      Left            =   375
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1620
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "SIMNAO"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   2
      Left            =   3255
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1140
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "MOEDA"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   1
      Left            =   1815
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1140
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "TEXTO"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1140
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Memo"
      Height          =   195
      Index           =   10
      Left            =   390
      TabIndex        =   21
      Top             =   2940
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numerico"
      Height          =   195
      Index           =   9
      Left            =   390
      TabIndex        =   20
      Top             =   2460
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fone"
      Height          =   195
      Index           =   8
      Left            =   3255
      TabIndex        =   19
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hora"
      Height          =   195
      Index           =   7
      Left            =   1815
      TabIndex        =   18
      Top             =   1905
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CGC"
      Height          =   195
      Index           =   6
      Left            =   375
      TabIndex        =   17
      Top             =   1905
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CPF"
      Height          =   195
      Index           =   5
      Left            =   3255
      TabIndex        =   16
      Top             =   1425
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Index           =   4
      Left            =   1815
      TabIndex        =   15
      Top             =   1425
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA"
      Height          =   195
      Index           =   3
      Left            =   375
      TabIndex        =   14
      Top             =   1425
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SIMNAO"
      Height          =   195
      Index           =   2
      Left            =   3255
      TabIndex        =   13
      Top             =   945
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MOEDA"
      Height          =   195
      Index           =   1
      Left            =   1815
      TabIndex        =   12
      Top             =   945
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEXTO"
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   11
      Top             =   945
      Width           =   540
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Const SQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
  
  With ideDSControl1

    .Conectar "SELECT * FROM TABELA", SQL & App.Path & "\Teste.mdb;"
    .MontarPesquisa "Código ID,ID,######|Nome,TEXTO,|Data,DATA,##/##/####|Sim Não,SIMNAO,|Numero,NUMRICO,######|HORA,HORA,##:##:##"
    .DataSource.AppName = App.EXEName
    .Permissoes = peTodos
    .ButtonsExtras = beSearch_Requery
  
    Dim oTxt As TextBox
  
    For Each oTxt In Text1
      Set oTxt.DataSource = .DataSource.rs
    Next
    
    If .DataSource.rs.RecordCount > 0 Then
      .DataSource.MoveLast
    End If
  End With
End Sub

Private Sub ideDSControl1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Static I As Integer
  
  I = I + 1
  Debug.Print "MoveComplete", I
End Sub

'Private Sub ideDSControl1_Operacao(ByVal eOperacao As OCXDSControl.eDSOperacao, ByVal eOperacaoAnterior As OCXDSControl.eDSOperacao)
''  Dim O As XDSField
''  For Each O In XDSField
''    O.TypeOperation = eOperacao
''  Next
'End Sub
