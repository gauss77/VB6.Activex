VERSION 5.00
Object = "{C6FEE5AC-DF5F-47A6-BE77-6DCE10AA8AB9}#4.2#0"; "ideDSControl.ocx"
Begin VB.Form frmSimple 
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7335
   WindowState     =   2  'Maximized
   Begin Insignia_DSControl.ideDSControl ideDSControl1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   714
      CaptionColor    =   -2147483630
      BackColor       =   13160660
      ButtonColor     =   13160660
      ButtonColorDesab=   9936289
      ButtonsExtras   =   0
      ButtonType      =   7
      Modelo          =   1
      Operacao        =   0
      Permissoes      =   0
   End
   Begin VB.TextBox Text1 
      DataField       =   "TEXTO"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   735
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "MOEDA"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   1
      Left            =   1530
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   735
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "SIMNAO"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   2
      Left            =   2970
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   735
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "DATA"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "CEP"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   4
      Left            =   1530
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "CPF"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   5
      Left            =   2970
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1215
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "CGC"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   6
      Left            =   90
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1695
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "HORA"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   7
      Left            =   1530
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1695
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "FONE"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   8
      Left            =   2970
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1695
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "NUMERICO"
      DataSource      =   "XDSMaster1"
      Height          =   285
      Index           =   9
      Left            =   4470
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   735
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      DataField       =   "MEMO"
      DataSource      =   "XDSMaster1"
      Height          =   765
      Index           =   10
      Left            =   4470
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSimple.frx":0000
      Top             =   1215
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEXTO"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   21
      Top             =   540
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MOEDA"
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   20
      Top             =   540
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SIMNAO"
      Height          =   195
      Index           =   2
      Left            =   2970
      TabIndex        =   19
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   18
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Index           =   4
      Left            =   1530
      TabIndex        =   17
      Top             =   1020
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CPF"
      Height          =   195
      Index           =   5
      Left            =   2970
      TabIndex        =   16
      Top             =   1020
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CGC"
      Height          =   195
      Index           =   6
      Left            =   90
      TabIndex        =   15
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hora"
      Height          =   195
      Index           =   7
      Left            =   1530
      TabIndex        =   14
      Top             =   1500
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fone"
      Height          =   195
      Index           =   8
      Left            =   2970
      TabIndex        =   13
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Numerico"
      Height          =   195
      Index           =   9
      Left            =   4470
      TabIndex        =   12
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Memo"
      Height          =   195
      Index           =   10
      Left            =   4470
      TabIndex        =   11
      Top             =   1020
      Width           =   435
   End
End
Attribute VB_Name = "frmSimple"
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
  
    Dim O As TextBox
    
    For Each O In Text1
      Set O.DataSource = .DataSource.RS
    Next
  End With
End Sub
