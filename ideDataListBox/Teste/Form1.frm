VERSION 5.00
Object = "{C5229D3F-48C7-45CD-88A2-1F2E29893091}#3.0#0"; "ideDataListBox.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboColunas 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   6465
      List            =   "Form1.frx":0010
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3930
      Width           =   750
   End
   Begin OCXDataListBox.ideDataListBox ideDataListBox1 
      Align           =   1  'Align Top
      Height          =   3690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   6509
      Caption         =   "Selecionar Itens:"
      CampoCodigo     =   ""
      CampoDescricao  =   ""
      CampoDescricaoFormat=   ""
      CampoAgrupamento=   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Selecionar"
      Height          =   405
      Left            =   5265
      TabIndex        =   2
      Top             =   3855
      Width           =   1050
   End
   Begin VB.TextBox Text 
      Height          =   405
      Left            =   2130
      TabIndex        =   1
      Text            =   "1,3,5,7,9"
      Top             =   3855
      Width           =   3060
   End
   Begin VB.CommandButton cmdCarregar 
      Caption         =   "Carregar"
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   3855
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pv_oRS As ADODB.Recordset

Private Sub cboColunas_Change()
    cboColunas_Click
End Sub

Private Sub cboColunas_Click()
    If IsNumeric(cboColunas.Text) Then
        ideDataListBox1.Columns = cboColunas.Text
    End If
End Sub

Private Sub cmdCarregar_Click()
    
    Call Carregar
'    ideDataListBox1.CampoCodigo = "CODFIL"
    ideDataListBox1.CampoDescricao = "CODFIL;CODUF;NOMFIL"
    ideDataListBox1.CampoDescricaoFormat = "0000;U;P"
    ideDataListBox1.CampoAgrupamento = "CODUF"
    
    Set ideDataListBox1.Recordset = pv_oRS
End Sub

Public Sub Carregar()
On Error GoTo TrataErro:
    Dim sSQL    As String
    Dim sUF     As String
    Dim nL1     As Integer
    
    Screen.MousePointer = vbHourglass
    
    Set pv_oRS = New ADODB.Recordset
    
    With pv_oRS
        .Fields.Append "CODFIL", adDouble
        .Fields.Append "NOMFIL", adVarChar, 50
        .Fields.Append "CODUF", adChar, 2
        .Open
        nL1 = 0
        While nL1 < 500
            nL1 = nL1 + 1
            .AddNew
            .Fields("CODFIL") = nL1
            .Fields("NOMFIL") = "FiLiAL DAS COVES "
            .Fields("CODUF") = IIf(nL1 <= 100, "AA", IIf(nL1 <= 200, "BB", IIf(nL1 <= 300, "CC", IIf(nL1 <= 400, "DD", "EE"))))
        Wend
    End With
    
Sair:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
TrataErro:
    Set pv_oRS = Nothing
    MsgBox Err.Number & " - " & Err.Source & vbCrLf & Err.Description
    GoTo Sair
End Sub

Private Sub cmdSelect_Click()
    ideDataListBox1.Selecionar Text.Text
End Sub
