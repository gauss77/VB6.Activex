VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EADE62FD-5B6B-444E-A6C6-26CFE520CF78}#1.0#0"; "ideToolBar.ocx"
Begin VB.Form FDTFieldConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuração de Campos"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8460
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAlterar 
      Caption         =   "Alterar DataField"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   4305
      TabIndex        =   3
      Top             =   510
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.OptionButton optAlterar 
      Caption         =   "Alterar Caption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2715
      TabIndex        =   2
      Top             =   510
      Width           =   1590
   End
   Begin Insignia_Toolbar.ideToolbar asxToolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   688
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   4
      ButtonStyle1    =   2
      ButtonStyle2    =   2
      ButtonCaption3  =   "&Salvar  "
      ButtonKey3      =   "Salvar"
      ButtonPicture3  =   "FDTFieldConfig.frx":0000
      ButtonToolTipText3=   "Salvar"
      ButtonCaption4  =   "&Cancelar  "
      ButtonKey4      =   "Cancelar"
      ButtonPicture4  =   "FDTFieldConfig.frx":0352
      ButtonToolTipText4=   "Cancelar"
   End
   Begin MSComctlLib.ListView LView 
      Height          =   5025
      Index           =   1
      Left            =   6105
      TabIndex        =   0
      Top             =   765
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   8864
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Campos da Tela"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView LView 
      Height          =   5025
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   765
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   8864
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Controle"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Index"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Caption"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DataField"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Caption         =   "Total de Campos: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6120
      TabIndex        =   4
      Top             =   525
      Width           =   2205
   End
End
Attribute VB_Name = "FDTFieldConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msKeyTexto    As String
Private msPathTabela  As String
Private mbConfirmado  As Boolean

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
  Select Case ButtonKey
    Case Is = "Salvar"
      Call Salvar
    Case Is = "Cancelar"
      mbConfirmado = False
      Unload Me
  End Select
End Sub

Public Sub AddListaOpcoes(ByRef pConnection As Variant, ByVal SQL As String)
  Dim RS  As ADODB.Recordset
  Dim i   As Integer
  
  On Error GoTo TrataErro:
  Set RS = New ADODB.Recordset
  RS.CursorLocation = adUseServer
  RS.Open SQL, pConnection
  On Error GoTo 0
  
  lblCount.Caption = "Total de Campos: " & RS.Fields.Count
  For i = 0 To RS.Fields.Count - 1
    LView(1).ListItems.Add , RS.Fields(i).Name, RS.Fields(i).Name
  Next
  
  RS.Close
TrataErro:
  Set RS = Nothing
End Sub

Public Function Abrir(ColCampos As Collection, ByVal PathTabela As String) As Boolean
  Dim CTR As Control
  
  MousePointer = vbHourglass

  msPathTabela = PathTabela
  For Each CTR In ColCampos
    With LView(0).ListItems
      .Add , , TypeName(CTR)
      .Item(.Count).SubItems(1) = CTR.Index
      .Item(.Count).SubItems(2) = CTR.Tag
      .Item(.Count).SubItems(3) = CTR.DataField
      Call CheckListDTField(CTR.DataField)
    End With
  Next
  
  MousePointer = vbDefault
  Show vbModal
  Abrir = mbConfirmado
End Function

Private Sub UnCheckListDTField(Key As String)
  Dim i As Integer
  
  If Key = "" Then Exit Sub
  
  For i = 1 To LView(1).ListItems.Count
    If LView(1).ListItems(i).Text = Key Then
      LView(1).ListItems(i).Checked = False
      Exit Sub
    End If
  Next
End Sub

Private Sub CheckListDTField(ByVal Key As String, Optional bGeral As Boolean)
  Dim i As Integer, n As Integer
  
  
  If bGeral = True Then
    For i = 1 To LView(1).ListItems.Count
      LView(1).ListItems(i).Checked = False
    Next
    
    For i = 1 To LView(0).ListItems.Count
      Key = LView(0).ListItems(i).Text
      For n = 1 To LView(1).ListItems.Count
        If LView(1).ListItems(n).Text = Key Then
          LView(1).ListItems(n).Checked = True
          Exit For
        End If
      Next
    Next
  
  Else
    If Key = "" Then Exit Sub
    
    For i = 1 To LView(1).ListItems.Count
      If LView(1).ListItems(i).Text = Key Then
        LView(1).ListItems(i).Checked = True
        Exit For
      End If
    Next
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set FDTFieldConfig = Nothing
End Sub

Private Sub LView_DblClick(Index As Integer)
  Dim sDS As String, sCap As String
  
  Select Case Index
    Case Is = 0
      With LView(Index).SelectedItem
        sCap = .SubItems(2)
        sDS = .SubItems(3)
        
        If optAlterar(0).Value Then
          'Alterar Caption
          sCap = InputBox("DataField do Campo: " & sDS, "Mudar Caption", sCap)
          If sCap <> "" Then
            .SubItems(2) = sCap
          End If
        Else
          'Alterar DataField
          sDS = InputBox("Caption do Campo: " & sCap, "Mudar DataField", sDS)
          If sDS <> "" Then
            .SubItems(3) = sDS
            Call CheckListDTField("", True)
          End If
        End If
      End With
  End Select
End Sub

Private Sub Salvar()
  Dim oConn       As ADODB.Connection
  Dim RSPrimary   As ADODB.Recordset
  Dim sCTR  As String, nIDX As Integer
  Dim sCap  As String, sField As String
  Dim SQLShape As String
  
  SQLShape = "SHAPE APPEND NEW adChar(20) AS Controle,NEW adInteger AS Index, " & _
                          "NEW adChar(30) AS Caption, NEW adChar(30) AS DataField, " & _
             "((SHAPE APPEND NEW adChar(20) AS Controle) as Secondary  " & _
             "RELATE Controle TO Controle)"
    
  'Iniciamos a conexão
  Set oConn = New ADODB.Connection
  With oConn
      .Provider = "MSDataShape"
      .ConnectionString = "Data Provider=None"
      .Open
  End With

  Set RSPrimary = New ADODB.Recordset
  With RSPrimary
    Set .ActiveConnection = oConn
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open SQLShape
  End With
  
  Dim i As Integer
  For i = 1 To LView(0).ListItems.Count
    With LView(0)
      sCTR = .ListItems(i).Text
      nIDX = .ListItems(i).SubItems(1)
      sCap = .ListItems(i).SubItems(2)
      sField = .ListItems(i).SubItems(3)
      RSPrimary.AddNew Array("Controle", "Index", "Caption", "DataField"), Array(sCTR, nIDX, sCap, sField)
    End With
  Next
  On Error Resume Next
  Kill msPathTabela
  On Error GoTo 0
  RSPrimary.Save msPathTabela, adPersistADTG
  
  RSPrimary.Close
  oConn.Close
  Set RSPrimary = Nothing
  Set oConn = Nothing
  
  mbConfirmado = True
  Unload Me
End Sub

Private Sub LView_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
  Dim sText As String
  If Index = 1 Then
    If Item.Checked = True Then
      sText = LView(0).SelectedItem.SubItems(3)
      If sText <> Item.Text Then
        Call UnCheckListDTField(sText)
        LView(0).SelectedItem.SubItems(3) = Item.Text
      End If
    End If
  End If
End Sub

Private Sub LView_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
  If Index = 0 Then msKeyTexto = ""
End Sub

Private Sub LView_KeyPress(Index As Integer, KeyAscii As Integer)
  Static sTexto As String
  Static i As Integer
  
  If Index = 0 Then
    If optAlterar(0).Value Then 'Se Opcao de Alterar Caption
      'Entao SubItem = 2 : Ref:Caption
      If i <> 2 Then msKeyTexto = ""
      i = 2
    Else
      'Se não SubItem = 2 : Ref:DataField
      If i <> 3 Then msKeyTexto = ""
      i = 3
    End If
    
    Select Case KeyAscii
      Case Is = vbKeyReturn
        If msKeyTexto <> "" Then
          LView(Index).SelectedItem.SubItems(i) = msKeyTexto
          If i = 3 Then Call CheckListDTField("", True)

        End If
        msKeyTexto = ""
      Case Is = vbKeyEscape
        msKeyTexto = ""
      Case Is = vbKeyBack
        If msKeyTexto <> "" Then msKeyTexto = Left(msKeyTexto, Len(msKeyTexto) - 1)
        LView(Index).SelectedItem.SubItems(i) = msKeyTexto
        If i = 3 Then Call CheckListDTField("", True)
      Case Else
        msKeyTexto = msKeyTexto & Chr(KeyAscii)
        LView(Index).SelectedItem.SubItems(i) = msKeyTexto
        If i = 3 Then Call CheckListDTField("", True)
    End Select
  End If
End Sub

Private Sub optAlterar_Click(Index As Integer)
  If Index = 0 Then
    optAlterar(Index).FontBold = optAlterar(Index).Value
    optAlterar(1).FontBold = optAlterar(1).Value
  Else
    optAlterar(Index).FontBold = optAlterar(Index).Value
    optAlterar(0).FontBold = optAlterar(0).Value
  End If
End Sub
