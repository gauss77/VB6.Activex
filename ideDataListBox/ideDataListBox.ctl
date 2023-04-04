VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.UserControl ideDataListBox 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3255
   ScaleWidth      =   8055
   ToolboxBitmap   =   "ideDataListBox.ctx":0000
   Begin VB.Frame fraBorda 
      BackColor       =   &H00C0C0C0&
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   7875
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   135
         Left            =   150
         TabIndex        =   5
         Top             =   3090
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.ComboBox cboGroup 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7485
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   135
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CheckBox chkList 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   195
         Width           =   195
      End
      Begin VB.ListBox lstItens 
         Columns         =   2
         Height          =   2760
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   465
         Width           =   7770
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecionar Itens:"
         Height          =   195
         Left            =   330
         TabIndex        =   2
         Top             =   195
         Width           =   1215
      End
   End
End
Attribute VB_Name = "ideDataListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private pv_bCheckList   As Boolean
Private pv_sInCodigos   As String
Private pv_oRS          As ADODB.RecordSet

'Event Declarations:
Public Event Change()
Public Event GroupClick()
Public Event ItemCheck(Item As Integer)

Private m_strCampoCodigo            As String
Private m_strCampoDescricao         As String
Private m_strCampoDescricaoArray()  As String
Private m_strCampoDescricaoFormat   As String
Private m_strCampoAgrupamento       As String

Public Enum eIDE_TipoGroup
  tpString = 0
  tpInteger = 1
End Enum
Private m_eCampoAgrupamentoTipo As eIDE_TipoGroup

Private Function TemSelecionado() As Boolean
  TemSelecionado = lstItens.SelCount > 0
End Function

Private Sub cboGroup_Click()
  Call MontarLista(cboGroup.Text)

  RaiseEvent GroupClick
End Sub

Private Sub chkList_Click()
  Dim n As Integer

  If chkList.Value = vbGrayed Then Exit Sub

  pv_bCheckList = True
  For n = 0 To lstItens.ListCount - 1
    lstItens.Selected(n) = chkList.Value
  Next
  pv_bCheckList = False
  RaiseEvent Change
End Sub

Private Sub lstItens_ItemCheck(Item As Integer)
  If pv_bCheckList Then Exit Sub

  If lstItens.ListIndex + 1 <> lstItens.ListCount Then
    lstItens.ListIndex = lstItens.ListIndex + 1
  End If
  chkList.Value = IIf(TemSelecionado, vbGrayed, vbUnchecked)

  RaiseEvent Change
  RaiseEvent ItemCheck(Item)
End Sub


Private Sub UserControl_Resize()
  On Local Error Resume Next
  fraBorda.Move 0, -90, UserControl.ScaleWidth + 5, UserControl.ScaleHeight + 90

  ProgressBar.Top = fraBorda.Height - ProgressBar.Height
  ProgressBar.Left = 0
  ProgressBar.Width = fraBorda.Width - 15

  lstItens.Width = fraBorda.Width - 115
  lstItens.Height = fraBorda.Height - 400

  If cboGroup.Width > fraBorda.Width / 2 Then
    cboGroup.Width = fraBorda.Width / 2
  End If
  cboGroup.Left = (lstItens.Width - cboGroup.Width) + 50
  DoEvents
End Sub

Public Property Get RecordSet() As ADODB.RecordSet
  Set RecordSet = pv_oRS
End Property

Public Property Set RecordSet(ByVal adoRS As ADODB.RecordSet)
  Call Carregar(adoRS)
End Property

Private Sub Carregar(ByVal pRecordset As ADODB.RecordSet)
  On Error GoTo TrataErro:
  Dim bExitSub As Boolean

  Dim sUF As String
  Dim nL1 As Integer
  Dim nL2 As Integer

  Screen.MousePointer = vbHourglass

  Set pv_oRS = Nothing
  Set pv_oRS = pRecordset
  lstItens.Clear

  Select Case True
  Case pv_oRS Is Nothing: bExitSub = True
  Case pv_oRS.RecordCount = 0: bExitSub = True
  End Select

  If bExitSub Then
    cboGroup.Clear
    cboGroup.Visible = False

    lstItens.Clear
    lstItens.AddItem "Nenhum registro encontrado para seleção!"
    lstItens.Enabled = False
    chkList.Enabled = False
  Else
    Call MontarAgrupamento
    Call MontarLista("")

    cboGroup.Visible = cboGroup.ListCount > 1
    DoEvents
  End If

Sair:
  Screen.MousePointer = vbDefault
  Exit Sub
  Resume

TrataErro:
  Set pv_oRS = Nothing
  Err.Raise Err.Number, Err.Source, Err.Description & vbCrLf & vbTab & "Step: [ideDataListBox.Carregar]"
  GoTo Sair
End Sub

Private Sub MontarAgrupamento()
  On Error GoTo TrataErro
  Dim sArray() As String
  Dim sFlag As String
  Dim iWidth As Integer

  Screen.MousePointer = vbHourglass

  If Trim(m_strCampoAgrupamento) = "" Then GoTo Sair

  cboGroup.Enabled = False

  cboGroup.Clear
  cboGroup.AddItem "  "

  ReDim sArray(0)
  sArray(0) = ""
  With pv_oRS
    .MoveFirst
    While Not .EOF
      sFlag = Trim(.Fields(m_strCampoAgrupamento))
      If InStr(Join(sArray), UCase(sFlag)) = 0 Then
        ReDim Preserve sArray(UBound(sArray) + 1)
        sArray(UBound(sArray)) = UCase(sFlag)

        If iWidth < TextWidth(sFlag) Then
          iWidth = TextWidth(sFlag)
        End If

        cboGroup.AddItem sFlag
      End If
      .MoveNext
      DoEvents
    Wend
    .MoveFirst
  End With

  cboGroup.Width = iWidth + 400  '400 tamanho do botao
  cboGroup.Enabled = True

  UserControl_Resize
  On Error GoTo 0
Sair:
  Screen.MousePointer = vbDefault
  Exit Sub

TrataErro:
  Err.Raise Err.Number, Err.Source, Err.Description & vbCrLf & vbTab & "Step: [ideDataListBox.MontarAgrupamento]"
  GoTo Sair:
End Sub

Private Sub MontarLista(Optional ByVal pAgrupar As String = "")
  On Error GoTo TrataErro
  Dim sCodi As String
  Dim sDesc As String
  Dim sArrayCampos() As String
  Dim sArrayFormat() As String
  Dim i As Integer
  Dim iCampos As Integer
  Dim iFormat As Integer

  Screen.MousePointer = vbHourglass

  If Len(Trim(m_strCampoDescricao)) > 0 Then
    sArrayCampos = Split(Trim(m_strCampoDescricao), ";")
    iCampos = UBound(sArrayCampos) + 1

    If Len(Trim(m_strCampoDescricaoFormat)) > 0 Then
      sArrayFormat = Split(Trim(m_strCampoDescricaoFormat), ";")
      iFormat = UBound(sArrayFormat) + 1
    End If
  End If

  lstItens.Visible = False
  ProgressBar.Visible = True

  lstItens.Clear
  chkList.Enabled = False
  chkList.Value = vbUnchecked

  With pv_oRS
    If m_strCampoAgrupamento <> "" Then
      If Trim(pAgrupar) = "" Then
        .Filter = adFilterNone
      Else
        If m_eCampoAgrupamentoTipo = tpString Then
          .Filter = m_strCampoAgrupamento & " = '" & pAgrupar & "'"
        Else
          .Filter = m_strCampoAgrupamento & " = " & pAgrupar
        End If
      End If
    End If

    ProgressBar.Max = .RecordCount
    .MoveFirst
    While Not .EOF
      If m_strCampoCodigo <> "" Then
        sCodi = .Fields(m_strCampoCodigo)
      Else
        sCodi = .Fields(0)
      End If

      sDesc = ""
      If iCampos > 0 Then
        For i = 0 To iCampos - 1
          If iFormat > 0 Then
            If iFormat - 1 >= i Then
              Select Case Mid(UCase(sArrayFormat(i)), 1, 1)
              Case "L"
                sDesc = sDesc & LCase(.Fields(m_strCampoDescricaoArray(i))) & " - "
              Case "U"
                sDesc = sDesc & UCase(.Fields(m_strCampoDescricaoArray(i))) & " - "
              Case "P"
                sDesc = sDesc & ProperCase(.Fields(m_strCampoDescricaoArray(i))) & " - "
              Case Else  'Case 0 To 9
                sDesc = sDesc & Format(.Fields(m_strCampoDescricaoArray(i)), sArrayFormat(i)) & " - "
                'Case Else
                '   sDesc = sDesc & .Fields(m_strCampoDescricaoArray(i)) & " - "
              End Select
            Else
              sDesc = sDesc & .Fields(m_strCampoDescricaoArray(i)) & " - "
            End If
          Else
            sDesc = sDesc & .Fields(m_strCampoDescricaoArray(i)) & " - "
          End If
        Next
        sDesc = Trim(Mid(sDesc, 1, Len(sDesc) - 3))

        'Se a quantidade de paramentros ultrapassar a quantidade de campos na descricao
        'entende se que este é o tamanho limite da descricao
        If iFormat > iCampos Then
          If IsNumeric(sArrayFormat(iCampos)) Then
            sDesc = Mid(sDesc, 1, sArrayFormat(iCampos))
          End If
        End If
      Else
        'Padrao campo codigo(0) e descricao(1)
        sDesc = sCodi & " - " & .Fields(1)
      End If

      lstItens.AddItem sDesc
      lstItens.ItemData(lstItens.NewIndex) = Val(sCodi)

      'Selecionando os codigos padrão
      If InStr(pv_sInCodigos, "," & CStr(sCodi) & ",") Then
        lstItens.Selected(lstItens.NewIndex) = True
      End If

      ProgressBar.Value = .AbsolutePosition
      .MoveNext
      DoEvents
    Wend
  End With

  lstItens.Enabled = True
  chkList.Enabled = True
  On Error GoTo 0
Sair:
  lstItens.Visible = True
  ProgressBar.Visible = False
  ProgressBar.Value = 0
  Screen.MousePointer = vbDefault
  Exit Sub
  Resume
TrataErro:
  Err.Raise Err.Number, Err.Source, Err.Description & vbCrLf & vbTab & "Step: [ideDataListBox.MontarLista]"
  GoTo Sair:
End Sub

Public Function Selecionados(Optional pFormat As String = "0000", Optional pUseAspas As Boolean) As String
  On Error GoTo TrataErro
  Dim sCod As String
  Dim n As Integer

  If Len(pFormat) > 4 Then pFormat = Right(pFormat, 4)

  For n = 0 To lstItens.ListCount - 1
    If lstItens.Selected(n) = True Then
      sCod = sCod & Right(lstItens.ItemData(n), Len(pFormat)) & ","
    End If
  Next

  sCod = Mid(sCod, 1, Len(sCod) - 1)
  If pUseAspas Then
    sCod = "'" & Replace(sCod, ",", "','") & "'"
  End If

  Selecionados = sCod

TrataErro:
End Function

Public Property Let SelecaoPadrao(ByVal pInCodigos As String)
  pInCodigos = Trim(pInCodigos)
  If Left(pInCodigos, 1) <> "," Then
    pInCodigos = "," & pInCodigos
  End If
  If Right(pInCodigos, 1) <> "," Then
    pInCodigos = pInCodigos & ","
  End If
  pv_sInCodigos = pInCodigos
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fraBorda,fraBorda,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = fraBorda.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  fraBorda.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  fraBorda.BackColor = PropBag.ReadProperty("BackColor", &HC0C0C0)
  lblCaption.Caption = PropBag.ReadProperty("Caption", "Selecionar Itens.")
  m_strCampoCodigo = PropBag.ReadProperty("CampoCodigo", tpString)
  m_strCampoDescricao = PropBag.ReadProperty("CampoDescricao", tpString)
  m_strCampoDescricaoArray = Split(m_strCampoDescricao, ";")
  m_strCampoDescricaoFormat = PropBag.ReadProperty("CampoDescricaoFormat", tpString)
  m_strCampoAgrupamento = PropBag.ReadProperty("CampoAgrupamento", tpString)
  m_eCampoAgrupamentoTipo = PropBag.ReadProperty("CampoAgrupamentoTipo", tpString)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  lstItens.Columns = PropBag.ReadProperty("Columns", 2)
  Set lblCaption.Font = PropBag.ReadProperty("CaptionFont", Ambient.Font)
  Set lstItens.Font = PropBag.ReadProperty("ListFont", Ambient.Font)
End Sub

Private Sub UserControl_Terminate()
  Set pv_oRS = Nothing
  pv_bCheckList = False
  pv_sInCodigos = ""

  lstItens.Clear
  cboGroup.Clear
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", fraBorda.BackColor, &HC0C0C0)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Selecionar Itens.")
  Call PropBag.WriteProperty("CampoCodigo", m_strCampoCodigo, tpString)
  Call PropBag.WriteProperty("CampoDescricao", m_strCampoDescricao, tpString)
  Call PropBag.WriteProperty("CampoDescricaoFormat", m_strCampoDescricaoFormat, tpString)
  Call PropBag.WriteProperty("CampoAgrupamento", m_strCampoAgrupamento, tpString)
  Call PropBag.WriteProperty("CampoAgrupamentoTipo", m_eCampoAgrupamentoTipo, tpString)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Columns", lstItens.Columns, 2)
  Call PropBag.WriteProperty("CaptionFont", lblCaption.Font, Ambient.Font)
  Call PropBag.WriteProperty("ListFont", lstItens.Font, Ambient.Font)
End Sub

Public Property Get Caption() As String
  Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal strCaption As String)
  lblCaption.Caption = strCaption
  Call UserControl.PropertyChanged("Caption")
End Property

Public Property Get CampoCodigo() As String
  CampoCodigo = m_strCampoCodigo
End Property

Public Property Let CampoCodigo(ByVal pCampo As String)
  m_strCampoCodigo = pCampo
  Call UserControl.PropertyChanged("CampoCodigo")
End Property

Public Property Get CampoDescricao() As String
  CampoDescricao = m_strCampoDescricao
End Property

Public Property Let CampoDescricao(ByVal pCamposDelimiter As String)
  m_strCampoDescricao = pCamposDelimiter
  m_strCampoDescricaoArray() = Split(pCamposDelimiter, ";")

  Call UserControl.PropertyChanged("CampoDescricao")
End Property

Public Property Get CampoDescricaoFormat() As String
  CampoDescricaoFormat = m_strCampoDescricaoFormat
End Property

Public Property Let CampoDescricaoFormat(ByVal NewValue As String)
  m_strCampoDescricaoFormat = NewValue
  Call UserControl.PropertyChanged("CampoDescricaoFormat")
End Property

Public Property Get CampoAgrupamento() As String
  CampoAgrupamento = m_strCampoAgrupamento
End Property

Public Property Let CampoAgrupamento(ByVal pCampo As String)
  m_strCampoAgrupamento = pCampo
  Call UserControl.PropertyChanged("CampoAgrupamento")
End Property

Public Property Get CampoAgrupamentoTipo() As eIDE_TipoGroup
  CampoAgrupamentoTipo = m_eCampoAgrupamentoTipo
End Property

Public Property Let CampoAgrupamentoTipo(ByVal eCampoAgrupamentoTipo As eIDE_TipoGroup)
  m_eCampoAgrupamentoTipo = eCampoAgrupamentoTipo
  Call UserControl.PropertyChanged("CampoAgrupamentoTipo")
End Property

Private Function ProperCase(ByVal pTexto As String) As String
  pTexto = StrConv(pTexto, vbProperCase)

  pTexto = Replace(pTexto, " Da ", " da ")
  pTexto = Replace(pTexto, " Das ", " das ")
  pTexto = Replace(pTexto, " De ", " de ")
  pTexto = Replace(pTexto, " Do ", " do ")
  pTexto = Replace(pTexto, " Dos ", " dos ")

  ProperCase = pTexto
End Function

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  UserControl.Enabled = NewValue
  chkList.Enabled = NewValue
  lstItens.Enabled = NewValue
  cboGroup.Enabled = NewValue

  Call UserControl.PropertyChanged("Enabled")
End Property

Public Property Get Columns() As Integer
  Columns = lstItens.Columns
End Property

Public Property Let Columns(ByVal NewValue As Integer)
  If NewValue < 0 Then NewValue = 1

  lstItens.Columns = NewValue
  Call UserControl.PropertyChanged("Columns")
End Property

Public Property Get SelCount() As Integer
  SelCount = lstItens.SelCount
End Property

Public Property Get CaptionFont() As Font
  Set CaptionFont = lblCaption.Font
End Property

Public Property Set CaptionFont(ByVal New_Font As Font)
  On Local Error Resume Next
  Set lblCaption.Font = New_Font

  UserControl.PropertyChanged "CaptionFont"
End Property

Public Property Get ListFont() As Font
  Set ListFont = lstItens.Font
End Property

Public Property Set ListFont(ByVal New_Font As Font)
  On Local Error Resume Next
  Set lstItens.Font = New_Font

  UserControl.PropertyChanged "ListFont"
End Property

Public Sub NaoSelecionar(Optional ByVal sInCodigos As String = "Todos")
  Dim i As Integer

  If LCase(Trim(sInCodigos)) = "todos" Then
    For i = 0 To lstItens.ListCount - 1
      lstItens.Selected(i) = False
    Next
  Else
    For i = 0 To lstItens.ListCount - 1
      If InStr(sInCodigos, lstItens.ItemData(i)) > 0 Then
        lstItens.Selected(i) = False
      End If
    Next
  End If
End Sub

Public Sub Selecionar(Optional ByVal sInCodigos As String = "Todos")
  Dim i As Integer

  If LCase(Trim(sInCodigos)) = "todos" Then
    For i = 0 To lstItens.ListCount - 1
      lstItens.Selected(i) = True
    Next
  Else
    For i = 0 To lstItens.ListCount - 1
      If InStr(sInCodigos, lstItens.ItemData(i)) > 0 Then
        lstItens.Selected(i) = True
      End If
    Next
  End If
End Sub

Public Property Get Text() As String
  Text = lstItens.Text
End Property

Public Property Get ListIndex() As Long
  ListIndex = lstItens.ListIndex
End Property

Public Property Let ListIndex(ByVal NewValue As Long)
  lstItens.ListIndex = NewValue
End Property

Public Property Get ListCount() As Long
  ListCount = lstItens.ListCount
End Property

Public Property Get ItemData(Index) As Long
  ItemData = lstItens.ItemData(Index)
End Property

Public Property Get List(Index) As String
  List = lstItens.List(Index)
End Property

Public Property Get Selected(Index) As Boolean
  Selected = lstItens.Selected(Index)
End Property

Public Property Get ValorCodigo(Index) As Long
    ValorCodigo = lstItens.ItemData(Index)
End Property
