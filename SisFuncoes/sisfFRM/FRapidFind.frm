VERSION 5.00
Object = "{EADE62FD-5B6B-444E-A6C6-26CFE520CF78}#1.0#0"; "ideToolBar.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Begin VB.Form FRapidFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pesquisa Rápida..."
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6195
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
   ScaleHeight     =   3045
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Insignia_Frame.ideFrame panel 
      Align           =   1  'Align Top
      Height          =   675
      Index           =   1
      Left            =   0
      Top             =   1935
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbOrdem 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   2295
      End
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   255
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontSize        =   8,25
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Localizar por:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   45
         Width           =   960
      End
   End
   Begin Insignia_Frame.ideFrame panel 
      Align           =   1  'Align Top
      Height          =   1935
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1800
         Left            =   75
         TabIndex        =   0
         Top             =   60
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3175
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Caption         =   "Listagem de Dados"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin Insignia_Toolbar.ideToolbar TBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   2670
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   661
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
      ButtonCount     =   12
      ButtonStyle1    =   2
      ButtonStyle2    =   2
      ButtonKey3      =   "NavPri"
      ButtonPicture3  =   "FRapidFind.frx":0000
      ButtonToolTipText3=   "Move para o primeiro registro"
      ButtonKey4      =   "NavAnt"
      ButtonPicture4  =   "FRapidFind.frx":0352
      ButtonToolTipText4=   "Move para o registro anterior"
      ButtonKey5      =   "NavPro"
      ButtonPicture5  =   "FRapidFind.frx":06A4
      ButtonToolTipText5=   "Move para o proximo registro"
      ButtonKey6      =   "NavUlt"
      ButtonPicture6  =   "FRapidFind.frx":09F6
      ButtonToolTipText6=   "Move para o último registro"
      ButtonStyle7    =   2
      ButtonCaption8  =   "Con&firmar  "
      ButtonKey8      =   "Confirmar"
      ButtonPicture8  =   "FRapidFind.frx":0D48
      ButtonToolTipText8=   "Confirma seleção      [ENTER]"
      ButtonCaption9  =   "&Cancelar  "
      ButtonKey9      =   "Cancelar"
      ButtonPicture9  =   "FRapidFind.frx":109A
      ButtonToolTipText9=   "Cancela seleção     [ESC]"
      ButtonStyle10   =   2
      ButtonCaption11 =   "&Incluir  "
      ButtonKey11     =   "Incluir"
      ButtonPicture11 =   "FRapidFind.frx":13EC
      ButtonToolTipText11=   "Abre janela de Cadastro    [INS]"
      ButtonCaption12 =   "&Editar  "
      ButtonKey12     =   "Editar"
      ButtonPicture12 =   "FRapidFind.frx":173E
      ButtonToolTipText12=   "Altera o registro selecionado    [Ctrl + INS]"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu menuPopUp 
         Caption         =   ""
         Index           =   1
      End
   End
End
Attribute VB_Name = "FRapidFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gOConn        As ADODB.Connection
Private mRS           As ADODB.Recordset
Attribute mRS.VB_VarHelpID = -1
Private mFormAddNew   As Form

Public Event Valores(ID As Long, Nome As String, Index As Integer)
Public Event ValColunas(Col As Columns)

Private msSQLPrimary    As String
Private msCapWidthGrid  As String
Private maDataField()   As String
Private mIDXRetorno     As Integer

Private Sub Destruir()
  On Error Resume Next
  mRS.Close
  Set mRS = Nothing
  On Error GoTo 0
  Set mFormAddNew = Nothing
  Set FRapidFind = Nothing
End Sub

Public Function Abrir(ByVal SQL As String, ByRef FormAddNew As Form, IdxCTRRetorno As Integer)
  mIDXRetorno = IdxCTRRetorno
  If Not FormAddNew Is Nothing Then
    Set mFormAddNew = FormAddNew
  Else
    TBar.ButtonEnabled("Incluir") = False
    TBar.ButtonEnabled("Editar") = False
  End If
  
  msSQLPrimary = SQL
  If SetRecordSet Then
    Me.Show vbModal
  Else
    Unload Me
  End If
End Function

Public Sub EnabledButtons(pConfirmar As Boolean, _
                          pCancelar As Boolean, _
                          pIncluir As Boolean, _
                          pEditar As Boolean)
  TBar.ButtonEnabled("Confirmar") = pConfirmar
  TBar.ButtonEnabled("Cancelar") = pCancelar
  TBar.ButtonEnabled("Incluir") = pIncluir
  TBar.ButtonEnabled("Editar") = pEditar
End Sub

Public Property Get Connection() As ADODB.Connection
  Set Connection = gOConn
End Property

Public Property Set Connection(vNewValue As ADODB.Connection)
  Set gOConn = vNewValue
End Property

Public Sub DestroyConnection()
  On Error Resume Next
  gOConn.Close
  Set gOConn = Nothing
End Sub

Private Function SetRecordSet() As Boolean
  Set mRS = Nothing
  Set mRS = CriarRSCliente(msSQLPrimary)
  
  If Not mRS Is Nothing Then
    Call SetarCaptions
    DataGrid1.BorderStyle = dbgNoBorder
    SetRecordSet = True
  End If
End Function

Private Sub SetarCaptions()
  Dim aCols() As String
  Dim aCapW() As String
  Dim i As Integer
  
  aCols = Split(msCapWidthGrid, "|")
  cmbOrdem.Clear
  Set DataGrid1.DataSource = Nothing
  Set DataGrid1.DataSource = mRS

  For i = 0 To UBound(aCols)
    aCapW = Split(aCols(i), ",")
    DataGrid1.Columns(i).Caption = aCapW(0)
    DataGrid1.Columns(i).Width = aCapW(1)
    ReDim Preserve maDataField(i)
    maDataField(i) = DataGrid1.Columns(i).DataField
    cmbOrdem.AddItem aCapW(0)
    cmbOrdem.ListIndex = 0
  Next
End Sub

Private Sub cmbOrdem_Click()
  mRS.Sort = maDataField(cmbOrdem.ListIndex)
  txtCampo.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case Is = vbKeyEscape
      TBar_ButtonClick 0, "Cancelar"
      KeyCode = 0
    Case Is = vbKeyReturn
      TBar_ButtonClick 0, "Confirmar"
      KeyCode = 0
    Case Is = vbKeyInsert
      If Not mFormAddNew Is Nothing Then
        If Shift = vbCtrlMask Then
          TBar_ButtonClick 0, "Editar"
        Else
          TBar_ButtonClick 0, "Incluir"
        End If
      End If
      KeyCode = 0
  End Select
End Sub

Private Sub TBar_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
  Select Case ButtonKey
    Case Is = "NavPri"
      If Not mRS.BOF Or Not mRS.EOF Then mRS.MoveFirst
    Case Is = "NavAnt"
      If mRS.RecordCount > 1 And mRS.AbsolutePosition > 1 Then mRS.MovePrevious
    Case Is = "NavPro"
      If mRS.RecordCount > 1 And mRS.AbsolutePosition < mRS.RecordCount Then mRS.MoveNext
    Case Is = "NavUlt"
      If mRS.RecordCount > 1 And mRS.AbsolutePosition < mRS.RecordCount Then mRS.MoveLast
      
    Case Is = "Incluir"
      Call Incluir
    Case Is = "Editar"
      Call Editar
      
    Case Is = "Confirmar"
      On Error GoTo ErrCancel
      With DataGrid1.Columns
        RaiseEvent Valores(.Item(0).Text, .Item(1).Text, mIDXRetorno)
        RaiseEvent ValColunas(DataGrid1.Columns)
      End With
      Unload Me
    Case Is = "Cancelar"
      Unload Me

  End Select
ErrCancel:
End Sub

Public Property Let CapWidthGrid(ByVal vNewValue As String)
  msCapWidthGrid = vNewValue
End Property

Private Sub Incluir()
  Dim lID As Long
  
  lID = mFormAddNew.Abrir
  If lID <> 0 Then
    RaiseEvent Valores(lID, "", mIDXRetorno)
    Unload Me
  End If
End Sub

Private Sub Editar()
  Dim lID As Long
  
  If mRS.RecordCount = 0 Then Exit Sub
  On Error GoTo Sair:

  lID = mFormAddNew.Abrir(mRS!ID)
  If lID <> 0 Then
    mRS.Requery
    Call SetarCaptions
    Call Posicionar(mRS, lID)
  End If
Sair:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call Destruir
End Sub

Private Sub txtCampo_KeyPress(KeyAscii As Integer)
  Call AutoPosition(mRS, maDataField(cmbOrdem.ListIndex), txtCampo.Text, KeyAscii)
End Sub

Private Function CriarRSCliente(ByVal pSQL As String, _
                               Optional ByVal pDesconectado As Boolean) As ADODB.Recordset
  Dim RS As ADODB.Recordset
  
  Set RS = New ADODB.Recordset
  
  On Error GoTo TrataErro
  With RS
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .CursorType = adOpenDynamic
    Set .ActiveConnection = gOConn
    .Open pSQL

    If pDesconectado Then Set .ActiveConnection = Nothing
  End With
  
  Set CriarRSCliente = RS
  Set RS = Nothing
  Exit Function
  
TrataErro:
  Set RS = Nothing
  MsgBox Err.Description & vbCrLf & Err.Source, vbCritical, "Módulo ADO - Criação de RSCliente"
  Err.Clear
End Function

Public Function CriarRSServer(pSQL As String) As ADODB.Recordset
  Dim RS As ADODB.Recordset
  
  Set RS = New ADODB.Recordset
  
  On Error GoTo TrataErro
  Set RS = gOConn.Execute(pSQL)
  If RS.RecordCount = 0 Then
    RS.Close
    Set RS = Nothing
  End If
  
  Set CriarRSServer = RS
  Set RS = Nothing
  Exit Function
TrataErro:
  Set RS = Nothing
  MsgBox Err.Description & vbCrLf & Err.Source, vbCritical, "Módulo ADO - Criação de RSServer"
  Err.Clear
End Function

Private Function Posicionar(ByRef pObjRecordSet As Object, ByVal pID As Long) As Boolean
   Dim nPos As Long
   
   If pObjRecordSet.RecordCount <> 0 Then
      nPos = pObjRecordSet.AbsolutePosition
      pObjRecordSet.Find "ID = " & pID, , adSearchForward, 1
      
      Posicionar = Not pObjRecordSet.EOF
      If pObjRecordSet.EOF Then
         pObjRecordSet.AbsolutePosition = nPos
      End If
   End If
End Function

Public Function AutoPosition(ByRef pRecordset As Object, ByVal DataField As String, _
                             ByVal Texto As String, KeyAscii As Integer) As Long
  Dim sTexto      As String
  Dim sCriterio   As String
  Dim nPos        As Long
  Dim RSConsulta  As ADODB.Recordset
  Dim sSQL        As String
  
  If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then Exit Function

  On Error Resume Next
  If KeyAscii = vbKeyBack Then
    sTexto = Mid$(Texto, 1, Len(Texto) - 1)
  Else
    sTexto = Texto & Chr(KeyAscii)
  End If
  On Error GoTo 0
  
  If sTexto = "" Then Exit Function
  
  On Error Resume Next
  
  Set RSConsulta = CriarRSServer(pRecordset.Source)
  If RSConsulta Is Nothing Then Exit Function
  
  With RSConsulta
    Select Case RSConsulta.Fields(DataField).Type
      Case Is = adChapter, adWChar, adVarChar, adVarWChar, adChar
        sCriterio = " LIKE '" & sTexto & "%'"
        
      Case Is = adNumeric, adInteger, adSmallInt, adDouble
        sCriterio = " = " & sTexto
        
      Case Is = adBinary, adBoolean
        sCriterio = " = " & sTexto
        
      Case Is = adDate, adDBDate, adDBTime
        sCriterio = " = '" & sTexto & "'"
        
      Case Else
        sCriterio = " LIKE '" & sTexto & "%'"
    End Select
    
    DataField = "[" & DataField & "]"
    .Find DataField & sCriterio, , adSearchForward, 1
    If Err.Number <> 0 Then
      GoTo TrataErro:
    End If
  
    If Not .BOF And Not .EOF Then
      Call Posicionar(pRecordset, RSConsulta!ID)
      .Close
    End If
  End With
  Set RSConsulta = Nothing
  
  On Error GoTo 0
  Exit Function
  
TrataErro:
  MsgBox Err.Description & vbCrLf & Err.Source, vbCritical, "Auto Posicionar"
End Function


