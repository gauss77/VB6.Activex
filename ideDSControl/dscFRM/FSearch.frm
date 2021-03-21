VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Begin VB.Form FSearch 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2175
   ClientLeft      =   1950
   ClientTop       =   2100
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin Insignia_Frame.ideFrame ideFrame 
      Align           =   1  'Align Top
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   609
      BorderExt       =   6
      BorderWidth     =   23
      BackColor       =   16777215
      BackColorB      =   14737632
      GradientStyle   =   4
      Caption         =   "Janela de Pesquisa..."
      CaptionAlign    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdButtons 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5265
         Style           =   1  'Graphical
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   75
         Picture         =   "FSearch.frx":058A
         Top             =   30
         Width           =   240
      End
   End
   Begin Insignia_Frame.ideFrame ideFrame 
      Align           =   1  'Align Top
      Height          =   1500
      Index           =   1
      Left            =   0
      Top             =   345
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   2646
      BorderExt       =   6
      BorderInt       =   6
      BorderPaint     =   10
      BorderWidth     =   20
      BackColor       =   14737632
      BackColorB      =   16777215
      GradientStyle   =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbCampos 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   315
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   1935
      End
      Begin AButtons.AButton abtButtons 
         Height          =   390
         Left            =   4080
         TabIndex        =   1
         Top             =   1020
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   688
         BTYPE           =   4
         TX              =   "&Pesquisar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   345
         Left            =   2340
         TabIndex        =   2
         ToolTipText     =   "Para pesquisar sobrenome digite primeiro sinal de porcentagem (%)"
         Top             =   585
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextCase        =   1
         RawText         =   0
         FontName        =   "Century Gothic"
         FontSize        =   8,25
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pesquizar por:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   5
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor de Pesquisa"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   2340
         TabIndex        =   4
         Top             =   330
         Width           =   1425
      End
   End
   Begin Insignia_Frame.ideFrame ideFrame 
      Align           =   2  'Align Bottom
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   1875
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   529
      BorderExt       =   6
      BorderWidth     =   5
      BackColor       =   16777215
      Caption         =   "Contato: codeuapp@gmail.com"
      ForeColor       =   10526880
      CaptionAlign    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "FSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'########################################################################################
'# Exemplo de psCapFieldMask =
'#  "Caption       ,Field    ,Maskara    |Caption     ,Field    ,Maskara  "
'#  "Data do Pedido,DATA_PEDI,!dd/mm/yyyy|N∫ do Pedido,NUME_PEDI,999999;0;"
'#  {OBS: as mascaras:
'#   '–dd/mm/yyyy' = inicial com – significa que a pesquisa utilizara os campos de data
'#   'ÒP'          = significa que o campo n„o deve entrar no combo de pesquisa
'#
'#  Alt 164 = Ò  /  Alt 209 = –
'########################################################################################

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private mDS As CDSControl

Private Const cPNLCaption As Byte = 0

Private Const cBTNSearch   As Byte = 1
Private Const cBTNFechar  As Byte = 0

Private maFields()      As String
Private maMasks()       As String

Public Sub MontarTela(ByVal psAliasFieldMask As String)

  Dim sCapt As String, sField As String, sMask As String
  Dim aI() As String, aL() As String

  aI = Split(psAliasFieldMask, "|")

  Dim i As Byte
  For i = 0 To UBound(aI)
    aL = Split(aI(i), ",")

    sCapt = Trim$(aL(0))
    sField = Trim$(aL(1))
    sMask = LTrim$(aL(2))
 
    'Se <> ent„o adicionando na Combo os Campos
    With cmbCampos
      .AddItem sCapt
      ReDim Preserve maFields(.ListCount)
      ReDim Preserve maMasks(.ListCount)
      maFields(.ListCount) = sField
      maMasks(.ListCount) = sMask
    End With
  Next
End Sub

Public Sub ShowForm(ByRef DataSource As CDSControl, _
                    Optional ByVal psAliasFieldMask As String, _
                    Optional ByVal psCaptionTela As String = "Janela de Pesquisa...")

  If DataSource Is Nothing Then
    Unload Me
    Exit Sub
  End If
  Set mDS = DataSource
 
  On Error GoTo TrataErro
  If psAliasFieldMask <> "" Then Call MontarTela(psAliasFieldMask)
  
  If cmbCampos.ListCount = 0 Then
    Unload Me
    Exit Sub
  End If
  
  cmbCampos.ListIndex = 0
  
  ideFrame(cPNLCaption).Caption = Space(3) & LTrim$(psCaptionTela)
  Me.Show vbModal
  
  Me.Hide
  Exit Sub
  
TrataErro:
  MsgBox Err.Source & ":(" & Err.Description & ")", vbCritical, "FormPesquisa.ShowForm"
  Unload Me
End Sub

Private Sub abtButtons_Click()
    cmdButtons_Click cBTNSearch
End Sub

Private Sub cmbCampos_Click()
  Dim sMask As String
  
  If cmbCampos.ListIndex <> -1 Then
    sMask = maMasks(cmbCampos.ListIndex + 1)
    
    txtCampo.Mask = sMask
    If sMask = "" Then txtCampo.MaxLength = 0
    
    On Error Resume Next
    txtCampo.SetFocus
    On Error GoTo 0
  End If
End Sub

Private Sub cmdButtons_Click(Index As Integer)
  Select Case Index
    Case Is = cBTNFechar
      On Error Resume Next
      SaveSetting App.EXEName, "Campo Pesquisa", gsParent, cmbCampos.ListIndex
      SaveSetting App.EXEName, "Valor Pesquisa", gsParent, txtCampo.Text
      On Error GoTo 0
      
      Me.Hide
    Case Is = cBTNSearch
    
      Dim sValor As String, sMask As String
    
      sValor = txtCampo.Text
      sMask = txtCampo.Mask
        
      Select Case sMask
        Case Is = "##/##/####"
          sMask = "dd/mm/yyyy"
          sValor = Format(sValor, sMask)
        
        Case Is = "##/##"
          sMask = "dd/mm"
          sValor = Format(sValor, sMask)
        
        Case Is = "##/####"
          sMask = "mm/yyyy"
          sValor = Format(sValor, sMask)
      End Select
      
      Call mDS.Search(maFields(cmbCampos.ListIndex + 1), sValor)
  End Select

End Sub

Private Sub Form_Load()
  Call GetSettings
End Sub

Private Sub ideFrame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Index = cPNLCaption Then Call DragForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set mDS = Nothing
  Set FSearch = Nothing
End Sub

Private Sub txtCampo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then cmdButtons_Click cBTNSearch
End Sub

Private Sub DragForm()
  On Local Error Resume Next
  'Move the borderless form...
  Call ReleaseCapture
  Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub GetSettings()
  ' Ler os valores dos campo de indexaÁ„o e Pesquisa usado pela ultima vezes
  With cmbCampos
    If .ListCount > 0 Then
      On Error GoTo ErrCombo
      .ListIndex = GetSetting(App.EXEName, "Campo Pesquisa", gsParent, "0")
      On Error GoTo 0
    Else
ErrCombo:
      Call SaveSetting(App.EXEName, "Campo Pesquisa", gsParent, 0)
      Err.Clear
    End If
  End With
  
  txtCampo.Text = GetSetting(App.EXEName, "Valor Pesquisa", gsParent, "")
End Sub


