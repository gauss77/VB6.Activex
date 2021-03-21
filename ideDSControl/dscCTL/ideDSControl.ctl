VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Begin VB.UserControl ideDSControl 
   Alignable       =   -1  'True
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11100
   DataSourceBehavior=   1  'vbDataSource
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   1110
   ScaleWidth      =   11100
   ToolboxBitmap   =   "ideDSControl.ctx":0000
   Begin Insignia_Frame.ideFrame fraBarra 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   714
      BorderExt       =   6
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox Combos 
         Height          =   315
         Index           =   0
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Ordem de exibição dos dados"
         Top             =   40
         Width           =   1875
      End
      Begin Insignia_Frame.ideFrame fraBarra 
         Height          =   345
         Index           =   2
         Left            =   4740
         Top             =   25
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   609
         BorderExt       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox chkContinuarInsert 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Continuar &incluindo."
            Height          =   195
            Left            =   2925
            TabIndex        =   10
            ToolTipText     =   "Se marcado continua incluido após a confirmação"
            Top             =   75
            Width           =   1755
         End
         Begin AButtons.AButton abtBarra3 
            Height          =   315
            Index           =   0
            Left            =   45
            TabIndex        =   8
            ToolTipText     =   "Confirmar Operação Ativa - [F8]"
            Top             =   15
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "Con&firmar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14737632
            FCOL            =   0
         End
         Begin AButtons.AButton abtBarra3 
            Height          =   315
            Index           =   1
            Left            =   1485
            TabIndex        =   9
            ToolTipText     =   "Cancelar Operação Ativa - [F9]"
            Top             =   15
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "&Cancelar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14737632
            FCOL            =   0
         End
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   7
         Left            =   1140
         TabIndex        =   3
         ToolTipText     =   "[Ctrl + F5] - Atualizar registros da tabela"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   7
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":0312
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   1
         ToolTipText     =   "[F6] - Alterar registro ativo"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   7
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":06AC
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   0
         ToolTipText     =   "[F5] - Incluir novo registro"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   7
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":0A46
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   2
         Left            =   780
         TabIndex        =   2
         ToolTipText     =   "[F7] - Excluir registro ativo"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   7
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":0DE0
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   3
         Left            =   3810
         TabIndex        =   6
         ToolTipText     =   "[F3] - Pesquisar no Banco de Dados"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":117A
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   4
         Left            =   4170
         TabIndex        =   7
         ToolTipText     =   "Imprimir Relatórios"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":1714
      End
      Begin AButtons.AButton abtBarra1 
         Height          =   315
         Index           =   5
         Left            =   3405
         TabIndex        =   5
         ToolTipText     =   "Alterna entre ordenação ascendente e descendente"
         Top             =   40
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":1CAE
      End
   End
   Begin Insignia_Frame.ideFrame fraBarra 
      Align           =   1  'Align Top
      Height          =   405
      Index           =   1
      Left            =   0
      Top             =   405
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   714
      BorderExt       =   6
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin AButtons.AButton abtBarra2 
         Height          =   315
         Index           =   3
         Left            =   1140
         TabIndex        =   14
         ToolTipText     =   "[End] - Movimenta para o último registro"
         Top             =   30
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
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
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":2048
      End
      Begin AButtons.AButton abtBarra2 
         Height          =   315
         Index           =   2
         Left            =   780
         TabIndex        =   13
         ToolTipText     =   "[PageDown] - Movimenta para o próximo registro"
         Top             =   30
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
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
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":25E2
      End
      Begin AButtons.AButton abtBarra2 
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   12
         ToolTipText     =   " [PageUp] - Movimenta para o registro anterior"
         Top             =   30
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
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
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":2B7C
      End
      Begin AButtons.AButton abtBarra2 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   11
         ToolTipText     =   "[Home] - Movimenta para o primeiro registro"
         Top             =   30
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
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
         BCOL            =   14737632
         FCOL            =   0
         PICTURE         =   "ideDSControl.ctx":3116
      End
      Begin Insignia_Frame.ideFrame fraBarra 
         Height          =   315
         Index           =   3
         Left            =   4755
         Top             =   30
         Visible         =   0   'False
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   556
         BorderExt       =   0
         BackColor       =   12632256
         Caption         =   "Pesquisar"
         CaptionAlign    =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combos 
            Height          =   315
            Index           =   1
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   -15
            Width           =   1875
         End
         Begin rdActiveText.ActiveText txtPesquisa 
            Height          =   315
            Left            =   2760
            TabIndex        =   16
            Top             =   -15
            Width           =   2505
            _ExtentX        =   4419
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
      End
      Begin Insignia_Frame.ideFrame fraPanel 
         Height          =   255
         Index           =   0
         Left            =   1530
         Top             =   75
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   450
         BorderExt       =   6
         Caption         =   "0 / 0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Shape shpRegistro 
            BackColor       =   &H00FF8080&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            DrawMode        =   6  'Mask Pen Not
            Height          =   240
            Left            =   15
            Top             =   0
            Width           =   165
         End
      End
      Begin Insignia_Frame.ideFrame fraPanel 
         Height          =   255
         Index           =   1
         Left            =   3390
         Top             =   75
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   450
         BorderExt       =   6
         Caption         =   "Identificador"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu menuPopUp 
      Caption         =   "menuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "Ordernar"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   1
      End
   End
End
Attribute VB_Name = "ideDSControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'################################################################
'# Projeto          : Controle XDataSource                      #
'# Data de início   : 25/10/2001                                #
'#«««««««««««««««««««««««« Descrição »»»»»»»»»»»»»»»»»»»»»»»»»»»#
'# Arquivo      :                     Criado: 05/11/2001        #
'# Comentário   : Controle para Navegação e Manipulação do      #
'#                RecordSet ADO.                                #
'# Analista     : Heliomar Pereira Marques dos Santos           #
'################################################################

Option Explicit

Private WithEvents mDS    As CDSControl
Attribute mDS.VB_VarHelpID = -1

Private maFields()        As String   'Fields de Pesquisa e Ordem
Private maMasks()         As String   'Mascara do Campos de Pesquisa

Private mbCancelEvent     As Boolean

Private meOperacao        As eDSOperacao
Private meBTNExtras       As eDSBotoesExtras
Private meDSPermissoes    As eDSPermissoes
Private mbFlatButtons     As Boolean
Private mbAddContFlag     As Boolean
Private meModelo          As eDSModelo
Private meButtonType      As AButtons.ButtonTypes
Private mBtnColor         As OLE_COLOR
Private mBtnColorInative  As OLE_COLOR
Private Const cCorDesab = &HC0C0C0

Private mbSortDesc        As Boolean

'======Constantes
Private Const cCmbOrder   As Byte = 0
Private Const cCmbPesq    As Byte = 1

Private Const cPNLCont    As Byte = 0
Private Const cPNLIdent   As Byte = 1
Private Const cPNLPesq    As Byte = 3

Private Const cTBarFunc   As Byte = 0
Private Const cTBarNavi   As Byte = 1

Private Const CorBtnCheck = &HC0FFFF
'================

Public Event MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Public Event FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Public Event Operacao(ByVal eOperacao As eDSOperacao, ByVal eOperacaoAnterior As eDSOperacao)

Public Event AntesAddNew(Cancel As Boolean)
Public Event AntesCancel(Cancel As Boolean)
Public Event AntesUpdate(Cancel As Boolean, eOperacao As eDSOperacao)
Public Event DepoisUpdate(eOperacao As eDSOperacao)
Public Event AntesRequery(Cancel As Boolean)
Public Event DepoisRequery()
Public Event AntesEdit(Cancel As Boolean)
Public Event AntesDelete(Cancel As Boolean, bNotMessage As Boolean)

Public Event ClickImprimir()
Public Event ClickPesquisa(Cancel As Boolean)
Public Event ClickButtonsCreate(ByVal ButtonKey As String)

Public Sub About()
Attribute About.VB_Description = "Sobre: Heliomar P. Marques \r\ncontato: heliomarpm@hotmail.com"
Attribute About.VB_UserMemId = -552
    Debug.Print ("Insignia_DSControl.ideDSControl.About")
  FormSplash.Show vbModal
End Sub

Public Sub SetNewDS()
    Debug.Print ("Insignia_DSControl.ideDSControl.SetNewDS")
  If mDS Is Nothing Then
    Set mDS = New CDSControl
  Else
    MsgBox "DataSource já esta conectado!"
  End If
End Sub

Private Sub abtBarra1_Click(Index As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.abtBarra1_Click")
  With mDS
    Select Case Index
      Case Is = 0:  Call .AddNew
      Case Is = 1:  Call Edit
      Case Is = 2:  Call Delete
      Case Is = 3:  Call OperacaoPesquisar '(Not TBar(Index).ButtonChecked(ButtonIndex))
      Case Is = 4:  RaiseEvent ClickImprimir
      Case Is = 5:  Call OrderAscDesc
      Case Is = 7:  If meOperacao = opVisualizacao Then Call .Requery
      Case Else 'Criados pelo codigo
'        RaiseEvent ClickButtonsCreate(ButtonKey)
    End Select
  End With
End Sub

Private Sub abtBarra2_Click(Index As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.abtBarra2_Click")
  With mDS
    Select Case Index
      Case Is = 0:  Call .MoveFirst
      Case Is = 1:  Call .MovePrevious
      Case Is = 2:  Call .MoveNext
      Case Is = 3:  Call .MoveLast
    End Select
  End With
End Sub

Private Sub abtBarra3_Click(Index As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.abtBarra3_Click")
  Select Case Index
    Case Is = 0:  Call Update
    Case Is = 1:  Call mDS.Cancel
  End Select

End Sub

Private Sub chkContinuarInsert_Click()
    Debug.Print ("Insignia_DSControl.ideDSControl.chkContinuarInsert_Click")
    Call AddNewContinue
End Sub

Private Sub UserControl_InitProperties()
    Debug.Print ("Insignia_DSControl.ideDSControl.UserControl_InitProperties")
    meOperacao = opVisualizacao
    meBTNExtras = beNone
    mbFlatButtons = True
    meDSPermissoes = peTodos
    mBtnColorInative = cCorDesab
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.UserControl_KeyDown")
    MFuncoes.KeyDown Me, KeyCode, Shift
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Debug.Print ("Insignia_DSControl.ideDSControl.UserControl_ReadProperties")
  meOperacao = PropBag.ReadProperty("Operacao", 0)
  Me.Modelo = PropBag.ReadProperty("Modelo", mdMaster)
  Me.CaptionColor = PropBag.ReadProperty("CaptionColor", Ambient.ForeColor) 'vbBlack)
  Me.Permissoes = PropBag.ReadProperty("Permissoes", meDSPermissoes)
  Me.ButtonType = PropBag.ReadProperty("ButtonType", AButtons.ButtonTypes.[Simple Flat])
  Me.ButtonsExtras = PropBag.ReadProperty("ButtonsExtras", meBTNExtras)
  Me.ButtonColorDesab = PropBag.ReadProperty("ButtonColorDesab", cCorDesab)
  Me.ButtonColor = PropBag.ReadProperty("ButtonColor", &HE0E0E0)
  Me.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor) '&HC0C0C0)
End Sub

Public Property Get IsConnected() As Boolean
    If Not mDS Is Nothing Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Property

Public Sub DesConectar()
    Debug.Print ("Insignia_DSControl.ideDSControl.DesConectar")
    If IsConnected Then
        mDS.DesConectar
        Set mDS = Nothing
    End If
End Sub

Public Function Conectar(ByVal Source As String, _
                         ByRef ActiveConnection, _
                         Optional ByVal pCursorLocation As CursorLocationEnum = adUseClient, _
                         Optional ByVal pCursorType As CursorTypeEnum = adOpenKeyset, _
                         Optional ByVal pLockType As LockTypeEnum = adLockPessimistic) As eDSConexao
    Debug.Print ("Insignia_DSControl.ideDSControl.Conectar")

  Dim nRet As eDSConexao
  
  gsParent = UserControl.Parent.Name
  
  If Not IsConnected Then Set mDS = New CDSControl
    
  nRet = mDS.Conectar(Source, ActiveConnection, pCursorLocation, pCursorType, pLockType)
  
  If meModelo = mdMaster Then
    Select Case nRet
      Case Is = cnSucesso
        fraPanel(cPNLIdent).Caption = "Banco de Dados Conectado!"
      Case Is = cnVazio
        fraPanel(cPNLIdent).Caption = "Banco de Dados Conectado!"
        Call EnabledNaveg(False, False, False, False, True)
        fraPanel(cPNLCont).Enabled = False
        Call ButtonsFuncEnabled(False)
    End Select
  End If
  
  If nRet = cnErroProcesso Then UserControl.Enabled = False

End Function

Private Sub ButtonsFuncEnabled(bEnabled As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.ButtonsFuncEnabled")
  Dim btn As AButton

  For Each btn In abtBarra1
    Select Case btn.Index
    Case Is = 0
      btn.Enabled = True
      
    Case Is = 1, 2
      btn.Enabled = bEnabled
    
    Case Is = 7 'ATUALIZAR
      btn.Enabled = bEnabled And (meBTNExtras = beAllButtons Or meBTNExtras = beRequery Or meBTNExtras = beSearch_Requery Or meBTNExtras = bePrinter_Requery)
      
    Case Is = 3 'PESQUISAR
      btn.Enabled = bEnabled And (meBTNExtras = beAllButtons Or meBTNExtras = beSearch Or meBTNExtras = beSearch_Requery Or meBTNExtras = bePrinter_Search)
    
    Case Is = 4 'IMPRIMIR
      btn.Enabled = bEnabled And (meBTNExtras = beAllButtons Or meBTNExtras = bePrinter Or meBTNExtras = bePrinter_Requery Or meBTNExtras = bePrinter_Search)
      
    Case Is = 5 'ORDERBY
      Combos(cCmbOrder).Enabled = bEnabled And Combos(cCmbOrder).ListCount > 0
      abtBarra1(5).Enabled = Combos(cCmbOrder).Enabled
    End Select
    
  Next
    
  
  'Se estiver mandando habilitar botoes entao deve
  'verificar se é permitido utiliza-lós
  If bEnabled Then
    Me.Permissoes = meDSPermissoes
  Else
    Me.ButtonColor = mBtnColor 'atualizando a cor dos botoes
  End If
End Sub

Public Property Get Operacao() As eDSOperacao
  Operacao = meOperacao
End Property

Public Property Let Operacao(ByVal vNewValue As eDSOperacao)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let Operacao")
  Dim bOpView As Boolean
  
  bOpView = vNewValue = opVisualizacao
  
  Dim btn As AButton
  For Each btn In abtBarra1
    btn.Visible = bOpView
  Next
  Combos(cCmbOrder).Visible = bOpView
    
  If bOpView Then
    Me.ButtonsExtras = meBTNExtras
  Else
    abtBarra1(3).Visible = False  'Localizar
    abtBarra1(4).Visible = False  'Imprimir
  End If
    
  fraBarra(2).Visible = Not bOpView
  chkContinuarInsert.Visible = vNewValue = opInclusao
  
  If meModelo = mdMaster Then
    Call EnabledPanelNav(bOpView)
    
    Dim sDesc As String
    Select Case vNewValue
      Case Is = opAlteracao
        sDesc = "Operação de Alteração"
      Case Is = opInclusao
        If mbAddContFlag Then
          sDesc = "Operação de Inclusão Continua"
        Else
          sDesc = "Operação de Inclusão"
        End If
      Case Is = opVisualizacao
        sDesc = "Operação de Visualização"
    End Select
    
    fraPanel(cPNLIdent).Caption = sDesc
  End If
  
  RaiseEvent Operacao(vNewValue, meOperacao)
  meOperacao = vNewValue
End Property

'Public Sub AddButton(ButtonKey As String, Caption As String, _
'                     Optional tpButton As eTBButtonStyles = bsButton, _
'                     Optional Picture As StdPicture, _
'                     Optional ToolTip As String)
'  TBar(cTBarFunc).AddButtonEx ButtonKey, tpButton, Picture, ToolTip, , , Caption
'End Sub

'Public Property Get FlatButtons() As Boolean
'  FlatButtons = mbFlatButtons
'End Property
'
'Public Property Let FlatButtons(ByVal vNewValue As Boolean)
'  Dim oT As asxToolbar
'  Dim i As Integer
'
'  If vNewValue Then
'    i = apFlat
'  Else
'    i = apStandard
'  End If
'
'  For Each oT In TBar
'    oT.Appearance = i
'  Next
'
'  mbFlatButtons = vNewValue
'
'  PropertyChanged "FlatButtons"
'End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = fraBarra(0).BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let BackColor")
  Dim oCtr As ideFrame
  
  For Each oCtr In fraBarra
    oCtr.BackColor = vNewValue
  Next
  Set oCtr = Nothing
  
  For Each oCtr In fraPanel
    oCtr.BackColor = vNewValue
  Next
  Set oCtr = Nothing

  chkContinuarInsert.BackColor = vNewValue
  
  PropertyChanged "BackColor"
End Property

Public Property Get ButtonColor() As OLE_COLOR
  ButtonColor = mBtnColor
End Property

Public Property Let ButtonColor(ByVal vNewValue As OLE_COLOR)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let ButtonColor")
  Dim btn As AButton
  
  For Each btn In abtBarra1
    btn.BackColor = IIf(btn.Enabled, vNewValue, mBtnColorInative)
    Set btn = Nothing
  Next

  For Each btn In abtBarra2
    btn.BackColor = IIf(btn.Enabled, vNewValue, mBtnColorInative)
    Set btn = Nothing
  Next
  
  For Each btn In abtBarra3
    btn.BackColor = IIf(btn.Enabled, vNewValue, mBtnColorInative)
    Set btn = Nothing
  Next
  
  mBtnColor = vNewValue
  PropertyChanged "ButtonColor"
End Property

Public Property Get ButtonColorDesab() As OLE_COLOR
  ButtonColorDesab = mBtnColorInative
End Property
Public Property Let ButtonColorDesab(ByVal vNewValue As OLE_COLOR)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let ButtonColorDesab")
  mBtnColorInative = vNewValue
  
  Me.ButtonColor = mBtnColor  'Força a mudança de cor de todos os botoes, aqui deve passar a cordo do botao ativo
  PropertyChanged "ButtonColor"
End Property


Public Property Get ButtonType() As AButtons.ButtonTypes
  ButtonType = meButtonType
End Property

Public Property Let ButtonType(ByVal vNewValue As AButtons.ButtonTypes)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let ButtonType")
  Dim btn As AButton

  For Each btn In abtBarra1
    btn.ButtonType = vNewValue
    Set btn = Nothing
  Next
  For Each btn In abtBarra2
    btn.ButtonType = vNewValue
    Set btn = Nothing
  Next
  For Each btn In abtBarra3
    btn.ButtonType = vNewValue
    Set btn = Nothing
  Next
  
  meButtonType = vNewValue
  PropertyChanged "ButtonType"
End Property

Public Property Get CaptionColor() As OLE_COLOR
  CaptionColor = fraPanel(0).ForeColor
End Property
Public Property Let CaptionColor(ByVal vNewValue As OLE_COLOR)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let CaptionColor")
  Dim oT As AButton

  For Each oT In abtBarra3
    oT.ForeColor = vNewValue
  Next
  Set oT = Nothing
  
  Dim oP As ideFrame
  For Each oP In fraPanel
    oP.ForeColor = vNewValue
  Next
  Set oP = Nothing
  
  chkContinuarInsert.ForeColor = vNewValue
  fraBarra(3).ForeColor = vNewValue
  
  PropertyChanged "CaptionColor"
End Property

Public Property Get ButtonsExtras() As eDSBotoesExtras
  ButtonsExtras = meBTNExtras
End Property

Public Property Let ButtonsExtras(ByVal vNewValue As eDSBotoesExtras)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let ButtonsExtras")
  Dim bI As Boolean, bL As Boolean, bA As Boolean

  Select Case vNewValue
    Case Is = bePrinter_Requery
      bI = True:  bA = True
    Case Is = bePrinter_Search
      bI = True:  bL = True
    Case Is = beSearch_Requery
      bL = True:  bA = True
    Case Is = beRequery
      bA = True
    Case Is = bePrinter
      bI = True
    Case Is = beSearch
      bL = True
    Case Is = beAllButtons
      bI = True:  bL = True:  bA = True
      
  End Select

  abtBarra1(3).Enabled = bL 'Localizar
  abtBarra1(4).Enabled = bI 'Imprimir
  abtBarra1(7).Enabled = bA 'Atualizar

  meBTNExtras = vNewValue
  PropertyChanged "ButtonsExtras"
End Property

Public Sub OperacaoPesquisar()
    Debug.Print ("Insignia_DSControl.ideDSControl.OperacaoPesquisar")
  If meModelo = mdMaster Then
    Dim bCheck As Boolean
        
    'Botao não esta visivel entao presume-se que nao se pode usa-lo
    'Caso esteja sendo chamado por tecla de atalho
    If Not abtBarra1(3).Visible Then Exit Sub   'Localizar
    
    If abtBarra1(3).BackColor = CorBtnCheck Then
      abtBarra1(3).BackColor = mBtnColor
      bCheck = False
    Else
      abtBarra1(3).BackColor = CorBtnCheck
      bCheck = True
    End If
    
    If bCheck Then
      mbCancelEvent = False
      RaiseEvent ClickPesquisa(mbCancelEvent)
      If mbCancelEvent Then Exit Sub
      
      If Combos(cCmbPesq).ListCount = 0 Then Exit Sub
    End If
    
    fraBarra(cPNLPesq).Visible = bCheck
    fraPanel(cPNLIdent).Visible = Not bCheck
    
    If bCheck Then
      On Error Resume Next
      txtPesquisa.SetFocus
      On Error GoTo 0
    End If
  
  ElseIf meModelo = mdSimples Then
    Static bUsou As Boolean
    
    'Botao não esta visivel entao presume-se que nao se pode usa-lo
    'Caso esteja sendo chamado pelo Teclado
    If Not abtBarra1(4).Visible Then Exit Sub ' TBar.ButtonVisible("Localizar") Then Exit Sub
    
    If FSearch.cmbCampos.ListCount = 0 Then
      Set FSearch = Nothing
      Exit Sub
    End If
    
    mbCancelEvent = False
    RaiseEvent ClickPesquisa(mbCancelEvent)
    If mbCancelEvent Then Exit Sub
    
    If Not bUsou Then
      FSearch.ShowForm mDS
      bUsou = True
    Else
      FSearch.Show vbModal    'o form já esta carregado
    End If
  End If
  
End Sub

Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
    Debug.Print ("Insignia_DSControl.ideDSControl.UserControl_GetDataMember")
 If Ambient.UserMode And IsConnected Then Call mDS.Class_GetDataMember(DataMember, Data)
End Sub

Private Sub UserControl_Initialize()
  shpRegistro.Left = 15
  shpRegistro.Width = 0
  
  fraBarra(2).BackColor = fraBarra(0).BackColor
  fraBarra(3).BackColor = fraBarra(0).BackColor
  
  Combos(cCmbOrder).Enabled = False
  abtBarra1(5).Enabled = False
End Sub

Private Sub UserControl_Terminate()
  On Error Resume Next
  SaveSetting App.EXEName, "Campo Order", gsParent, Combos(cCmbOrder).ListIndex
  SaveSetting App.EXEName, "Campo Pesquisa", gsParent, Combos(cCmbPesq).ListIndex
  SaveSetting App.EXEName, "Valor Pesquisa", gsParent, txtPesquisa.Text
  On Error GoTo 0
      
  Me.DesConectar
End Sub

Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.KeyDown")
  If Ambient.UserMode And IsConnected Then
    MFuncoes.KeyDown Me, KeyCode, Shift
  End If
End Sub

Private Sub UserControl_Resize()
    Debug.Print ("Insignia_DSControl.ideDSControl.UserControl_Resize")
    Dim iH As Integer
    
    If meModelo = mdMaster Then
        If Width < 6200 Then Width = 6200
        iH = 1
    ElseIf meModelo = mdSimples Then
        If Width < 4665 Then Width = 4665
        iH = 0
    End If
    
    If Height <> (fraBarra(iH).Top + fraBarra(iH).Height) Then
        Height = (fraBarra(iH).Top + fraBarra(iH).Height)
    End If
    
    If meModelo = mdMaster Then
        With fraPanel(cPNLIdent)
            .Width = (Width - .Left) - 80
    
            fraBarra(cPNLPesq).Move .Left, .Top
            fraBarra(cPNLPesq).Height = .Height
            fraBarra(cPNLPesq).Width = .Width
        End With
    End If
    
    fraBarra(2).Left = 15
End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let Enabled")
  UserControl.Enabled = vNewValue
  
  Dim oC As Control
  
  On Local Error Resume Next
  For Each oC In Controls
    oC.Enabled = vNewValue
  Next
  On Error GoTo 0
End Property

'Private Sub fraPanel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'   Select Case Index
'      Case Is = cPNLCont
'         UserControl.MousePointer = 9
'         If Button = vbLeftButton Then
'            Call MoveRegShape(x)
'         End If
'   End Select
'End Sub
'
'Private Sub fraPanel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'   Select Case Index
'      Case Is = cPNLCont
'         If Button = vbLeftButton Then
'            Call MoveRegShape(x)
'         End If
'   End Select
'End Sub
'
'Private Sub fraPanel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'   Select Case Index
'      Case Is = cPNLCont
'         UserControl.MousePointer = 0
'         mAPI.SoltarCursor
'   End Select
'End Sub

'Private Sub MoveRegShape(x As Single)
'  Dim nRegAtual As Long
'  Dim nRegistros As Long
'  Dim nFator As Long
'  Dim nWidth As Integer
'
'  If mDS Is Nothing Then Exit Sub
'
'  Call mAPI.PrenderCursor(fraPanel(cPNLCont).hwnd)
'  nWidth = fraPanel(cPNLCont).Width
'  With mDS
'    nRegistros = .RS.RecordCount
'    nRegAtual = .AbsolutePosition
'
'    If nRegistros > 0 Then
'      nFator = nWidth / nRegistros
'      If x <= nWidth And x >= 0 Then
'        If x < 100 Then
'          .MoveFirst
'        ElseIf x > (nWidth - 80) Then
'          .MoveLast
'        Else
'          On Error Resume Next
'          nRegAtual = Int(x / nFator) + 1
'          If nRegAtual > nRegistros Then
'            .AbsolutePosition = nRegistros
'          Else
'            .AbsolutePosition = nRegAtual
'          End If
'        End If
'      End If
'    End If
'  End With
'End Sub

Public Property Get AddNewContinueFlag() As Boolean
  AddNewContinueFlag = mbAddContFlag
End Property

Public Sub AddNewContinue()
    Debug.Print ("Insignia_DSControl.ideDSControl.AddNewContinue")
   mbAddContFlag = CBool(chkContinuarInsert.Value)
End Sub

Public Sub AddNew(Optional ByVal FieldList, Optional ByVal Values)
    Debug.Print ("Insignia_DSControl.ideDSControl.AddNew")
  'Botao não esta visivel entao presume-se que nao se pode usa-lo
  If Not abtBarra1(0).Enabled Then Exit Sub

  If meOperacao = opVisualizacao Then
    mDS.AddNew FieldList, Values
  End If
End Sub

Public Sub Delete()
    Debug.Print ("Insignia_DSControl.ideDSControl.Delete")
  'Botao não esta habilitado entao presume-se que nao se pode usa-lo
  If Not abtBarra1(2).Enabled Then Exit Sub
  
  Dim bNotMessage As Boolean
  mbCancelEvent = False
  RaiseEvent AntesDelete(mbCancelEvent, bNotMessage)
  If Not mbCancelEvent Then
    If meOperacao = opVisualizacao Then mDS.Delete Not bNotMessage
  End If
End Sub

Public Sub Edit()
    Debug.Print ("Insignia_DSControl.ideDSControl.Edit")
  'Botao não esta visivel entao presume-se que nao se pode usa-lo
  If Not abtBarra1(1).Enabled Then Exit Sub

  If meOperacao = opVisualizacao Then
    mbCancelEvent = False
    RaiseEvent AntesEdit(mbCancelEvent)
    If mbCancelEvent Then Exit Sub
   
    With mDS
      .Resync
      If Not .RS.EOF And Not .RS.BOF Then .RS.Move 0  'atualiza os controles (Atualização forçada)
    End With
    mDS.Operacao = opAlteracao
  End If
End Sub

Public Sub Update()
    Debug.Print ("Insignia_DSControl.ideDSControl.Update")
  Dim nRCountOld As Integer
  Dim meOpeOld   As eDSOperacao
  
  meOpeOld = meOperacao
  nRCountOld = mDS.RS.RecordCount
  If mDS.Update Then
    If mbAddContFlag Then
      Select Case meOpeOld
        Case Is = opInclusao
          mDS.AddNew
          Exit Sub
      End Select
    End If
  
    Call ButtonsFuncEnabled(mDS.RS.RecordCount > 0)
    
    If meModelo = mdMaster Then
      If mDS.RS.RecordCount = 0 Then
        'Call ButtonsFuncEnabled(False)
        shpRegistro.Width = MFuncoes.ContadorWidth(fraPanel(cPNLCont), meOperacao, 0, 0)
        Call EnabledNaveg(False, False, False, False, False)
      
'      ElseIf nRCountOld = 0 Then
'        Call ButtonsFuncEnabled(True)
      End If
    End If
  End If
End Sub

Private Sub mDS_AntesAddNew(Cancel As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_AntesAddNew")
  RaiseEvent AntesAddNew(Cancel)
End Sub

Private Sub mDS_AntesCancel(Cancel As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_AntesCancel")
  RaiseEvent AntesCancel(Cancel)
End Sub

Private Sub mDS_AntesRequery(Cancel As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_AntesRequery")
  RaiseEvent AntesRequery(Cancel)
End Sub

Private Sub mDS_AntesUpdate(Cancel As Boolean, eOperacao As eDSOperacao)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_AntesUpdate")
  RaiseEvent AntesUpdate(Cancel, eOperacao)
End Sub

Private Sub mDS_DepoisRequery()
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_DepoisRequery")
  RaiseEvent DepoisRequery
End Sub

Private Sub mDS_DepoisUpdate(eOperacao As eDSOperacao)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_DepoisUpdate")
  RaiseEvent DepoisUpdate(eOperacao)
End Sub

Private Sub mDS_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_FieldChangeComplete")
  RaiseEvent FieldChangeComplete(cFields, Fields, pError, adStatus, pRecordset)
End Sub

Private Sub mDS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_MoveComplete")
  If meModelo = mdMaster Then
    Dim nRegCount As Long
    Dim nPos As Long
  
    On Error GoTo Sair:
     
    With mDS.RS
      nRegCount = .RecordCount
      If nRegCount > 0 Then
        If .EOF And .BOF Then .AbsolutePosition = 1
          nPos = .AbsolutePosition
          Else: nPos = 0
        End If
  
      nRegCount = .RecordCount
      If nRegCount <= 1 Then
        Call EnabledNaveg(False, False, False, False)
        shpRegistro.Width = MFuncoes.ContadorWidth(fraPanel(cPNLCont), meOperacao, nPos, nRegCount)
        Exit Sub
      End If
  
      Select Case nPos
        Case Is = adPosBOF, 1
          Call EnabledNaveg(False, False, True, True)
        Case Is = adPosEOF, nRegCount
          Call EnabledNaveg(True, True, False, False)
        Case Else
          Call EnabledNaveg(True, True, True, True)
      End Select
      shpRegistro.Width = MFuncoes.ContadorWidth(fraPanel(cPNLCont), meOperacao, nPos, nRegCount)
      DoEvents
    End With
  End If
Sair:
  RaiseEvent MoveComplete(adReason, pError, adStatusCancel, pRecordset)
  On Error GoTo 0
End Sub

Private Sub mDS_Operacao(ByVal eOPAtual As eDSOperacao, ByVal eOPAterior As eDSOperacao)
    Debug.Print ("Insignia_DSControl.ideDSControl.mDS_Operacao")
  Me.Operacao = eOPAtual
End Sub

Private Sub EnabledPanelNav(ByVal pbValue As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.EnabledPanelNav")
   Dim nRegCount As Long
   Dim nPos As Long
   
   On Error GoTo TrataErro:
   If Not pbValue Then
      Call EnabledNaveg(False, False, False, False, True)
   Else
      nPos = mDS.AbsolutePosition
      nRegCount = mDS.RS.RecordCount
      If nRegCount <= 1 Then
         Call EnabledNaveg(False, False, False, False, True)
         Exit Sub
      End If
      
      Select Case nPos
         Case Is = adPosBOF, 1
            Call EnabledNaveg(False, False, True, True, True)
         Case Is = adPosEOF, nRegCount
            Call EnabledNaveg(True, True, False, False, True)
         Case Else
            Call EnabledNaveg(True, True, True, True, True)
      End Select
      DoEvents
   End If
   fraPanel(cPNLCont).Enabled = pbValue
   Exit Sub
TrataErro:
'   MsgBox Err.Description & vbCrLf & Err.Source, vbCritical, "XDataSource_EnabledPanelNav"
End Sub

Private Sub EnabledNaveg(bNavPri As Boolean, _
                         bNavAnt As Boolean, _
                         bNavPro As Boolean, _
                         bNavUlt As Boolean, _
                         Optional bForcaExec As Boolean)
    Debug.Print ("Insignia_DSControl.ideDSControl.EnabledNaveg")
                         
  If meOperacao <> opVisualizacao And bForcaExec = False Then Exit Sub
  
  abtBarra2(0).Enabled = bNavPri
  abtBarra2(1).Enabled = bNavAnt
  abtBarra2(2).Enabled = bNavPro
  abtBarra2(3).Enabled = bNavUlt
  
  Dim btn As AButton
  
  For Each btn In abtBarra2
    btn.BackColor = IIf(btn.Enabled, mBtnColor, mBtnColorInative)
    Set btn = Nothing
  Next
End Sub

Public Property Get DataSource() As CDSControl
  Set DataSource = mDS
End Property

Public Property Get Permissoes() As eDSPermissoes
  Permissoes = meDSPermissoes
End Property
Public Property Let Permissoes(ByVal vNewValue As eDSPermissoes)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let Permissoes")
    meDSPermissoes = vNewValue
    
    Dim i As Integer
    Dim sKey As String
    
    Dim bI As Boolean, bA As Boolean, bD As Boolean, bC As Boolean
        
    Select Case vNewValue
        Case peTodos:     bI = True:  bA = True:  bD = True:  bC = True:
        Case peIncluir:   bI = True
        Case peAlterar:   bA = True
        Case peExcluir:   bD = True
        Case peNenhuma:   'Todos desabilitados
        Case peIncluir_Excluir:   bI = True:  bD = True
        Case peIncluir_Alterar:   bI = True:  bA = True
        Case peAlterar_Excluir:   bA = True:  bD = True
    End Select
        
    abtBarra1(0).Enabled = bI
    abtBarra1(1).Enabled = bA
    abtBarra1(2).Enabled = bD
    
    Me.ButtonColor = mBtnColor 'Força a atualização das cores de todos os botoes
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Debug.Print ("Insignia_DSControl.ideDSControl.UserControl_WriteProperties")
    Call PropBag.WriteProperty("Operacao", meOperacao, 0)
    Call PropBag.WriteProperty("Modelo", meModelo, mdMaster)
    Call PropBag.WriteProperty("CaptionColor", abtBarra1(0).ForeColor, vbBlack)
    Call PropBag.WriteProperty("Permissoes", meDSPermissoes, peTodos)
    Call PropBag.WriteProperty("ButtonType", meButtonType, AButtons.ButtonTypes.[Simple Flat])
    Call PropBag.WriteProperty("ButtonsExtras", meBTNExtras, 0)
    Call PropBag.WriteProperty("ButtonColorDesab", mBtnColorInative, cCorDesab)
    Call PropBag.WriteProperty("ButtonColor", mBtnColor, &HE0E0E0)
    Call PropBag.WriteProperty("BackColor", fraBarra(0).BackColor, &HC0C0C0)
End Sub

Public Sub MontarPesquisa(ByVal psAliasFieldMask As String)
  Debug.Print ("Insignia_DSControl.ideDSControl.MontarPesquisa")
  Dim sCapt As String, sField As String, sMask As String
  Dim aI() As String, aL() As String

'  abtBarra1(3).Enabled = True  'TBar(cTBarFunc).ButtonEnabled("Localizar") = True

  aI = Split(psAliasFieldMask, "|")

  Dim i As Byte
  For i = 0 To UBound(aI)
    aL = Split(aI(i), ",")

    sCapt = Trim$(aL(0))
    sField = Trim$(aL(1))
    sMask = LTrim$(aL(2))
     
    If sMask <> "ñP" Then '"ñP" Significa que não entra na pesquisa
      With Combos(cCmbOrder)
        .AddItem sCapt
        
        ReDim Preserve maFields(.ListCount)
        maFields(.ListCount) = sField
      End With
      
      If meModelo = mdMaster Then 'Se incluir este se for tipo completo
        With Combos(cCmbPesq)
          .AddItem sCapt
          
          ReDim Preserve maMasks(.ListCount)
          maMasks(.ListCount) = sMask
        End With
      End If
    End If
  Next
  
  With Combos(cCmbOrder)
    .Enabled = .ListCount > 0
    abtBarra1(5).Enabled = .Enabled
    Me.ButtonColor = mBtnColor
  End With
  
  If meModelo = mdSimples Then
    Load FSearch
    FSearch.MontarTela psAliasFieldMask
  End If
  Call LerREG
End Sub

Private Sub OrderAscDesc()
    Debug.Print ("Insignia_DSControl.ideDSControl.OrderAscDesc")
  mbSortDesc = Not mbSortDesc
  abtBarra1(5).BackColor = IIf(mbSortDesc, CorBtnCheck, mBtnColor)
  
  Combos_Click cCmbOrder
End Sub

Private Sub Combos_Click(Index As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.Combos_Click")
  Select Case Index
    Case Is = cCmbOrder
      With Combos(Index)
        If .ListIndex >= 0 Then
          On Error GoTo Sair:
          mDS.Sort = maFields(.ListIndex + 1) & IIf(mbSortDesc, " DESC", " ASC")
          On Error GoTo 0
        End If
      End With
    
    Case Is = cCmbPesq
      Dim sMask As String
      
      With Combos(Index)
        If .ListIndex < 0 Then .ListIndex = 0
        
        sMask = maMasks(.ListIndex + 1)
        txtPesquisa.DataField = maFields(.ListIndex + 1)
      End With
      
      With txtPesquisa
        .Mask = sMask
        If sMask = "" Then .MaxLength = 0
        
        On Error Resume Next
        .SetFocus
        On Error GoTo 0
      End With
  End Select
Sair:
End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
    Debug.Print ("Insignia_DSControl.ideDSControl.txtPesquisa_KeyPress")
  If KeyAscii = vbKeyReturn Then
    With txtPesquisa
      Call mDS.Search(.DataField, .Text)
    
      On Error Resume Next
      .SetFocus
      On Error GoTo 0
    End With
  End If
End Sub

Private Sub LerREG()
    Debug.Print ("Insignia_DSControl.ideDSControl.LerREG")
  ' Ler os valores dos campo de indexação e Pesquisa usado pela ultima vezes
  With Combos(cCmbOrder)
    If .ListCount > 0 Then
      On Error GoTo ErrIndex
      .ListIndex = GetSetting(App.EXEName, "Campo Order", gsParent, "0")
      On Error GoTo 0
    Else
ErrIndex:
      Call SaveSetting(App.EXEName, "Campo Order", gsParent, 0)
      Err.Clear
    End If
  End With
  
  With Combos(cCmbPesq)
    If .ListCount > 0 Then
      On Error GoTo ErrPesquisa
      .ListIndex = GetSetting(App.EXEName, "Campo Pesquisa", gsParent, "0")
      On Error GoTo 0
    Else
ErrPesquisa:
      Call SaveSetting(App.EXEName, "Campo Pesquisa", gsParent, 0)
      Err.Clear
    End If
  End With
  
  txtPesquisa.Text = GetSetting(App.EXEName, "Valor Pesquisa", gsParent, "")
End Sub

Public Property Get Modelo() As eDSModelo
  Modelo = meModelo
End Property

Public Property Let Modelo(ByVal vNewValue As eDSModelo)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let Modelo")
  meModelo = vNewValue
  fraBarra(1).Visible = meModelo = mdMaster
  
  UserControl_Resize
End Property

Public Property Let Informe(ByVal vNewValue As String)
    Debug.Print ("Insignia_DSControl.ideDSControl.Property Let Informe")
  fraPanel(cPNLIdent).Caption = vNewValue
End Property
