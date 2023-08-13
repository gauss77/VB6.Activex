VERSION 5.00
Object = "{C6FEE5AC-DF5F-47A6-BE77-6DCE10AA8AB9}#4.2#0"; "ideDSControl.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Begin VB.Form frmClientes 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10665
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPesquisa 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   2520
      Picture         =   "frmClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4095
      Width           =   300
   End
   Begin VB.CommandButton cmdPesquisa 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   2520
      Picture         =   "frmClientes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2745
      Width           =   300
   End
   Begin VB.CheckBox chkCampo 
      Alignment       =   1  'Right Justify
      Caption         =   "&Bloquear Registro"
      DataField       =   "REGBLOQ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   0
      Left            =   7440
      TabIndex        =   2
      ToolTipText     =   "Registros Bloqueados não apareceram nas listagens de pesquisa "
      Top             =   960
      Width           =   2265
   End
   Begin VB.TextBox txtMemo 
      DataField       =   "OBS"
      Height          =   690
      Index           =   0
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   5880
   End
   Begin Insignia_DSControl.ideDSControl dscMaster 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
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
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "NOME"
      Height          =   300
      Index           =   2
      Left            =   1800
      TabIndex        =   27
      Tag             =   "Obrigatorio"
      ToolTipText     =   "Nome ou Razão Social"
      Top             =   1815
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      TextCase        =   1
      RawText         =   0
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "DT_NASC"
      Height          =   300
      Index           =   3
      Left            =   8040
      TabIndex        =   28
      Top             =   1815
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "ENDERECO"
      Height          =   300
      Index           =   6
      Left            =   1800
      TabIndex        =   29
      Top             =   3165
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      TextCase        =   1
      RawText         =   0
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "BAIRRO"
      Height          =   300
      Index           =   7
      Left            =   1800
      TabIndex        =   30
      Top             =   3615
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   30
      TextCase        =   1
      RawText         =   0
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "CEP"
      Height          =   300
      Index           =   8
      Left            =   5145
      TabIndex        =   31
      Top             =   3615
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   9
      TextMask        =   6
      RawText         =   6
      Mask            =   "#####-###"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "RG_IE"
      Height          =   300
      Index           =   10
      Left            =   1800
      TabIndex        =   32
      ToolTipText     =   "Informe o Registro Geral ou Insc. Estadual se Pessoa Júridica"
      Top             =   4515
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   529
      Alignment       =   2
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextCase        =   1
      RawText         =   0
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "CGC_CPF"
      Height          =   300
      Index           =   11
      Left            =   4620
      TabIndex        =   33
      ToolTipText     =   "Informe o CPF ou CNPJ se Pessoa Júridica"
      Top             =   4515
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   18
      RawText         =   0
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "FONE2"
      Height          =   300
      Index           =   14
      Left            =   8040
      TabIndex        =   34
      Top             =   3600
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   14
      TextMask        =   9
      RawText         =   9
      Mask            =   "(##)#####-####"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "FAX"
      Height          =   300
      Index           =   16
      Left            =   8040
      TabIndex        =   35
      Top             =   4500
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      Alignment       =   2
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   14
      TextMask        =   9
      RawText         =   9
      Mask            =   "(##)#####-####"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "FONE1"
      Height          =   300
      Index           =   13
      Left            =   8040
      TabIndex        =   36
      Top             =   3150
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   14
      TextMask        =   9
      RawText         =   9
      Mask            =   "(##)#####-####"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "DT_CADASTRO"
      Height          =   300
      Index           =   1
      Left            =   5145
      TabIndex        =   37
      Tag             =   "Obrigatorio"
      Top             =   1365
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   529
      Alignment       =   2
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "PESSOA"
      Height          =   300
      Index           =   4
      Left            =   1800
      TabIndex        =   38
      Tag             =   "Obrigatorio"
      Top             =   1365
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Alignment       =   2
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   1
      TextMask        =   9
      TextCase        =   1
      RawText         =   9
      FloatFormat     =   2
      Mask            =   "?"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "ID"
      Height          =   300
      Index           =   0
      Left            =   1800
      TabIndex        =   39
      Tag             =   "Obrigatorio"
      Top             =   915
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      Alignment       =   1
      Appearance      =   0
      BackColor       =   15987699
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FloatFormat     =   2
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "LIMITE_CREDITO"
      DataSource      =   "XDSStandard"
      Height          =   300
      Index           =   5
      Left            =   8040
      TabIndex        =   40
      Top             =   2265
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   529
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   18
      Text            =   "R$ 0,00"
      TextMask        =   4
      RawText         =   4
      FloatFormat     =   2
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "ID_CIDADE"
      Height          =   300
      Index           =   9
      Left            =   1800
      TabIndex        =   41
      Top             =   4065
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   529
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "CELULAR"
      Height          =   300
      Index           =   15
      Left            =   8040
      TabIndex        =   42
      Top             =   4050
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   14
      TextMask        =   9
      RawText         =   9
      Mask            =   "(##)#####-####"
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "NOME_FANTASIA"
      Height          =   300
      Index           =   18
      Left            =   1800
      TabIndex        =   43
      Top             =   2250
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      TextCase        =   1
      RawText         =   0
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampo 
      DataField       =   "ID_GRUPO"
      Height          =   300
      Index           =   12
      Left            =   1800
      TabIndex        =   44
      Top             =   2715
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   529
      Alignment       =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   2
      Text            =   "0"
      TextMask        =   3
      RawText         =   3
      FloatFormat     =   2
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampoFK 
      DataField       =   "NOME"
      Height          =   300
      Index           =   12
      Left            =   2880
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2715
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   529
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   30
      TextCase        =   1
      RawText         =   0
      FloatFormat     =   2
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txtCampoFK 
      DataField       =   "NOME"
      Height          =   300
      Index           =   9
      Left            =   2880
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4065
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   529
      Appearance      =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   30
      TextCase        =   1
      RawText         =   0
      FloatFormat     =   2
      FontName        =   "Microsoft Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label lblRotulos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Código ID:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008D550A&
      Height          =   225
      Index           =   0
      Left            =   450
      TabIndex        =   26
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Celular:"
      Height          =   195
      Index           =   111
      Left            =   7425
      TabIndex        =   25
      Top             =   4095
      Width           =   525
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   450
      TabIndex        =   24
      Top             =   4110
      Width           =   555
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limite Crédito:"
      Height          =   195
      Index           =   34
      Left            =   6960
      TabIndex        =   23
      Top             =   2310
      Width           =   990
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cadastro:"
      ForeColor       =   &H008D550A&
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   22
      Top             =   1410
      Width           =   1065
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome/R. Social:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008D550A&
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   21
      Top             =   1860
      Width           =   1140
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DT Nasc.:"
      Height          =   195
      Index           =   14
      Left            =   7215
      TabIndex        =   20
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   450
      TabIndex        =   19
      Top             =   3210
      Width           =   735
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   450
      TabIndex        =   18
      Top             =   3660
      Width           =   480
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP:"
      Height          =   195
      Index           =   5
      Left            =   4635
      TabIndex        =   17
      Top             =   3660
      Width           =   360
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RG / IE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   450
      TabIndex        =   16
      Top             =   4560
      Width           =   510
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ/CPF:"
      Height          =   195
      Index           =   11
      Left            =   3615
      TabIndex        =   15
      Top             =   4560
      Width           =   825
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone Comercial:"
      Height          =   195
      Index           =   100
      Left            =   6810
      TabIndex        =   14
      Top             =   3645
      Width           =   1140
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX:"
      Height          =   195
      Index           =   12
      Left            =   7605
      TabIndex        =   13
      Top             =   4590
      Width           =   345
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observações..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   450
      TabIndex        =   12
      Top             =   5025
      Width           =   1125
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone Residêncial:"
      Height          =   195
      Index           =   6
      Left            =   6675
      TabIndex        =   11
      Top             =   3195
      Width           =   1275
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pessoa:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008D550A&
      Height          =   195
      Index           =   4
      Left            =   450
      TabIndex        =   10
      Top             =   1410
      Width           =   570
   End
   Begin VB.Label lblIdent 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FÍSICA"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   2115
      TabIndex        =   9
      Top             =   1410
      Width           =   615
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Fantasia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   450
      TabIndex        =   8
      Top             =   2310
      Width           =   1125
   End
   Begin VB.Label lblIdent 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente desde "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00212121&
      Height          =   195
      Index           =   1
      Left            =   6390
      TabIndex        =   7
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label lblRotulos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   450
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblDocInvalido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº documento inválido!"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   4620
      TabIndex        =   5
      Top             =   4845
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim sSQL As String
  Dim sConn As String
    
  Me.MousePointer = vbHourglass
  
  sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\QRAuto.mdb"
  sSQL = "SELECT * FROM TB_CLIENTES WHERE REGBLOQ = FALSE"
  
  If dscMaster.Conectar(sSQL, sConn) = cnErroProcesso Then
    Me.MousePointer = vbDefault
    Unload Me
  Else
    Load Me
    Call ConfigurarDados
    Me.MousePointer = vbDefault
    Me.Show
  End If
End Sub

Private Sub ConfigurarDados()
  Dim oT  As ActiveText

  Static bMontPesq As Boolean
  Dim sPesq As String, sMask As String
  
  For Each oT In txtCampo
    Set oT.DataSource = dscMaster.DataSource.rs
  
    If Not bMontPesq Then
      Select Case oT.Index
        Case Is = 0, 4, 2, 1, 10, 11, 18
          Select Case oT.TextMask
            Case Is = [Integer Mask]
              sMask = "############"
            Case Else
              sMask = oT.Mask
          End Select

          sPesq = sPesq & _
                  lblRotulos(oT.Index).Caption & "," & _
                  oT.DataField & "," & sMask & "|"
      End Select
    End If
    Set oT = Nothing
  Next
  
  Set chkCampo(0).DataSource = dscMaster.DataSource.rs
  Set txtMemo(0).DataSource = dscMaster.DataSource.rs
  
  If Not bMontPesq Then
    bMontPesq = True
    sPesq = Mid$(sPesq, 1, Len(sPesq) - 1)
    dscMaster.MontarPesquisa sPesq
  End If
End Sub

