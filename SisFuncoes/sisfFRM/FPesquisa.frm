VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Begin VB.Form FPesquisa 
   Caption         =   "Formulário de Pesquisa"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin Insignia_Frame.ideFrame panel 
      Align           =   2  'Align Bottom
      Height          =   1065
      Index           =   1
      Left            =   0
      Top             =   2655
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbOrdem 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   2700
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Executar"
         Height          =   330
         Left            =   4560
         TabIndex        =   2
         Top             =   645
         Width           =   1575
      End
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   645
         Width           =   4395
         _ExtentX        =   7752
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
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   315
         Index           =   1
         Left            =   2850
         TabIndex        =   5
         Top             =   270
         Width           =   1650
         _ExtentX        =   2910
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   315
         Index           =   2
         Left            =   4545
         TabIndex        =   6
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         FontSize        =   8,25
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Campo de Pesquisa"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data De"
         Height          =   195
         Index           =   1
         Left            =   2850
         TabIndex        =   8
         Top             =   60
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Até"
         Height          =   195
         Index           =   2
         Left            =   4545
         TabIndex        =   7
         Top             =   60
         Width           =   645
      End
   End
   Begin Insignia_Frame.ideFrame panel 
      Align           =   2  'Align Bottom
      Height          =   360
      Index           =   0
      Left            =   0
      Top             =   2295
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   635
      Caption         =   " Qtd. Registros:"
      CaptionAlign    =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check1 
         Caption         =   "Pesquisar valor em branco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3900
         TabIndex        =   1
         Top             =   75
         Width           =   2235
      End
   End
   Begin MSDataGridLib.DataGrid DTGrid 
      Height          =   2205
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   3889
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
Attribute VB_Name = "FPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private msSQLConsulta As String
'Private msSQLCount As String
'
'Public Property Let FormCaption(sCaption As String)
'  Me.Caption = sCaption
'End Property
'
'Public Sub Config(SQLConsulta As String, aCaptionDTFieldMaskWidth As String, sSQLCount As String)
''  msSQLConsulta = SQLConsulta
''  msSQLCount = sSQLCount
''
''  Dim sC() As String
''  Dim sI() As String
''  Dim i As Integer
''  sC = Split(aCaptionDTFieldMaskWidth, "|")
''  For i = 0 To UBound(sC)
''    sI = Split(sC(i), ",")
'''    DTGrid.Columns (i)
''  Next
''
''
''  Dim aCols() As String
''  Dim aCapW() As String
''  Dim i As Integer
''
''  aCols = Split(msCapWidthGrid, "|")
''  cmbOrdem.Clear
''  Set DataGrid1.DataSource = Nothing
''  Set DataGrid1.DataSource = mRS
''
''  For i = 0 To UBound(aCols)
''    aCapW = Split(aCols(i), ",")
''    DataGrid1.Columns(i).Caption = aCapW(0)
''    DataGrid1.Columns(i).Width = aCapW(1)
''    ReDim Preserve maDataField(i)
''    maDataField(i) = DataGrid1.Columns(i).DataField
''    cmbOrdem.AddItem aCapW(0)
''    cmbOrdem.ListIndex = 0
''  Next
'
'
'End Sub
'
''Retorna um array dos Valores dos Campos do registro selecionado
'Public Function ShowPesquisa() As String()
'  Me.Show vbModal
'End Function
