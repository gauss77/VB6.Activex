VERSION 5.00
Begin VB.Form FPathDB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Caminho do Bando de Dados"
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   60
   ClientWidth     =   6660
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FCFCFC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   150
      TabIndex        =   1
      Top             =   465
      Width           =   6360
      Begin VB.ComboBox cmbODBC 
         Height          =   315
         ItemData        =   "FPathDB.frx":0000
         Left            =   75
         List            =   "FPathDB.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1635
         Width           =   6180
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   75
         TabIndex        =   5
         Top             =   1050
         Width           =   5880
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Path do Banco de Dados"
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   2160
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Driver ODBC"
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   3
         Top             =   1425
         Width           =   2160
      End
      Begin VB.ComboBox cmbProvedor 
         Height          =   315
         ItemData        =   "FPathDB.frx":00B4
         Left            =   75
         List            =   "FPathDB.frx":00C4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   6180
      End
      Begin VB.Label lblButton 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   2
         Left            =   6015
         TabIndex        =   10
         Top             =   1065
         Width           =   180
      End
      Begin VB.Shape shpButton 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   270
         Index           =   2
         Left            =   5955
         Top             =   1050
         Width           =   285
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Provedor"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   6
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   3
      Left            =   6360
      TabIndex        =   11
      Top             =   90
      Width           =   150
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancelar"
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   1
      Left            =   5220
      TabIndex        =   9
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Con&firmar"
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   0
      Left            =   3795
      TabIndex        =   8
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   6285
      Top             =   60
      Width           =   285
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   330
      Index           =   1
      Left            =   5205
      Top             =   2700
      Width           =   1350
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   330
      Index           =   0
      Left            =   3780
      Top             =   2700
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dados de Conexão do Banco de Dados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   2910
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   435
      Index           =   2
      Left            =   15
      Top             =   2655
      Width           =   6645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   405
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   6645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FCFCFC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   2265
      Index           =   1
      Left            =   15
      Top             =   405
      Width           =   6645
   End
End
Attribute VB_Name = "FPathDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbOK As Boolean

Private msSource As String
Private msProvedor As String

Public Sub Abrir(ByRef psProvedor As String, ByRef psSource As String)
  mbOK = True   'Para poder saber que esta sendo aberta do modo correto
  
  Me.Show vbModal
  
  psProvedor = msProvedor
  psSource = msSource
End Sub

Private Sub Form_Load()
  If mbOK = False Then
    MsgBox "Abrir esta janela, utilize o procedimento [Abrir]!", vbCritical, "Path do Banco de Dados"
    Unload Me
  Else
    With cmbProvedor
      .Clear
      
      .AddItem "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.4.0;"
      .AddItem "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
      .AddItem "PROVIDER=MSDataShape;Data PROVIDER=MSDASQL;"
      .AddItem "PROVIDER=MSDASQL;"
      .AddItem "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.ACE.OLEDB.12.0;"
      .AddItem "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
      
      .ListIndex = 0
    End With
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set FPathDB = Nothing
End Sub

Private Sub lblButton_Click(Index As Integer)
  Select Case Index
    Case Is = 0
      If Option1(0).Value = True Then
        msProvedor = cmbProvedor.Text
        msSource = txtPath.Text
      Else
        msProvedor = cmbODBC.Text
        msSource = msProvedor
      End If
      Unload Me
      
    Case Is = 1, 3
      msProvedor = ""
      msSource = ""
      
      Unload Me
    
    Case Is = 2
      Dim sFilter As String
      Dim sTitle As String
      Dim sDir As String
      
      sFilter = "Microsoft Access (*.MDB,*.MDE)" & Chr(0) & "*.mdb;*.mde"
      sTitle = "Abrir Banco de Dados Access"
      sDir = GetSetting(App.EXEName, "Paths", "PathDB", App.Path)
      
      txtPath.Text = mFuncoes.OpenFile(sFilter, sTitle, sDir)
            
      If txtPath.Text <> "" Then
        SaveSetting App.EXEName, "Paths", "PathDB", txtPath.Text
      End If
  End Select
End Sub
