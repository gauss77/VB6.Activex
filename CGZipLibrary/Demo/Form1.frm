VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup e Restore (ZipLibrary)"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3150
      TabIndex        =   9
      Top             =   2250
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1020
      Width           =   6375
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "..."
      Height          =   270
      Index           =   1
      Left            =   6540
      TabIndex        =   7
      Top             =   1665
      Width           =   300
   End
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   5160
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "..."
      Height          =   270
      Index           =   0
      Left            =   6540
      TabIndex        =   6
      Top             =   1035
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   1650
      Width           =   6375
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Fechar"
      Height          =   495
      Index           =   2
      Left            =   5655
      TabIndex        =   4
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Restaurar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4410
      TabIndex        =   0
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   90
      Top             =   2115
      Width           =   6810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   45
      Index           =   1
      Left            =   105
      Top             =   615
      Width           =   6765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Backup do Banco de Dados do Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   675
      TabIndex        =   3
      Top             =   285
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "Form1.frx":1042
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblCurDir 
      AutoSize        =   -1  'True
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   780
      Width           =   690
   End
   Begin VB.Label lblTempDir 
      AutoSize        =   -1  'True
      Caption         =   "Restaurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   1
      Top             =   1410
      Width           =   885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Unzip/Zip Client program for the CGZipLibrary ActiveXDLL
' Chris Eastwood, July 1999

Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub cmdButton_Click(Index As Integer)
  Select Case Index
    Case Is = 0
      Call Backup(Text1(0).Text)
    Case Is = 1
      Call Restore(Text1(1).Text)
    Case Is = 2
      Unload Me
  End Select
End Sub

Private Sub cmdDialog_Click(Index As Integer)
  Dim sTitle As String, sFilter As String
  
  Select Case Index
    Case Is = 0
      sFilter = "Microsoft Access (*.MDB,*.MDE)|*.mdb;*.mde"
      sTitle = "Seleção de Banco de Dados para Backup"
    Case Is = 1
      sFilter = "Microsoft Access (*.BKP)|*.bkp"
      sTitle = "Seleção de Arquivo para Restaurar"
  End Select
  
  With cmnDialog
    .CancelError = False
    .DialogTitle = sTitle
    .Filter = sFilter
    .ShowOpen
  
    Text1(Index).Text = .FileName
  End With
End Sub

Private Sub Restore(sPathFile As String)
  On Error GoTo vbErrorHandler
  
  If sPathFile = "" Then
    MsgBox "Selecione o Arquivo de Restauração...", vbInformation
    Text1(1).SetFocus
    Exit Sub
  End If
  
  Dim oSFunc As SisFuncoes.cSisFuncoes
  Set oSFunc = New cSisFuncoes
  
  Dim sDestino As String
  sDestino = oSFunc.ShowBrowseFolders(Me.hWnd, "Destino da Restauração")
  
  If sDestino = "" Then Exit Sub

  Dim oUnZip As CGUnzipFiles
  
  Set oUnZip = New CGUnzipFiles
    
  With oUnZip
    ' What Zip File ?
    .ZipFileName = sPathFile
    
    ' Where are we zipping to ?
    .ExtractDir = sDestino

    ' Keep Directory Structure of Zip ?
    .HonorDirectories = False
    
    ' Unzip and Display any errors as required
    If .Unzip <> 0 Then
      MsgBox .GetLastMessage
    End If
  End With
    
  Set oUnZip = Nothing
  MsgBox "Origem:  " & sPathFile & vbCrLf & _
         "Destino: " & sDestino, vbInformation, "Restauração Completa!"

  Exit Sub

vbErrorHandler:
  MsgBox Err.Number & " " & "Restauração" & " " & Err.Description

End Sub

Private Sub Backup(sPathFile As String)
  On Error GoTo vbErrorHandler
  
  If sPathFile = "" Then
    MsgBox "Selecione o Arquivo para Backup...", vbInformation
    Text1(0).SetFocus
    Exit Sub
  End If

  Dim sDestino As String
  
  With cmnDialog
    .CancelError = True
    .DialogTitle = "Salvar Arquivo de Bachup..."
    .Filter = "Arquivo de Backup(*.bkp)|*.bkp"
    .FileName = Format(Date, "dd-mm-yyyy") & ".bkp"
    On Error GoTo CancelSave
    .ShowSave
    
    sDestino = .FileName
  End With
  
  Dim oZip As CGZipFiles
   
  Set oZip = New CGZipFiles
    
  With oZip
    .ZipFileName = sDestino
    ' Are we updating a Zip File ?
    ' - This doesn't seem to work - check InfoZip homepage for more info.
    .UpdatingZip = False ' ensures a new zip is created

    ' Add in the files to the zip - in this case, we
    ' want all the ones in the current directory
    .AddFile sPathFile
    
    ' Make the zip file & display any errors
    If .MakeZipFile <> 0 Then
      MsgBox .GetLastMessage ' any errors
    End If
  End With
    
  Set oZip = Nothing
  MsgBox "Origem:  " & sPathFile & vbCrLf & _
         "Destino: " & sDestino, vbInformation, "Backup Completo!"
    
  Exit Sub

vbErrorHandler:
  MsgBox Err.Number & " " & "Backup" & " " & Err.Description
CancelSave:
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Function GetTempPathName() As String
  Dim sBuffer As String
  Dim lRet As Long
  
  sBuffer = String$(255, vbNullChar)
  lRet = GetTempPath(255, sBuffer)
  
  If lRet > 0 Then sBuffer = Left$(sBuffer, lRet)
  
  GetTempPathName = sBuffer
End Function

