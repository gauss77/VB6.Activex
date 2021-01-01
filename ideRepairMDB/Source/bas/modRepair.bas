Attribute VB_Name = "modRepair"
Option Explicit

Public Const STR_FILTER As String = "Microsoft Access (*.MDB,*.MDE)|*.mdb;*.mde"
Public Const STR_TITLE As String = "Abrir Arquivo - Compactar e Corrigir [.MDB] "
Public Const STR_INITIAL_PATH As String = "C:\"

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Sub CompactMDB(Optional ByVal psPathFile As String = "")
    If psPathFile = "" Then
        psPathFile = OpenFile(STR_FILTER, STR_TITLE, App.Path)
        
        If psPathFile = "" Then Exit Sub 'Cancelado
    End If
    
    Call Repair(psPathFile)
End Sub

Public Function OpenFile(ByVal psFilter As String, _
                         Optional ByVal psTitle As String, _
                         Optional ByVal psInitialDir As String) As String
  
    Dim OFName As OPENFILENAME
    
    With OFName
        .lStructSize = Len(OFName)
'        .hwndOwner = Me.hWnd            'Set the parent window
        .hInstance = App.hInstance      'Set the application's instance
        .lpstrFile = Space$(254)        'create a buffer for the file
        .nMaxFile = 255                 'set the maximum length of a returned file
        .lpstrFileTitle = Space$(254)   'Create a buffer for the file title
        .nMaxFileTitle = 255            'Set the maximum length of a returned file title
        .flags = 0                      'No flags
        
        .lpstrInitialDir = psInitialDir 'Set the initial directory
        .lpstrFilter = psFilter
        .lpstrTitle = psTitle           'Set the title
    End With
    
    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        OpenFile = Trim$(OFName.lpstrFile)
        'OpenFile = Mid(Trim$(OFName.lpstrFile), 1, Len(Trim$(OFName.lpstrFile)) - 1)
    End If
  
End Function

Private Sub Repair(psPathFile As String)
    Dim sMsg As String
    
    DoEvents
'    If MsgBox("Compactação e Correção do Banco de Dados pode demorar varios minutos dependendo do tamanho do arquivo!" & _
'              vbCrLf & "Deseja continuar?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
    
    On Error GoTo ErrRepair
    'repara o Banco de Dados
    DBEngine.RepairDatabase (psPathFile)
    DoEvents
    GoTo Compact
  
ErrRepair:
  sMsg = "Erro ao Corrigir DataBase." & vbCrLf
  
Compact:
    On Error GoTo ErrCompact
    'compacta banco de dados e renomeia
    DBEngine.CompactDatabase psPathFile, psPathFile & "_Compact"
    DoEvents
    'apaga o BD antigo
    Kill psPathFile
    
    'renomeia o Banco de Dados
    Name psPathFile & "_Compact" As psPathFile
    
    MsgBox "Compactação e Correção de Banco de Dados encerrada!", vbInformation, "Compactar e Corrigir DataBase"
    Exit Sub

ErrCompact:
    sMsg = sMsg & "Erro ao Compactar DataBase."
    
    If sMsg <> "" Then
        MsgBox sMsg, , "Compactar e Corrigir DataBase"
    End If
End Sub

