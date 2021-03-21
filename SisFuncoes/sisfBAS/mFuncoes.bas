Attribute VB_Name = "mFuncoes"
Option Explicit

'Dialogo para Abertura de Arquivos
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

Public Function OpenFile(ByVal psFilter As String, _
                         Optional ByVal psTitle As String, _
                         Optional ByVal psInitialDir As String) As String
  
  Dim OFName As OPENFILENAME
  
  OFName.lStructSize = Len(OFName)

'  OFName.hwndOwner = Me.hWnd   'Set the parent window
  OFName.hInstance = App.hInstance 'Set the application's instance
 
  OFName.lpstrFile = Space$(254) 'create a buffer for the file
  OFName.nMaxFile = 255 'set the maximum length of a returned file
  OFName.lpstrFileTitle = Space$(254) 'Create a buffer for the file title
  OFName.nMaxFileTitle = 255 'Set the maximum length of a returned file title
  
  OFName.lpstrFilter = psFilter 'Set Filtro do Arquivos
  OFName.lpstrTitle = psTitle 'Set the title
  OFName.lpstrInitialDir = psInitialDir 'Set the initial directory
    
  OFName.flags = 0 'No flags
  
  'Show the 'Open File'-dialog
  If GetOpenFileName(OFName) Then
    OpenFile = Mid(Trim$(OFName.lpstrFile), 1, Len(Trim$(OFName.lpstrFile)) - 1)
  End If
End Function

Public Sub Reparar(psArquivo As String)
  Dim sMsg As String
  
  If MsgBox("Compactação e Correção do Banco de Dados pode demorar alguns minutos!" & _
            vbCrLf & "Deseja continuar?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  On Error Resume Next
  'repara o Banco de Dados
  DBEngine.RepairDatabase (psArquivo)
    
  If Err.Number <> 0 Then
    sMsg = "Corrigir: " & Err.Description & vbCrLf
    Err.Clear
  End If
  

  On Error GoTo ErrCompact
  'compacta banco de dados e renomeia
  DBEngine.CompactDatabase psArquivo, psArquivo & "_Compact"
    
  'apaga o BD antigo
  Kill psArquivo
  
  'renomeia o Banco de Dados
  Name psArquivo & "_Compact" As psArquivo
  
  MsgBox "Compactação e Correção de Banco de Dados encerrada!", vbInformation, "Compactar e Corrigir DataBase"
  Exit Sub

ErrCompact:
  sMsg = sMsg & "Compactar: " & Err.Description

  If sMsg <> "" Then
    MsgBox sMsg, vbCritical, "Erros Ocorridos!"
  End If
End Sub

