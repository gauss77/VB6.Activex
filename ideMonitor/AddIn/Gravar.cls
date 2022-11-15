VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gravar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ARQUIVO As String = "\" & "Monitor.log"

Public Function Log(sTipo As String, sProcName As String) As String
    Dim objFso As Scripting.FileSystemObject
    Dim objTexto As Scripting.TextStream
    Dim sLog As String

    Set objFso = New Scripting.FileSystemObject
    
    If objFso.FileExists(App.Path & ARQUIVO) Then
        Set objTexto = objFso.OpenTextFile(App.Path & ARQUIVO, ForAppending, True)
    Else
        Set objTexto = objFso.OpenTextFile(App.Path & ARQUIVO, ForWriting, True)
    End If
    
    'sLog = Contador(App.Path & sCaminhoArquivo) & ". " & sTipo & "  " & sProcName
    'objTexto.WriteLine sLog
    objTexto.WriteLine (Now & vbTab & sTipo & vbTab & sProcName)
    
    objTexto.Close
    
    Set objTexto = Nothing
    Set objFso = Nothing
    
    Log = App.Path & ARQUIVO
End Function

Private Function Contador(sCaminho As String) As Integer
   Dim objFso As Scripting.FileSystemObject
   Dim objArquivo As TextStream
   
   Set objFso = New Scripting.FileSystemObject
   Set objArquivo = objFso.OpenTextFile(sCaminho, ForReading)
   
   While Not objArquivo.AtEndOfStream
      Contador = objArquivo.Line
      objArquivo.ReadLine
   Wend
   
   Contador = Contador + 1
   
   objArquivo.Close
   
   Set objArquivo = Nothing
   Set objFso = Nothing
End Function
