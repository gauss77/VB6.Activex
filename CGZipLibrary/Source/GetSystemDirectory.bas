Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Function GetSysDir() As String
  'KPD-Team 1998
  'URL: http://www.allapi.net/
  'E-Mail: KPDTeam@Allapi.net
  Dim sSave As String, Ret As Long
  'Create a buffer
  sSave = Space(255)
  'Get the system directory
  Ret = GetSystemDirectory(sSave, 255)
  'Remove all unnecessary chr$(0)'s
  sSave = Left$(sSave, Ret)
  
  GetSysDir = sSave
End Function
