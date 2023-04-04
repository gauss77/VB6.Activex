Attribute VB_Name = "modExtractZipLib"
Option Explicit
Public Enum eZipFile
  zf_Zip32
  zf_UnZip32
End Enum
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

Public Sub ExtractZipFiles(eFile As eZipFile)
  Dim mFreeFile As Integer
  Dim FileTmp As String, Buffer As String
  Dim fZip As Variant, B As Long
  Dim sPathSys As String, sFileName As String
  
  mFreeFile = FreeFile
  FileTmp = "File.tmp~"
  
  If eFile = zf_Zip32 Then
    sFileName = "ZIP32.DLL"
  Else
    sFileName = "UNZIP32.DLL"

  End If

  fZip = LoadResData(sFileName, "Custom")
  
  Open FileTmp For Binary As mFreeFile
    Put mFreeFile, , fZip
  Close mFreeFile

  B = FileLen(FileTmp)
  Buffer = String(B - 12, " ")

  Open FileTmp For Binary As mFreeFile
    Seek mFreeFile, 13
    Get mFreeFile, , Buffer
  Close mFreeFile
  Kill FileTmp

  'Salvando mo arquivo correto
  sPathSys = GetSysDir & "\" & sFileName
  
  Open sPathSys For Binary As mFreeFile
    Put mFreeFile, , Buffer
  Close mFreeFile

End Sub

Public Sub CheckExistZipFiles()
  Dim sPathSys As String
  
  sPathSys = GetSysDir
  
  If FileExist(sPathSys & "\UNZIP32.DLL") = "" Then
    Call ExtractZipFiles(zf_UnZip32)
  End If
  
  If FileExist(sPathSys & "\ZIP32.DLL") = "" Then
    Call ExtractZipFiles(zf_Zip32)
  End If

End Sub

Public Function FileExist(PathFile As String) As String
  FileExist = Dir(PathFile, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)
End Function
