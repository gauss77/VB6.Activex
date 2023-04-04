VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArqINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***** DECLARAÇÕES DAS API'S *************************
Private Declare Function GetPrivateProfileString Lib "kernel32" _
                         Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                         ByVal lpKeyName As Any, _
                         ByVal lpDefault As String, _
                         ByVal lpReturnedString As String, _
                         ByVal nSize As Long, _
                         ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
                         Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                         ByVal lpKeyName As Any, _
                         ByVal lpString As Any, _
                         ByVal lpFileName As String) As Long
'***************************

Private msPathFile As String
                                                 
Public Function Ler(ByVal sSeção As String, _
                    ByVal sChave As String, _
                    ByVal sDefault As String) As String
   
   Dim sTemp As String * 256
   Dim nLength As Integer
   
   sTemp = Space$(256)
   nLength = GetPrivateProfileString(sSeção, sChave, sDefault, sTemp, 255, msPathFile)
   Ler = Left$(sTemp, nLength)
End Function

Public Sub Gravar(ByVal sSeção As String, _
                  ByVal sChave As String, _
                  ByVal sValor As String)
   Dim n As Integer
                    
   n = WritePrivateProfileString(sSeção, sChave, sValor, msPathFile)
End Sub

'***** PROPRIEDADES **********************************
Public Property Let PathFile(ByVal sPathFile As String)
   msPathFile = sPathFile
End Property
