Attribute VB_Name = "INI"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Sub AddToINI()
   Dim rc As Long
   rc = WritePrivateProfileString("Add-Ins32", "Monitor.Connect", _
        "0", "VBADDIN.INI")
   MsgBox "A entrada no arquivo VBADDIN.INI foi feita com sucesso!!!"
End Sub



