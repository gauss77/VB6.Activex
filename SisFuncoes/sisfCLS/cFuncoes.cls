VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSisFuncoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private moArqINI As cArqINI

Public Event Printer(ByVal NameReport As String, ByVal View As Boolean)
'Public Event FindValores(ID As Long, Name As String, Index As Integer)
'Public Event FindValColunas(Col As Columns)

Private WithEvents mFPrinterRel As FPrinterRel
Attribute mFPrinterRel.VB_VarHelpID = -1

Public Sub PathDB(ByRef psProvedor As String, psSource As String)
  Call FPathDB.Abrir(psProvedor, psSource)
End Sub

Public Property Get ArqINI() As cArqINI
  If moArqINI Is Nothing Then
    Set moArqINI = New cArqINI
  End If
  
  Set ArqINI = moArqINI
End Property

Private Sub Class_Terminate()
  Set moArqINI = Nothing
End Sub

Public Sub ShowLOpcoes(ByVal pArrayKeyDesc As String, _
                       ByRef psRetKey As String, ByRef psRetDescricao As String)

  FOpcoes.ShowLOpcoes pArrayKeyDesc, psRetKey, psRetDescricao
End Sub

Public Sub ShowLRelatorios(ByVal ListArray As String)
  If (mFPrinterRel Is Nothing) Then Set mFPrinterRel = FPrinterRel
  
  mFPrinterRel.ShowLRelatorios ListArray
End Sub

Public Sub CompactMDB(Optional ByVal psPathBanco As String = "")
  Dim sFilter As String
  Dim sTitle As String
  
  If psPathBanco = "" Then
    sFilter = "Microsoft Access (*.MDB,*.MDE)" & Chr(0) & "*.mdb;*.mde"
    sTitle = "Abrir Arquivo - Compactar e Corrigir [.MDB]"
    
    psPathBanco = mFuncoes.OpenFile(sFilter, sTitle, App.Path)
    
    If psPathBanco = "" Then Exit Sub 'Cancelado
  
  End If
  
  Call Reparar(psPathBanco)
End Sub

Public Function ShowBrowseFolders(hOwnerWnd As Long, Optional ByVal sInstruct As String, Optional vSelPath As Variant, Optional vTopFolder As Variant) As String
  ShowBrowseFolders = BrowseFolders(hOwnerWnd, sInstruct, vSelPath, vTopFolder)
End Function

Private Sub mFPrinterRel_Printer(ByVal NameReport As String, ByVal View As Boolean)
  RaiseEvent Printer(NameReport, View)
End Sub
