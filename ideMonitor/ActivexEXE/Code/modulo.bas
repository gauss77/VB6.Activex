Attribute VB_Name = "modulo"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -&H1
Private Const HWND_NOTOPMOST = -&H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private mObjetos()  As Variant
Private mbAlterou   As Boolean

Public Function AlwaysOnTop(pForm As Form, Optional TopMost As Boolean = True)
    AlwaysOnTop = SetWindowPos(pForm.hwnd, IIf(TopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Sub Main()
    If App.StartMode = vbSModeStandalone Then
        frmObjetos.Show
    End If
End Sub

Public Sub AddObjeto(sObjeto As String)
    Dim n As Integer
    Dim bExiste As Boolean
    
    On Error GoTo IniciaArray
    For n = 0 To UBound(mObjetos, 2)
        If mObjetos(1, n) = sObjeto Then
            bExiste = True
            mObjetos(2, n) = mObjetos(2, n) + 1
        End If
    Next
    On Error GoTo 0
    
    If UBound(mObjetos, 2) = 0 And mObjetos(1, 0) = "" Then
        mObjetos(1, 0) = sObjeto
        mObjetos(2, 0) = 1
    ElseIf bExiste = False Then
        ReDim Preserve mObjetos(1 To 2, UBound(mObjetos, 2) + 1)
        mObjetos(1, UBound(mObjetos, 2)) = sObjeto
        mObjetos(2, UBound(mObjetos, 2)) = 1
    End If
    
    mbAlterou = True
    
    Call frmObjetos.Atualizar(mObjetos)
    
    Exit Sub
    
IniciaArray:
    ReDim mObjetos(1 To 2, 0)
    Resume
End Sub

Public Sub RemoveObjeto(sObjeto As String)
    Dim n As Integer
    
    For n = 0 To UBound(mObjetos, 2)
        If mObjetos(1, n) = sObjeto Then
            mObjetos(2, n) = mObjetos(2, n) - 1
        End If
    Next
    mbAlterou = True
    
    Call frmObjetos.Atualizar(mObjetos)
End Sub

Public Sub ShowMonitor()
    Call frmObjetos.Atualizar(mObjetos)
    mbAlterou = False
    frmObjetos.Show
End Sub

Public Property Get ArrayObjetos() As Variant()
    ArrayObjetos = mObjetos
    mbAlterou = False
End Property

Public Property Get Alterou() As Boolean
    Alterou = mbAlterou
End Property
