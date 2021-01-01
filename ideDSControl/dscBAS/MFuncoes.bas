Attribute VB_Name = "MFuncoes"
Option Explicit

Public gsParent As String

Global Const cAPPNome = "ideDSControl 2.8.2003"

Public Sub ErrRaise()
  MsgBox Err.Number & ":(" & Err.Description & ")", vbCritical, "Erro: " & App.FileDescription
  Err.Clear
  On Error GoTo 0
End Sub

Public Sub KeyDown(XDSource As Object, KeyCode As Integer, Shift As Integer)
  With XDSource
    If .Operacao <> opVisualizacao And Shift = 0 Then
      Select Case KeyCode
        Case Is = vbKeyF8: KeyCode = 0: .Update              'Confirmando Operacao
        Case Is = vbKeyF9: KeyCode = 0: .DataSource.Cancel   'Cancelando Operacao
        Case Else:        Exit Sub
      End Select
    Else  'Visualizacao
      If Shift = 0 Then
        Select Case KeyCode
          Case Is = vbKeyF3: KeyCode = 0:     .OperacaoPesquisar
          Case Is = vbKeyF4: KeyCode = 0:     .AddNewContinue
          Case Is = vbKeyF5: KeyCode = 0:     .AddNew                   'Adicionando um novo registro
          Case Is = vbKeyF6: KeyCode = 0:     .Edit
          Case Is = vbKeyF7: KeyCode = 0:     .Delete
          Case vbKeyEnd:     KeyCode = 0:     .DataSource.MoveLast
          Case vbKeyHome:    KeyCode = 0:     .DataSource.MoveFirst
          Case vbKeyPageUp:  KeyCode = 0:     .DataSource.MovePrevious
          Case vbKeyPageDown: KeyCode = 0:    .DataSource.MoveNext
        End Select
        
      ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF5 Then
        KeyCode = 0
        .DataSource.Requery          'Atualizando RecordSet
      End If
    End If
  End With
End Sub

Public Function ContadorWidth(ByRef pPanel As ideFrame, pOperacao As eDSOperacao, _
                              ByVal pRegAtual As Long, ByVal pRegCount As Long) As Integer
    Dim nFATOR As Long
    Dim NTOTAL As Long
    
    On Error Resume Next
    DoEvents
    pPanel.Caption = Format(pRegAtual, "000000") & " / " & Format(pRegCount, "000000")
    nFATOR = (pPanel.Width - 45) \ pRegCount
    
    If nFATOR = 0 Then nFATOR = 1
    
    If Err.Number <> 0 Then
        ContadorWidth = 0
        pPanel.Enabled = False
        Err.Number = 0
    Else
        If pOperacao = opVisualizacao Then
            pPanel.Enabled = True
        End If
        NTOTAL = nFATOR * pRegAtual
        If NTOTAL >= 80 Then ContadorWidth = (NTOTAL) - 80
    End If
    On Error GoTo 0
End Function

'=========== Guardar
''Trocar Nome para CódigoSequencial (Procedure nao esta sendo usada)
'Public Function CodigoAutomatico(ByVal psTabela, ByVal psCampo As String) As String
'  Dim Conn As New ADODB.Connection
'  Dim RS   As New ADODB.RecordSet
'  Dim sSQL As String
'
'  On Error GoTo ExibirErro
'  Conn.CursorLocation = adUseServer
'  Conn.Open msConnectionString
'
'  sSQL = "SELECT MAX(" & psCampo & ") as MaxCod From " & psTabela
'  Set RS = Conn.Execute(sSQL)
'  GoTo PularMsg
'
'ExibirErro:
'  MsgBox Err.Description & vbCrLf & Err.Source, vbCritical, "XDataSource_CodigoAutomatico"
'  Err.Clear
'  On Error GoTo 0
'  GoTo Destroy
'
'PularMsg:
'  If Not IsNull(RS!MaxCod) Then
'     CodigoAutomatico = Val(RS!MaxCod + 1)
'  Else
'     CodigoAutomatico = 1
'  End If
'
'Destroy:
'  On Error GoTo 0
'  On Error Resume Next
'  RS.Close
'  Conn.Close
'  On Error GoTo 0
'  Set RS = Nothing
'  Set Conn = Nothing
'End Function

'Public Property Get BookMark() As Long
'  On Error Resume Next
'  If mRSADO.RecordCount > 0 Then BookMark = mRSADO.BookMark
'End Property
'
'Public Property Let BookMark(ByVal vNewValue As Long)
'  If mRSADO.RecordCount > 0 Then
'    If vNewValue <> mRSADO.BookMark Then
'      mRSADO.BookMark = vNewValue
'    End If
'  End If
'End Property

'Public Sub Move(NumRecords As Long, Optional Start As Variant)
'  Dim nReg As Long
'
'  On Error Resume Next
'  nReg = mRSADO.RecordCount
'  If nReg = 1 Then
'    Exit Sub
'  End If
'  If nReg > 0 And _
'     nReg >= NumRecords Then
'    mRSADO.Move NumRecords, Start
'  End If
'  On Error GoTo 0
'End Sub
