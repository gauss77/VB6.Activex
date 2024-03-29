VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDSControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mRsADO As ADODB.Recordset
Attribute mRsADO.VB_VarHelpID = -1
Private mRsFieldClone As ADODB.Recordset

Private msParentEXEName As String

Private meOperacao    As eDSOperacao
Private mbCancelEvent As Boolean
Private mbFlagRequery As Boolean

Public Event MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Public Event FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Public Event Operacao(ByVal eOperacao As eDSOperacao, ByVal eOperacaoAnterior As eDSOperacao)

Public Event AntesAddNew(Cancel As Boolean)
Public Event AntesCancel(Cancel As Boolean)
Public Event AntesUpdate(Cancel As Boolean, eOperacao As eDSOperacao)
Public Event DepoisUpdate(eOperacao As eDSOperacao)
Public Event AntesRequery(Cancel As Boolean)
Public Event DepoisRequery()

Public Event RecordDeleted()

Private mbDisableRaiseEventFieldChange As Boolean

Public Sub DesConectar()
    Debug.Print ("Insignia_DSControl.CDSControl.DesConectar")
  On Error Resume Next
  
  If Not mRsADO Is Nothing Then
    mRsADO.Close
    Set mRsADO = Nothing
  End If
  
  On Error GoTo 0
  Err.Clear
End Sub

Friend Sub Class_GetDataMember(DataMember As String, Data As Object)
    Debug.Print ("Insignia_DSControl.CDSControl.Class_GetDataMember")
  ' TODO:  Return the appropriate recordset based on DataMember. For example:

  If mbFlagRequery Then Exit Sub

  Select Case DataMember
    Case Is = ""
      Set Data = mRsADO
    Case Else
      Set Data = mRsADO(DataMember).UnderlyingValue
  End Select
End Sub

Private Sub Class_Initialize()
    Debug.Print ("Insignia_DSControl.CDSControl.Class_Initialize")

End Sub

Private Sub Class_Terminate()
    Debug.Print ("Insignia_DSControl.CDSControl.Class_Terminate")

End Sub

Private Sub mRsADO_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print ("Insignia_DSControl.CDSControl.mRsADO_FieldChangeComplete")
 ' verifica dados invalidos inseridos pelo usuario
 ' este evento sera disparado quando o campo que esta sendo editado
 ' sera disparado novamente depois do metodo update ser invocado
 If mbDisableRaiseEventFieldChange Then Exit Sub
 
 RaiseEvent FieldChangeComplete(cFields, Fields, pError, adStatus, pRecordset)
End Sub

Private Sub mRsADO_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print ("Insignia_DSControl.CDSControl.mRsADO_MoveComplete")
  Static bDisparado As Boolean
  
  'Esta passando duas vezes qdo usa o controle xdsmaster
  If Not bDisparado And meOperacao = opVisualizacao Then
    bDisparado = True
    If Not IsRecordDeleted Then RaiseEvent MoveComplete(adReason, pError, adStatus, pRecordset)
    bDisparado = False
  End If

End Sub

Public Property Get Connection() As ADODB.Connection
  On Error GoTo TrataErro
  If mRsADO Is Nothing Then Exit Property
  Set Connection = mRsADO
  Exit Property
TrataErro:
  Call MFuncoes.ErrRaise
End Property

Public Property Get RS() As ADODB.Recordset
  Set RS = mRsADO
End Property

Public Property Set RS(ByVal pRsADO As ADODB.Recordset)
  Set mRsADO = pRsADO
End Property


Public Function Conectar(ByVal Source As String, _
                         ByRef ActiveConnection As Variant, _
                         Optional ByVal pCursorLocation As CursorLocationEnum = adUseClient, _
                         Optional ByVal pCursorType As CursorTypeEnum = adOpenKeyset, _
                         Optional ByVal pLockType As LockTypeEnum = adLockPessimistic) As eDSConexao
    Debug.Print ("Insignia_DSControl.CDSControl.Conectar")
  
  On Error GoTo ErrHandle
  
  If mRsADO Is Nothing Then
    Set mRsADO = New ADODB.Recordset
  Else
    If MsgBox("Base de Dados j� se encontra conectada!" & vbCrLf & _
               "Deseja Re-Conectar?", vbQuestion + vbYesNo, cAPPNome) = vbYes Then
      Call DesConectar
      Set mRsADO = New ADODB.Recordset
    Else
      Exit Function
    End If
  End If
  
  With mRsADO
    .CursorLocation = pCursorLocation
    .Open Source, ActiveConnection, pCursorType, pLockType
    
    If .RecordCount = 0 Then
      Conectar = cnVazio
    Else
      Conectar = cnSucesso
    End If
  End With
  
  Exit Function

ErrHandle:
  Set mRsADO = Nothing
  Conectar = cnErroProcesso
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") na procedure CDSControl.Conectar"
  Err.Clear
End Function

Public Property Get EditingRecord() As Boolean
  If Not mRsADO Is Nothing Then _
    EditingRecord = (mRsADO.EditMode <> adEditNone)
End Property

Public Property Get AbsolutePosition() As Long
  On Error Resume Next
  AbsolutePosition = mRsADO.AbsolutePosition
  On Error GoTo 0
End Property
Public Property Let AbsolutePosition(vNewValue As Long)
    Debug.Print ("Insignia_DSControl.CDSControl.Property Let AbsolutePosition")
  On Error Resume Next
  mRsADO.AbsolutePosition = vNewValue
  On Error GoTo 0
End Property

Public Sub AddNew(Optional ByVal FieldList, Optional ByVal Values)
    Debug.Print ("Insignia_DSControl.CDSControl.AddNew")
  On Error GoTo AddErr
  
  If meOperacao = opVisualizacao Then
    mbCancelEvent = False
    RaiseEvent AntesAddNew(mbCancelEvent)
    If mbCancelEvent = True Then Exit Sub
    
    If Not IsMissing(FieldList) Then
      mRsADO.AddNew FieldList, Values
    Else
      mRsADO.AddNew
    End If
    
    Operacao = opInclusao
  End If
  
  Exit Sub
AddErr:
  Call MFuncoes.ErrRaise
End Sub

Public Function Delete(Optional ByVal ShowMessage As Boolean = True) As Boolean
    Debug.Print ("Insignia_DSControl.CDSControl.Delete")
  
  If ShowMessage Then _
      If MsgBox("O registro ser� EXCLU�DO definitivamente." & _
                vbCrLf & vbCrLf & "Deseja Continua ?", _
                vbYesNo + vbExclamation, _
                "Confirmar Exclus�o") <> vbYes Then Exit Function
  
  Dim nTentativas As Integer
GoInicio:
  Err.Clear
  
  On Error Resume Next
  mRsADO.Delete adAffectCurrent
    
  If Err.Number <> 0 And nTentativas < 2 Then
    nTentativas = nTentativas + 1
    mRsADO.CancelUpdate
    Me.Resync nTentativas < 2
    GoTo GoInicio
  End If
  
  Select Case Err.Number
    Case 0: 'exclusao foi um sucesso
      Call MoveNext
      Delete = True
      
    Case -2147217864
    Case -2147217885
'      MsgBox "Este registro foi exclu�do ou est� sendo exclu�do por outro usu�rio!", vbInformation
'      mRsADO.CancelUpdate
'    Call MoveNext
'    Delete = True
      Me.Operacao = opVisualizacao
      RaiseEvent RecordDeleted
    
    Case -2147467259
      MsgBox "As altera��es realizadas n�o podem ser gravadas no momento." _
             & "O registro encontra-se bloqueado por outro usu�rio." _
             & vbCrLf & "Voc� pode cancelar as altera��es ou tentar salvar mais tarde...", _
             vbExclamation, "Erro de grava��o"
      Exit Function
    Case Else
      MsgBox "O registro n�o pode ser exclu�do." + vbCrLf + vbCrLf + Err.Description
      mRsADO.CancelUpdate
  End Select
  On Error GoTo 0
End Function

Public Sub Requery(Optional ByVal Options As Long = -1)
    Debug.Print ("Insignia_DSControl.CDSControl.Requery")
  Dim nPos As Long
  
  On Error GoTo RefreshErr

  mbCancelEvent = False
  RaiseEvent AntesRequery(mbCancelEvent)
  If mbCancelEvent Then Exit Sub
    
  mbFlagRequery = True
  
  Screen.MousePointer = vbHourglass
  With mRsADO
    nPos = .AbsolutePosition
    .Requery Options
    .AbsolutePosition = nPos
  End With
  Screen.MousePointer = vbDefault
  
  mbFlagRequery = False
  
'  For i = 1 To DataMembers.Count
'    DataMemberChanged DataMembers(i)
'  Next

  RaiseEvent DepoisRequery

  Exit Sub
  
RefreshErr:
  mbFlagRequery = False
  Screen.MousePointer = vbDefault
  Call MFuncoes.ErrRaise
End Sub

Public Function Update() As Boolean
    Debug.Print ("Insignia_DSControl.CDSControl.Update")
  Dim bResult As Boolean
  Dim eOpRetorno As eDSOperacao
  Dim bEditAdd As Boolean
  Dim vErr As MFuncoes.tpError
  
  mbCancelEvent = False
  eOpRetorno = meOperacao
  RaiseEvent AntesUpdate(mbCancelEvent, meOperacao)
  If mbCancelEvent Then Exit Function
  
  Dim vFieldArray(), x As Integer
  
  On Error Resume Next 'limpa o objeto error
  
  bEditAdd = mRsADO.EditMode = adEditAdd
  With mRsADO
    .ActiveConnection.Errors.Clear
  
    Screen.MousePointer = vbHourglass
    .AbsolutePosition = .AbsolutePosition   'Isso for�a a gravacao
'    If .Supports(adUpdateBatch) Then
'      .UpdateBatch
'    Else
      .Update
'    End If
    Screen.MousePointer = vbNormal
  
    'Caso tenha ocorrido o erro na linha cima,
    'guarda o numero do erro(O objeto error sera resetado pela proxima linha)
    vErr = MFuncoes.CopyErrObject
    
    On Error GoTo TrataErro
    Select Case vErr.Number
      Case 0:
        If .ActiveConnection.Errors.Count = 0 Then 'nao ocorreu nenhum erro
          If bEditAdd Then
            Me.Resync 'Atualiza o Registro Atual
'            .Move 0 'forca uma atualiacao dos controles para exibir os dados
          End If
          
          Me.Operacao = opVisualizacao
          bResult = True
          If Not mRsFieldClone Is Nothing Then
            If mRsFieldClone.State = adStateOpen Then mRsFieldClone.Close
            Set mRsFieldClone = Nothing
          End If
        Else
          Dim oError As Error, sErros As String
          
          For Each oError In .ActiveConnection.Errors
            If oError.Number = -2147217864 Then
              sErros = "Este registro j� foi recentemente alterado por outro usu�rio! "
              Me.Operacao = opVisualizacao
            Else
              sErros = sErros & oError.Description & vbCr
            End If
          Next
          Set oError = Nothing
          
          If sErros <> "" Then
            sErros = "Os seguintes erros foi(ram) " _
                     & "definidos pelo provedor : " & vbCr & sErros
            MsgBox sErros, vbCritical, App.FileDescription ' exibe todos os erros
          End If
        End If
        
      Case 3640 + vbObjectError 'registro alterado por outro usuario
        If MsgBox("Este registro foi alterado enquanto voc� realizava sua altera��o." _
                  & vbCr & vbCr & "Se continuar voc� ir� substituir todas as altera��es do outro usu�rio." _
                  & vbCr & vbCr & "Deseja sobrescrever as altera��es do outro usu�rio?" _
                  , vbExclamation + vbYesNoCancel, "Conflito na Informa��es") = vbYes Then
          'forca uma sobrescrita dos dados devemos armazenar os dados em um buffer
          ReDim vFieldArray(.Fields.Count - 1)
          For x = 0 To .Fields.Count - 1
            vFieldArray(x) = .Fields(x).Value
          Next
              
          .CancelUpdate
          If Me.Resync Then
            For x = 0 To .Fields.Count - 1 'salva as alteracoes no banco de dados
              If .Fields(x).Value <> vFieldArray(x) Then .Fields(x) = vFieldArray(x)
            Next
                
            .Update
            
          End If
          Me.Operacao = opVisualizacao

        Else 'usuario escolheu nao sobrescrever os dados
          .CancelUpdate
          Me.Resync
          .Move 0
        End If
        
      Case -2147467259
        If MFuncoes.IsDebug Then
          MsgBox "Number: " & vbTab & vErr.Number & vbCrLf _
               & "Description: " & vbTab & vErr.Description & vbCrLf _
               , vbError, "Origem: CDSControl.Update"
        Else
        MsgBox "Suas altera��es n�o podem ser gravadas no momento! " & vbCrLf _
               & vbCr & "O registro encontra-se bloqueado por outro usu�rio." _
               & vbCr & "Voc� pode cancelar as altera��es ou tentar salvar mais tarde...", vbExclamation, "Erro de grava��o"
        End If
        Exit Function
      Case Else:
        MsgBox "Erro: " & vbTab & vErr.Number & vbCrLf _
               & "Descri��o: " & vbTab & vErr.Description & vbCrLf _
               & "Origem: CDSControl.Update", vbError, AppName
        Me.Operacao = opVisualizacao
    End Select
  End With
  
  If bResult Then RaiseEvent DepoisUpdate(eOpRetorno)
    
  Update = bResult
  Exit Function
    
TrataErro:
  Call MFuncoes.ErrRaise
End Function

Public Sub Cancel()
    Debug.Print ("Insignia_DSControl.CDSControl.Cancel")
  mbCancelEvent = False
  RaiseEvent AntesCancel(mbCancelEvent)
  If mbCancelEvent = True Then Exit Sub
  
  On Error GoTo ErrCancel
  With mRsADO
    If Not .EOF And Not .BOF Then
      Call ReValueFields
      .CancelUpdate
    End If
'    If .EditMode = adEditAdd And (Not .EOF And Not .BOF) Then .MoveLast
  End With

  Me.Operacao = opVisualizacao
  Exit Sub
  
ErrCancel:
  Call MFuncoes.ErrRaise
End Sub

Public Sub MoveFirst()
    Debug.Print ("Insignia_DSControl.CDSControl.MoveFirst")
  On Error GoTo TrataErro:
  If mRsADO.RecordCount > 0 Then mRsADO.MoveFirst
  Exit Sub
  
TrataErro:
  Call MFuncoes.ErrRaise
End Sub

Public Sub MoveLast()
    Debug.Print ("Insignia_DSControl.CDSControl.MoveLast")
  On Error GoTo TrataErro
  mRsADO.MoveLast
  Exit Sub
TrataErro:
  Call MFuncoes.ErrRaise
End Sub

Public Sub MoveNext()
    Debug.Print ("Insignia_DSControl.CDSControl.MoveNext")
  On Error GoTo GoNextError
  If Not mRsADO.EOF And mRsADO.RecordCount > 0 Then mRsADO.MoveNext
  If mRsADO.EOF And mRsADO.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    mRsADO.MoveLast
  End If
  Exit Sub

GoNextError:
  Call MFuncoes.ErrRaise
End Sub

Public Sub MovePrevious()
    Debug.Print ("Insignia_DSControl.CDSControl.MovePrevious")
  On Error GoTo TrataErro
  If Not mRsADO.BOF Then mRsADO.MovePrevious
  If mRsADO.BOF And mRsADO.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    mRsADO.MoveFirst
  End If
  Exit Sub

TrataErro:
  Call MFuncoes.ErrRaise
End Sub

Public Sub AutoSearch(DataField As String, Texto As String, KeyAscii As Integer)
    Debug.Print ("Insignia_DSControl.CDSControl.AutoSearch")
  Dim sSearch As String, sTexto As String
  Dim nPos As Long
  
  If mRsADO Is Nothing Then _
      MsgBox "Banco de Dados n�o est� conectado.", vbInformation, App.FileDescription
      
  If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then Exit Sub

  On Error Resume Next
  If KeyAscii = vbKeyBack Then
    sTexto = Mid$(Texto, 1, Len(Texto) - 1)
  Else
    sTexto = Texto & Chr(KeyAscii)
  End If
  On Error GoTo 0
  
  If sTexto = "" Then Exit Sub
  
  sSearch = "[" & DataField & "] LIKE '" & sTexto & "%'"
  
  On Error Resume Next
  nPos = mRsADO.AbsolutePosition
  
  mRsADO.Find "[" & DataField & "] LIKE '" & sTexto & "%'", , adSearchForward, 1
  If Err.Number = -2147217825 Then 'Erro de Operador
    mRsADO.Find "[" & DataField & "] = '" & sTexto & "'", , adSearchForward, 1
    Err.Clear
  End If
  
  If mRsADO.EOF Or mRsADO.BOF Then mRsADO.AbsolutePosition = nPos
  On Error GoTo 0
End Sub

Public Sub Search(ByVal DataField As String, ByVal Texto As String)
    Debug.Print ("Insignia_DSControl.CDSControl.Search")
  Dim sTexto      As String
  Dim sCriterio   As String
  
  Dim RSConsulta  As ADODB.Recordset
  
  Static sStaticField As String
  Static sStaticCriterio As String
  
  'Preparando texto de busca p/ indiferenca de acento
  sTexto = Texto
  If sTexto = "" Then Exit Sub
  
  On Error GoTo TrataErro
  Set RSConsulta = mRsADO.Clone(adLockReadOnly)
  With RSConsulta
    If mRsADO.Sort <> "" Then .Sort = mRsADO.Sort
    
    DataField = Replace(DataField, "[", "") 'Apenas para evitar erros
    DataField = Replace(DataField, "]", "")
    
    Select Case RSConsulta.Fields(DataField).Type
      Case Is = adNumeric, adInteger, adSmallInt, adDouble, adBinary
        sCriterio = " = " & sTexto
        
      Case Is = adBoolean
        Dim s As String
        
        sTexto = Trim$(sTexto)
        s = UCase(Left$(sTexto, 1))
        
        Select Case s
          Case Is = "S", "V", "-", "T" 'S=Sim , V=Verdadeiro, - = -1, T = True
            sTexto = "True"
          Case Else
            sTexto = "False"
        End Select
        sCriterio = " = " & sTexto
        
      Case Is = adDate, adDBDate, adDBTime
        sCriterio = " = #" & sTexto & "#"
        
  '    Case Is = adChapter, adWChar, adVarChar, adVarWChar, adChar
   '     sCriterio = " LIKE '" & sTexto & "%'"
      
      Case Else
        sCriterio = " LIKE '" & sTexto & "%'"
    End Select
    
    DataField = "[" & DataField & "]"
    
    Dim nStartFind As Integer
    
    If sStaticField <> DataField Then
      sStaticField = DataField
      nStartFind = 0
    Else
      If mRsADO.AbsolutePosition < mRsADO.RecordCount Then
        nStartFind = mRsADO.AbsolutePosition
      Else
        nStartFind = 0
      End If
    End If
    
    If sStaticCriterio <> sCriterio Then
      sStaticCriterio = sCriterio
      nStartFind = 0
    Else
      If mRsADO.AbsolutePosition < mRsADO.RecordCount Then
        nStartFind = mRsADO.AbsolutePosition
      Else
        nStartFind = 0
      End If
    End If
    
    On Error GoTo 0
    
    On Error Resume Next
    If nStartFind > 0 Then
      .Find DataField & sCriterio, nStartFind, adSearchForward
    Else
      'Neste caso inicia a busca a partir do primeiro registro
      .Find DataField & sCriterio, , adSearchForward, 1
    End If
    
    If Err.Number <> 0 Then
      GoTo TrataErro:
    End If
  
    If Not .BOF And Not .EOF Then
      mRsADO.AbsolutePosition = .AbsolutePosition
    End If
  End With
  
  RSConsulta.Close
  On Error GoTo 0

  Set RSConsulta = Nothing
  
  Exit Sub
  
TrataErro:
  Call MFuncoes.ErrRaise

  On Error Resume Next
  RSConsulta.Close
  Set RSConsulta = Nothing
  On Error GoTo 0
End Sub

Public Property Get Filter() As Variant
  On Error GoTo ErrRaise
  Filter = mRsADO.Filter
  Exit Property
ErrRaise:
  Call MFuncoes.ErrRaise

End Property

Public Property Let Filter(ByVal sFilter As Variant)
    Debug.Print ("Insignia_DSControl.CDSControl.Property Let Filter")
  On Error GoTo ErrRaise
  mRsADO.Filter = sFilter
  Exit Property
ErrRaise:
  Call MFuncoes.ErrRaise
End Property

Public Property Get Field(Index) As ADODB.Field
  On Error GoTo TrataErro
  Set Field = mRsADO(Index)
Exit Property
TrataErro:
  Call MFuncoes.ErrRaise
End Property

Public Sub Find(Criterio As String, Optional SkipRecords As Long, _
                Optional SearchDirection As SearchDirectionEnum = adSearchForward, Optional Start)
    Debug.Print ("Insignia_DSControl.CDSControl.Find")
  On Error GoTo TrataErro
  mRsADO.Find Criterio, SkipRecords, SearchDirection, Start
  Exit Sub
TrataErro:
  Call MFuncoes.ErrRaise
End Sub

Public Sub SeekValue(ByVal pIndexName As String, KeyValues, Optional SeekOption As SeekEnum = adSeekFirstEQ)
    Debug.Print ("Insignia_DSControl.CDSControl.SeekValue")
  On Error GoTo ErrRaise
  If (mRsADO.Supports(adIndex) And mRsADO.Supports(adSeek)) Then
      mRsADO.Index = pIndexName
      mRsADO.Seek KeyValues, SeekOption
  Else
    MsgBox "Provedor de banco de dados n�o oferece suporte para fun��o [Seek]", App.CompanyName
  End If
  Exit Sub
ErrRaise:
  Call MFuncoes.ErrRaise
End Sub

Public Property Let Sort(ByVal sField As String)
    Debug.Print ("Insignia_DSControl.CDSControl.Property Let Sort")
  On Error GoTo TrataErro
  mRsADO.Sort = sField
  Exit Property
  
TrataErro:
  If Err.Number = 3265 Then
    MsgBox "Sistema: " & Me.AppName & vbCrLf & vbCrLf & _
           "Numero do Erro: 3265 " & vbCrLf & _
           "Descri��o: O item n�o pode ser encontrado na cole��o " & vbCrLf & _
           "           correspondente ao nome " & UCase(sField) & vbCrLf & vbCrLf & _
           "Consulte o Administrador de Sistema e o informe sobre o erro!", vbInformation, Me.AppName
    Err.Clear
  Else
    Call MFuncoes.ErrRaise
  End If
End Property

Public Property Get AppName() As String
  If msParentEXEName = "" Then msParentEXEName = gsParent
  
  If msParentEXEName = "" Then
    msParentEXEName = App.FileDescription
    gsParent = App.FileDescription
  End If
  
  AppName = msParentEXEName
End Property
Public Property Let AppName(ByVal sEXEName As String)
    Debug.Print ("Insignia_DSControl.CDSControl.Property Let AppName")
  msParentEXEName = sEXEName
End Property

Public Property Get Operacao() As eDSOperacao
  Operacao = meOperacao
End Property
Public Property Let Operacao(ByVal vNewValue As eDSOperacao)
    Debug.Print ("Insignia_DSControl.CDSControl.Property Let Operacao")
  RaiseEvent Operacao(vNewValue, meOperacao)
  meOperacao = vNewValue
End Property

Public Function OperacaoPendente() As Boolean
    Debug.Print ("Insignia_DSControl.CDSControl.OperacaoPendente")
  Dim sCap As String
  
  If gsParent = "" Then
    sCap = "na Janela " & UCase(gsParent)
  Else
    sCap = "na Janela Ativa"
  End If
  
  If meOperacao <> opVisualizacao Then
    Beep
    MsgBox "Existe uma opera��o pendente " & sCap & "." & Chr$(13) & _
           "CONFIRME ou CANCELE a opera��o antes de proceguir...", vbExclamation, "ATEN��O"
    OperacaoPendente = True
  End If
End Function

Friend Function Resync(Optional hideError As Boolean = False) As Boolean
  Debug.Print ("Insignia_DSControl.CDSControl.Resync")
  'Atualiza o Registro Atual funciona apenas em tabelas com chave primaria
'  On Error Resume Next
'  If mRsADO.Supports(adResync) Then mRsADO.Resync adAffectCurrent
'  On Error GoTo 0
  
  On Error GoTo Finaliza
    DoEvents
    If mRsADO.Supports(adResync) Then mRsADO.Resync adAffectCurrent
    Resync = True
    
Finaliza:
    If hideError = False And Err.Number <> 0 Then
      If Err.Number = -2147217885 Or Err.Number = -2147217864 Then
        MsgBox "Este registro foi exclu�do ou est� sendo exclu�do por outro usu�rio!" & vbCrLf & vbCrLf & _
               "Por favor, abra novamente essa tela para atualizar seus dados", vbCritical
        
        Me.Operacao = opVisualizacao
        RaiseEvent RecordDeleted
      Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Resync of Module de classe CDSControl", vbCritical
      End If
    End If
End Function

Friend Sub CloneRegFields()
    Debug.Print ("Insignia_DSControl.CDSControl.CloneRegFields")
  Dim fld As ADODB.Field

  Set mRsFieldClone = New ADODB.Recordset
    
  'Cria a estrutura dos campos
  For Each fld In mRsADO.Fields
    If Not fld.Properties("ISAUTOINCREMENT") Then
      mRsFieldClone.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
      If fld.Precision > 0 Then
          mRsFieldClone.Fields(fld.Name).Precision = fld.Precision
      End If
      If fld.NumericScale > 0 Then
          mRsFieldClone.Fields(fld.Name).NumericScale = fld.NumericScale
      End If
    End If
  Next
  
  'Adiciona os valores
  mRsFieldClone.CursorLocation = adUseServer
  mRsFieldClone.Open , , adOpenKeyset
  mRsFieldClone.AddNew
  
  For Each fld In mRsADO.Fields
    If Not fld.Properties("ISAUTOINCREMENT") Then
      mRsFieldClone.Fields(fld.Name).Value = fld.Value
    End If
  Next
  
End Sub

'Volta os valores originais apos cancelar edicao
Friend Sub ReValueFields()
    Debug.Print ("Insignia_DSControl.CDSControl.ReValueFields")
  Dim fld As ADODB.Field
  
  If mRsFieldClone Is Nothing Or mRsADO Is Nothing Then Exit Sub
  
  mbDisableRaiseEventFieldChange = True
  For Each fld In mRsFieldClone.Fields
    ' EXCETO CHAVE PRIMEIRA E CAMPO AUTOINCREMENT
    If Not mRsADO.Fields(fld.Name).Properties("KEYCOLUMN").Value And Not mRsADO.Fields(fld.Name).Properties("ISAUTOINCREMENT").Value Then
      Debug.Print fld.Name, fld.Value, mRsADO.Fields(fld.Name).Value
      mRsADO.Fields(fld.Name).Value = fld.Value
    End If
  Next
  mbDisableRaiseEventFieldChange = False

  mRsFieldClone.Close
  Set mRsFieldClone = Nothing
  
  Me.Resync
End Sub

Public Function IsRecordDeleted()
On Error Resume Next
  Dim book As Double
  book = mRsADO.Bookmark
  IsRecordDeleted = (Err.Number = -2147217885 Or Err.Number = -2147217864)
  
  Err.Clear
On Error GoTo 0
End Function
