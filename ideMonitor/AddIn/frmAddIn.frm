VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmAddIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monitor"
   ClientHeight    =   6390
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5445
      Left            =   3660
      TabIndex        =   11
      Top             =   0
      Width           =   8250
      Begin VB.TextBox txtLOG 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Text            =   "frmAddIn.frx":0000
         Top             =   450
         Width           =   8130
      End
      Begin VB.Label lblLabel1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Arquivo de LOG"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   45
         TabIndex        =   12
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Adicionar Código"
      Enabled         =   0   'False
      Height          =   690
      Left            =   60
      TabIndex        =   6
      Top             =   5490
      Width           =   1740
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Remover Código"
      Enabled         =   0   'False
      Height          =   690
      Left            =   1845
      TabIndex        =   5
      Top             =   5490
      Width           =   1740
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3060
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0011
            Key             =   "Key0"
            Object.Tag             =   "vbext_ct_Project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":016B
            Key             =   "Key1"
            Object.Tag             =   "vbext_ct_StdModule"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":02C5
            Key             =   "Key2"
            Object.Tag             =   "vbext_ct_ClassModule"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":041F
            Key             =   "Key5"
            Object.Tag             =   "vbext_ct_VBForm"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0579
            Key             =   "Key5c"
            Object.Tag             =   "vbext_ct_VBMDIChild"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":06D3
            Key             =   "Key6"
            Object.Tag             =   "vbext_ct_VBMDIForm"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":082D
            Key             =   "Key7"
            Object.Tag             =   "vbext_ct_PropPage"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0987
            Key             =   "Key8"
            Object.Tag             =   "vbext_ct_UserControl"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0AE1
            Key             =   "Key11"
            Object.Tag             =   "vbext_ct_Designers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0C3B
            Key             =   "KeyF"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExcluirLog 
      Caption         =   "Excluir &Log"
      Height          =   690
      Left            =   10125
      TabIndex        =   4
      Top             =   5505
      Width           =   1740
   End
   Begin MSComctlLib.ProgressBar barProgresso 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   6255
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3660
      Begin VB.OptionButton optEscopo 
         Caption         =   "Módulo &Atual"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   1065
         Width           =   1275
      End
      Begin VB.OptionButton optEscopo 
         Caption         =   "Em &todos os Módulos"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   1830
      End
      Begin VB.CheckBox chkObjeto 
         Caption         =   "Criação e Destruição de &Objetos"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   405
         Width           =   2625
      End
      Begin VB.CheckBox chkProcesso 
         Caption         =   "Etapas dos &Processos"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   165
         Width           =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   150
         X2              =   3405
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   150
         X2              =   3405
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4140
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   1305
      Width           =   3660
      Begin MSComctlLib.TreeView trvObj 
         Height          =   3915
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   6906
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImgList"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Modo
   [Todos] = 1
   [Atual] = 2
End Enum

Private Enum eExecução
   [Inclui] = 1
   [Elimina] = 2
End Enum

Private Enum eMonitor
   [Processo] = 1
   [Objeto] = 2
End Enum

Public VBInstance As VBIDE.VBE
Public Connect As Connect
Private objCodeModule As CodeModule

Private Const CALL_NAME As String = "Call CallStackMonitor.Gravar.Log"

Private Sub cmdExcluirLog_Click()
    Dim objFso As Scripting.FileSystemObject
   
    Set objFso = New Scripting.FileSystemObject
    
    On Error Resume Next
    objFso.DeleteFile App.Path & "\Monitor.log", True
    On Error GoTo 0
    
    txtLOG.Text = "Arquivo vazio..."
End Sub

'Carrega o trvObj com todos os Projetos e seus respectivos componentes
Private Sub Form_Load()
    On Error GoTo TrataErro
    
    Dim vbP As VBProject
    Dim vbC As VBComponent
    Dim vbR As Reference
    
    Dim sProjeto As String
    Dim bRefOK As Boolean
    Dim sKeySRoot As String
    Dim sKeyIcon As String
    
    Me.Top = 60
    
    Call PreencherTextBox
        
    Set VBInstance = Connect.VBInstance
        
    With VBInstance
        For Each vbP In .VBProjects
            bRefOK = False
            
'            'Verifica se projet está usando a referencia do monitor
'            For Each vbR In vbP.References
'                If (vbR.Name = App.EXEName) Then
'                    bRefOK = True
'                    Exit For
'                End If
''                If (vbR.Name = "AddInMonitor" And vbR.Guid = "{E91DC3C5-F386-4D2A-9FF3-E4B6DB0F481B}") Then
''                    bRefOK = True
''                    Exit For
''                End If
'            Next
'
'            'Adicionando a referencia se não existir
'             If (bRefOK = False) Then
'                vbP.References.AddFromFile App.LogPath
'                'vbP.References.AddFromGuid "{E91DC3C5-F386-4D2A-9FF3-E4B6DB0F481B}", App.Major, App.Minor
''                .VBProjects(p).References.AddFromGuid("asd", 0, 0)
                'Call AddMyRef(vbP)
'            End If
            
            
            sProjeto = vbP.Name
            trvObj.Nodes.Add , , sProjeto, sProjeto, "Key0"
'            trvObj.Nodes(sProjeto).Image = 1
             
            For Each vbC In vbP.VBComponents
                sKeyIcon = "Key" & vbC.Type
                
                Select Case vbC.Type
                    Case 1 'Modules
                        sKeySRoot = AddSubRoot(sProjeto, "Modules")
                    Case 2 'Class Modules
                        sKeySRoot = AddSubRoot(sProjeto, "Class Modules")
                    Case 5 'Forms
                        sKeySRoot = AddSubRoot(sProjeto, "Forms")
                        If (CBool(vbC.Properties("MDIChild").Value) = True) Then
                            sKeyIcon = sKeyIcon & "c"
                        End If
                    Case 6 'MDIForm
                        sKeySRoot = AddSubRoot(sProjeto, "MDIForm")
                    Case 7 'Property Pages
                        sKeySRoot = AddSubRoot(sProjeto, "Property Pages")
                    Case 8 'User Controls
                        sKeySRoot = AddSubRoot(sProjeto, "User Controls")
                    Case 11 'Designers
                        sKeySRoot = AddSubRoot(sProjeto, "Designers")
                    Case Else
                        sKeyIcon = ""
                End Select
                
                If sKeyIcon <> "" Then
                    trvObj.Nodes.Add sKeySRoot, tvwChild, sKeySRoot & "_" & vbC.Name, vbC.Name, sKeyIcon
                End If
            Next
            
            trvObj.Nodes(sProjeto).Sorted = True
        Next
        
    End With
    
    Exit Sub

TrataErro:
   MsgBox "frmAddIn.Form_Load" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & _
          "Fonte: " & Err.Source

End Sub

Public Sub AddMyRef(ByRef pProj As VBProject)
    On Error GoTo TrataErro
    
    pProj.References.AddFromFile App.Path & "\" & App.EXEName
   
TrataErro:

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    
    Set frmAddIn = Nothing
End Sub

Private Sub chkObjeto_Click()
    Dim bEnabled As Boolean
    
    bEnabled = chkProcesso.Value = vbChecked Or chkObjeto.Value = vbChecked
    
    optEscopo([Todos]).Enabled = bEnabled
    optEscopo([Atual]).Enabled = bEnabled
    cmdIncluir.Enabled = bEnabled
    cmdEliminar.Enabled = bEnabled
End Sub

Private Sub chkProcesso_Click()
   Call chkObjeto_Click
End Sub

Private Sub cmdEliminar_Click()
   Call Executar([Elimina], False, False)
End Sub

Private Sub cmdIncluir_Click()
   Call Executar([Inclui], chkProcesso.Value, chkObjeto.Value)
End Sub

Private Sub optEscopo_Click(Index As Integer)
   Dim nodComponentes As Node
   Dim objCodePane As CodePane
   Dim sAtual As String, sNode As String
   
   Select Case Index
      Case [Todos]
         'Marca todos os Nodes do trvObj
         For Each nodComponentes In trvObj.Nodes
            nodComponentes.Checked = True
         Next
      Case [Atual]
         'Desmarca todos os Nodes do trvObj
         For Each nodComponentes In trvObj.Nodes
            nodComponentes.Checked = False
         Next
         
         'Marca o Node associado ao CodePane Ativo
         On Error Resume Next
         sNode = VBInstance.ActiveCodePane.CodeModule.Parent.Name
         Set nodComponentes = trvObj.Nodes(sNode)
         On Error GoTo 0
            
         If Not nodComponentes Is Nothing Then
            nodComponentes.Checked = True
            nodComponentes.Parent.Expanded = True
         Else
            MsgBox "Nenhuma Janela de Código está Ativa", vbCritical
         End If
   End Select

End Sub

'Permite que quando se clique num Node-Pai, todos os Node-Filhos também são marcados
Private Sub trvObj_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo TrataErro
    Dim nd As Node
    Dim sTemp As String

    If (Node.Children <> 0) Then
        For Each nd In trvObj.Nodes
            If (Not nd.Parent Is Nothing) Then
                If (Not nd.Parent.Parent Is Nothing) Then
                    If (nd.Parent.Parent.Key = Node.Key) Then
                        nd.Checked = Node.Checked
                    End If
                End If
                If (nd.Parent.Key = Node.Key) Then
                    nd.Checked = Node.Checked
                End If
            End If
        Next
    End If

Exit Sub

TrataErro:
   MsgBox "frmAddIn.trvObj_NodeCheck" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & _
          "Fonte: " & Err.Source

End Sub

Private Function AddSubRoot(sRoot As String, sSRoot As String)
    On Error GoTo AddNode
    Dim sKey As String
    
    sKey = sRoot & "_" & sSRoot
    
    AddSubRoot = trvObj.Nodes(sKey).Key
    Exit Function
    
AddNode:
    trvObj.Nodes.Add sRoot, tvwChild, sKey, sSRoot, "KeyF"
    AddSubRoot = sKey
End Function

'Percorre todo o trvObj em busca de Nodes selecionados
Private Sub Executar(nExecução As eExecução, bProcesso As Boolean, bObjeto As Boolean)
    On Error Resume Next
    Dim ndComp As Node
    Dim sProjeto As String
    
    barProgresso.Enabled = True
   
    barProgresso.Max = trvObj.Nodes.Count
    
    For Each ndComp In trvObj.Nodes
        If ndComp.Children = False And ndComp.Checked Then
            sProjeto = ndComp.Parent.Parent.Text
            
            Set objCodeModule = VBInstance.VBProjects(sProjeto).VBComponents(ndComp.Text).CodeModule
         
            If nExecução = [Inclui] Then
                Call Incluir(sProjeto, bProcesso, bObjeto)
            Else
                Call Remover
            End If
        End If
        
        barProgresso.Value = barProgresso.Value + 1
    Next
    
    On Error GoTo 0
   
   barProgresso.Value = 0
   
   Set objCodeModule = Nothing
   Set ndComp = Nothing
End Sub

Private Sub Incluir(ByVal pProjectName As String, ByVal pProcesso As Boolean, ByVal pObjeto As Boolean)
   
   If pObjeto Then
      Call MonitorarObjeto(pProjectName)
   End If

   If pProcesso Then
      Call MonitorarProcesso(pProjectName)
   End If
   
   objCodeModule.CodePane.Window.Close
End Sub

Private Sub MonitorarProcesso(ByVal pProjectName As String)
   Dim objMembro As Member
   Dim nLinha As Integer, i As Integer
   Dim sMembro As String, sNomeLog As String
   
   On Error GoTo TrataErro
      
   For Each objMembro In objCodeModule.Members
      
      sMembro = objMembro.Name
      
      If objMembro.Type = vbext_mt_Method Then
         'Os eventos UserControl ou Class Initialize e Terminate são
         'encarados pelo VB como um método
         If InStr(1, sMembro, "_Initialize") = 0 And InStr(1, sMembro, "_Terminate") = 0 Then
            nLinha = objCodeModule.ProcBodyLine(sMembro, vbext_pk_Proc)
            Call GravaLog(pProjectName, sMembro, [Processo], nLinha)
         End If
         
      ElseIf objMembro.Type = vbext_mt_Property Then
         For i = 1 To 3
            
            On Error Resume Next
            nLinha = objCodeModule.ProcBodyLine(sMembro, i)
            On Error GoTo TrataErro
            
            If nLinha <> 0 Then
               
               Select Case i
                  Case 1
                     sNomeLog = "Property Let " & sMembro
                  Case 2
                     sNomeLog = "Property Set " & sMembro
                  Case 3
                     sNomeLog = "Property Get " & sMembro
               End Select
               
               Call GravaLog(pProjectName, sNomeLog, [Processo], nLinha)
               Exit For
            End If
         Next
      End If
   Next
   
   Exit Sub
   
TrataErro:
   MsgBox "frmAddIn.MonitorarProcesso" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & _
          "Fonte: " & Err.Source
          
End Sub

Private Sub MonitorarObjeto(ByVal pProjectName As String)
    Dim sTipo As String
    
    Select Case objCodeModule.Parent.Type
    Case 2:     sTipo = "Class"
    Case 5:     sTipo = "Form"
    Case 6:     sTipo = "MDIForm"
    Case 8:     sTipo = "UserControl"
    Case 11
        If (objCodeModule.Parent.DesignerID = "MSDataReportLib.DataReport") Then
            sTipo = "DataReport"
        Else
            sTipo = "DataEnvironment"
        End If
    End Select
    
    If sTipo <> "" Then
        Call GravaLog(pProjectName, sTipo & "_Initialize", Objeto)
        Call GravaLog(pProjectName, sTipo & "_Terminate", Objeto)
    End If
End Sub

Private Sub GravaLog(ByVal pProjectName As String, ByVal sMembro As String, Monitorar As eMonitor, Optional nLinha As Integer = 0)
    On Error GoTo TrataErro
    Dim sCall As String, sProc As String, sPosProc As String
    Dim sEvento As String, sTipo As String
    Dim nLinhaProc As Integer
    
    sCall = "    " & CALL_NAME & "(""{0}"", ""{1}.{2}.{3}"")"
       
    sCall = Replace(sCall, "{1}", pProjectName)
    sCall = Replace(sCall, "{2}", objCodeModule.Parent.Name)
    sCall = Replace(sCall, "{3}", sMembro)
    
    If Monitorar = Processo Then
        sCall = Replace(sCall, "{0}", "Processo")
                
        'Retorna uma String contendo a Procedure
        nLinhaProc = nLinha
        sProc = objCodeModule.Lines(nLinhaProc, 1)
      
        'O Loop é necessário para evitar inserção de código errado
        'em procedures com mais de uma linha
        While Right(sProc, 1) = "_"
            nLinhaProc = nLinhaProc + 1
            sProc = objCodeModule.Lines(nLinhaProc, 1)
        Wend
      
        'Retorna a linha REAL após a procedure
        sPosProc = objCodeModule.Lines(nLinhaProc + 1, 1)
      
        If InStr(1, sPosProc, CALL_NAME) = 0 Then
            If nLinhaProc > objCodeModule.CountOfDeclarationLines Then
                objCodeModule.InsertLines nLinhaProc + 1, sCall
            End If
        End If
    
    Else
        sCall = Replace(sCall, "{0}", "Objeto")
        
        On Error Resume Next
        nLinhaProc = objCodeModule.ProcBodyLine(sMembro, vbext_pk_Proc)
        On Error GoTo TrataErro
      
        If nLinhaProc = 0 Then
            sEvento = Mid(sMembro, InStrRev(sMembro, "_") + 1)
            sTipo = Mid(sMembro, 1, InStr(1, sMembro, "_") - 1)
         
            nLinhaProc = objCodeModule.CreateEventProc(sEvento, sTipo)
        End If
      
        sPosProc = objCodeModule.Lines(nLinhaProc + 1, 1)
      
        If InStr(1, sPosProc, CALL_NAME) = 0 Then
            objCodeModule.InsertLines nLinhaProc + 1, sCall
        End If
    End If
       
    Exit Sub
   
TrataErro:
   MsgBox "frmAddIn.GravaLog" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & _
          "Fonte: " & Err.Source, , objCodeModule.Parent.Name

End Sub

'Remova todas as chamadas à procedure Gravar
Private Sub Remover()
   Dim sLinha As String
   Dim i As Integer
   
   On Error GoTo TrataErro
   
   For i = 1 To objCodeModule.CountOfLines
      sLinha = objCodeModule.Lines(i, 1)
      
      If InStr(1, sLinha, CALL_NAME, vbTextCompare) <> 0 Then
         objCodeModule.DeleteLines i
      End If
   
   Next
   
   Exit Sub

TrataErro:
   MsgBox "frmAddIn.Remover" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & _
          "Fonte: " & Err.Source
End Sub

Private Sub PreencherTextBox()
   Dim objFso As Scripting.FileSystemObject
   Dim objTexto As TextStream
   Dim sTextoLog As String
   
   Set objFso = New Scripting.FileSystemObject
   
   On Error GoTo TrataErro
   
   Set objTexto = objFso.OpenTextFile(App.Path & "\Monitor.Log", ForReading)
   
   txtLOG.Text = ""
   sTextoLog = objTexto.ReadAll
   
   txtLOG.Text = sTextoLog
   
   objTexto.Close
   
   Set objTexto = Nothing
   Set objTexto = Nothing
      
   Exit Sub

TrataErro:
   txtLOG.Text = "Arquivo Vazio..."
   Exit Sub
End Sub
