VERSION 5.00
Begin VB.Form frmRepair 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compactar e Corrigir [.MDB]"
   ClientHeight    =   1815
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepair.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAguarde 
      BackColor       =   &H80000014&
      Height          =   645
      Left            =   2760
      ScaleHeight     =   585
      ScaleWidth      =   2235
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label lblAguarde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reparando"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   285
         Left            =   435
         TabIndex        =   11
         Top             =   75
         Width           =   1290
      End
   End
   Begin VB.Frame fraProcessamentoCPU 
      Caption         =   "Processamento CPU (Clonagem)"
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   1185
      Width           =   2970
      Begin VB.OptionButton optProcCPU 
         Caption         =   "Alta"
         Height          =   195
         Index           =   2
         Left            =   2010
         TabIndex        =   9
         Top             =   285
         Width           =   840
      End
      Begin VB.OptionButton optProcCPU 
         Caption         =   "Normal"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   8
         Top             =   285
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton optProcCPU 
         Caption         =   "Baixa"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&Clonar"
      Height          =   465
      Left            =   4635
      TabIndex        =   4
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informe o arquivo [MDB] para ser compactado e corrigido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   5
      TabIndex        =   1
      Top             =   0
      Width           =   6225
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   135
         TabIndex        =   3
         Text            =   "C:\"
         Top             =   360
         Width           =   5565
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   5715
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblMsgProgresso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblLabel1"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   900
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdRepair 
      Caption         =   "&Reparar"
      Height          =   465
      Left            =   3045
      TabIndex        =   0
      Top             =   1275
      Width           =   1530
   End
End
Attribute VB_Name = "frmRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sPastaErro    As String
Private m_sMsgProgress  As String
Private m_bCancelar     As Boolean

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdRepair_Click()
    m_bCancelar = False
    
    If MsgBox("Compactação e Correção do Banco de Dados pode demorar varios minutos dependendo do tamanho do arquivo!" & _
              vbCrLf & "Deseja continuar?", vbQuestion + vbYesNo) = vbYes Then
        picAguarde.Visible = True
        lblAguarde.Caption = "Compactação e Reparação em andamento..."
        
        Call modRepair.CompactMDB(Text1.Text)
        
        picAguarde.Visible = False
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = modRepair.OpenFile(modRepair.STR_FILTER, modRepair.STR_TITLE, Text1.Text)
End Sub

Private Sub cmdClone_Click()
    m_bCancelar = False
    Call ClonarMDB(Text1.Text, True)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        m_bCancelar = True
    End If
End Sub

Private Sub Form_Load()
    
    lblMsgProgresso.Caption = ""
    
    If (Command$ <> "") Then
        Text1.Text = Command$
    Else
        Text1.Text = GetSetting(App.ProductName, "Parametros", "PathDB", App.Path)
    End If
'    lblMsgProgresso.Caption = "Especifique o caminho e nome do aquivo (MDB)" & vbCrLf & _
'                     "a ser reparado e compactado..." & vbCrLf & vbCrLf & _
'                     "Isto pode demorar!"

End Sub

Private Sub Form_Resize()
    picAguarde.Align = vbAlignLeft
    picAguarde.Width = Me.Width - 95
    
    With lblAguarde
        .Left = (Me.Width / 2) - (.Width / 2)
        .Top = ((Me.Height / 2) - (.Height / 2)) - 100
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    SaveSetting App.ProductName, "Parametros", "PathDB", Text1.Text
End Sub

Private Sub ClonarMDB(ByVal pPathDataFile As String, ByVal pCopyData As Boolean)
    Dim db As DAO.Database, db2 As DAO.Database
    Dim tb As DAO.TableDef, tb2 As DAO.TableDef
    Dim qy As DAO.QueryDef, qy2 As DAO.QueryDef
    Dim fd As DAO.Field, fd2 As DAO.Field
    Dim ix As DAO.Index, ix2 As DAO.Index
    Dim fdx As DAO.Field
        
    Dim iProg As Integer
    Dim x As Integer
    Dim sNovaDB As String
    
    Me.MousePointer = vbHourglass
    Me.Refresh
    
    Set db = DAO.OpenDatabase(pPathDataFile)
    
    sNovaDB = Replace(pPathDataFile, Dir(pPathDataFile), "") & Format(Now, "yyyyMMdd.hhmmss") & ".mdb"
    Set db2 = DAO.CreateDatabase(sNovaDB, dbLangGeneral)
    
    cmdRepair.Enabled = False
    Command2.Enabled = False
    cmdClone.Enabled = False
    
    iProg = (db.TableDefs.Count * 2) + db.QueryDefs.Count + 1
    
    m_sMsgProgress = "Clonando estrutura de dados (%)"
    lblMsgProgresso.Caption = Replace(m_sMsgProgress, "%", iProg)
    
    For Each tb In db.TableDefs
        If m_bCancelar Then GoTo GoSair:
        
        If Not (tb.Name Like "MSys*") Then
        
            Set tb2 = db2.CreateTableDef(tb.Name)
            For Each fd In tb.Fields
                If m_bCancelar Then GoTo GoSair:
                
                Set fd2 = tb2.CreateField(fd.Name, fd.Type, fd.Size)
                fd2.Attributes = fd.Attributes
'                Debug.Assert tb.Name & fd.Name <> "VENDAS_CONSIGNACAO_ITENS"
                
                'fd2.AllowZeroLength = fd.AllowZeroLength
                If (fd.Type = 10 Or fd.Type = 12) Then
                    fd2.AllowZeroLength = True
                End If
                fd2.DefaultValue = fd.DefaultValue
                fd2.Required = fd.Required
                fd2.ValidationRule = fd.ValidationRule
                fd2.ValidationText = fd.ValidationText
                
                tb2.Fields.Append fd2
                tb2.Fields.Refresh
            Next
                        
            db2.TableDefs.Append tb2
        End If
        
        iProg = iProg - 1
        lblMsgProgresso.Caption = Replace(m_sMsgProgress, "%", iProg)
        
        DoEvents
    Next
    
    For Each tb In db.TableDefs
        If m_bCancelar Then GoTo GoSair:
        
        If Not (tb.Name Like "MSys*") Then
            
            For Each tb2 In db2.TableDefs
                If m_bCancelar Then GoTo GoSair:
                
                If (tb.Name = tb2.Name) Then
                
                    For Each ix In tb.Indexes
                        If m_bCancelar Then GoTo GoSair:
                        
                        Set ix2 = tb2.CreateIndex(ix.Name)
                        
                        For Each fdx In ix.Fields
                            ix2.Fields.Append ix2.CreateField(fdx.Name)
                        Next
                        
                        ix2.Required = ix.Required
                        ix2.Primary = ix.Primary
                        ix2.Unique = ix.Unique
                        ix2.Clustered = ix.Clustered
                        ix2.IgnoreNulls = ix.IgnoreNulls
                        
                        tb2.Indexes.Append ix2
                        tb2.Indexes.Refresh
                    Next
            
                End If
            Next
            
        End If
        
        iProg = iProg - 1
        lblMsgProgresso.Caption = Replace(m_sMsgProgress, "%", iProg)

        DoEvents
    Next
        
    For Each qy In db.QueryDefs
        If m_bCancelar Then GoTo GoSair:
        
        Set qy2 = db2.CreateQueryDef(qy.Name, qy.SQL)
        
        iProg = iProg - 1
        lblMsgProgresso.Caption = Replace(m_sMsgProgress, "%", iProg)
        DoEvents
    Next
    
GoSair:
    db.Close
    db2.Close
    
    If (pCopyData And m_bCancelar = False) Then
        Call IniciarCopia(pPathDataFile, sNovaDB)
    End If
        
    If m_bCancelar Then
        lblMsgProgresso.Caption = "CLONAGEM CANCELADA!"
    Else
        lblMsgProgresso.Caption = ""
    End If
    
    
    cmdRepair.Enabled = True
    Command2.Enabled = True
    cmdClone.Enabled = True
    
    Me.MousePointer = vbDefault
End Sub

Private Sub IniciarCopia(ByVal pDBOrigem As String, ByVal pDBDestino As String)

    Const PROVIDER As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%PATHFILE%;Persist Security Info=False"
    
    Dim cn As ADODB.Connection, cn2 As ADODB.Connection
    Dim ca As ADOX.Catalog
    Dim tb As ADOX.Table
    
    Me.MousePointer = vbHourglass
    DoEvents
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = Replace(PROVIDER, "%PATHFILE%", pDBOrigem)
    cn.CursorLocation = adUseClient
    cn.Open
    
    Set cn2 = New ADODB.Connection
    cn2.ConnectionString = Replace(PROVIDER, "%PATHFILE%", pDBDestino)
    cn2.CursorLocation = adUseServer
    cn2.Open
    
    Set ca = New ADOX.Catalog
    Set ca.ActiveConnection = cn
    
    m_sPastaErro = Replace(pDBDestino, ".mdb", "")
    
    Dim nTbls As Long
    nTbls = ca.Tables.Count
    
    For Each tb In ca.Tables
        If m_bCancelar Then GoTo GoSair
'        If tb.Name = "VENDAS_CONSIGNACAO_ITENS" Then
        m_sMsgProgress = "Copiando dados na tabela #(" & nTbls & " / %)"

        lblMsgProgresso.Caption = Replace(m_sMsgProgress, "#", tb.Name)
        lblMsgProgresso.Caption = Replace(m_sMsgProgress, "%", 0)
        If tb.Type = "TABLE" Then
            Call CopiarDados(tb.Name, cn, cn2)
        End If
        
        nTbls = nTbls - 1
        DoEvents
'        End If
    Next
    
GoSair:
    Set tb = Nothing
    Set ca = Nothing
    Set cn = Nothing
    Set cn2 = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub CopiarDados(ByVal pNomeTabela As String, ByRef pConnOrigem As ADODB.Connection, ByRef pConnDestino As ADODB.Connection)
On Error GoTo Finaliza

    Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset
    Dim nRCount As Long
    
    Set rs = New ADODB.Recordset
    rs.Open pNomeTabela, pConnOrigem
    
    If (rs.RecordCount > 0) Then
        nRCount = rs.RecordCount
        
'        Set rs2 = New ADODB.Recordset
'        rs2.Open pNomeTabela, pConnDestino, adOpenDynamic, adLockOptimistic
'        Debug.Assert pNomeTabela <> "BANCOS"
        Do While Not rs.EOF
            If m_bCancelar Then GoTo GoSair
            
            lblMsgProgresso.Caption = Replace(m_sMsgProgress, "#", pNomeTabela)
            lblMsgProgresso.Caption = Replace(lblMsgProgresso.Caption, "%", nRCount)
            Call GravarDados(pNomeTabela, pConnDestino, rs.Fields)
            rs.MoveNext
            
            Call SleepCPU
            nRCount = nRCount - 1
            DoEvents
        Loop
'        rs2.Close
    End If
    
GoSair:
    rs.Close
    
Finaliza:
    Set rs = Nothing
'    Set rs2 = Nothing
    
    If Err.Number <> 0 Then
        Call GravarLogTabelaCorrompida(pNomeTabela)
    End If
    
    On Error GoTo 0
End Sub

Private Sub GravarDados(ByVal pNomeTabela As String, ByRef pConnDestino As ADODB.Connection, ByRef pValues As ADODB.Fields)
On Error GoTo Finaliza
    Dim x As Integer
    
    Dim sSQLFields As String
    Dim sSQLValues As String
    
    'pRSDestino.AddNew
    For x = 0 To pValues.Count - 1
        'Debug.Print pValues(x).Name, pValues(x).Type
        
        With pValues(x)
            sSQLFields = sSQLFields & .Name & ", "
            If (IsNull(.Value)) Then
                sSQLValues = sSQLValues & "Null, "
            Else
                Select Case .Type
                    Case 3, 5   'integer, float
                        sSQLValues = sSQLValues & Replace(.Value, ",", ".") & ", "
                    Case 7      'date
                        sSQLValues = sSQLValues & Format(.Value & "", "'yyyy-MM-dd'") & ", "
                    Case 11     'boolean
                        sSQLValues = sSQLValues & CInt(.Value) & ", "
                    Case Else
                        If (.Value = "") Then
                            sSQLValues = sSQLValues & "' ', "
                        Else
                            sSQLValues = sSQLValues & "'" & Replace(.Value, "'", "''") & "', "
                        End If
                End Select
            End If
        End With
        DoEvents
        Me.Refresh
    Next
    
    sSQLFields = Mid(sSQLFields, 1, Len(sSQLFields) - 2)
    sSQLValues = Mid(sSQLValues, 1, Len(sSQLValues) - 2)
    
    Debug.Print pNomeTabela
    Debug.Print sSQLFields
    Debug.Print sSQLValues
        
    pConnDestino.Execute "insert into [" & pNomeTabela & "] (" & sSQLFields & ")values(" & sSQLValues & ")"
    'pRSDestino.Update
    
Finaliza:
    If Err.Number <> 0 Then
'        pRSDestino.CancelUpdate
'        Resume
        Call GravarDadosErro(pNomeTabela, pValues)
    End If
    
    On Error GoTo 0
End Sub


Private Sub GravarDadosErro(ByVal pNomeTabela As String, ByRef pValues As ADODB.Fields)
'     Open Mid(sFileD, 1, Len(sFileD) - 4) & "_INFO.TXT" For Output As #1
'        Print #1, txtResumo.Text
'    Close #1
    
    Dim oFSO        As Scripting.FileSystemObject
    Dim oTexto      As Scripting.TextStream
    Dim sPathFile   As String
    Dim sLog        As String
    Dim fld         As ADODB.Field
    
    Set oFSO = New Scripting.FileSystemObject
            
    If Not oFSO.FolderExists(m_sPastaErro) Then
        oFSO.CreateFolder m_sPastaErro
    End If
    
    sPathFile = m_sPastaErro & "\" & pNomeTabela & ".csv"
    If oFSO.FileExists(sPathFile) Then
        Set oTexto = oFSO.OpenTextFile(sPathFile, ForAppending, True)
    Else
        Set oTexto = oFSO.OpenTextFile(sPathFile, ForWriting, True)
        
        'Monta cabecalho primeira linha
        For Each fld In pValues
            sLog = sLog & fld.Name & ";"
        Next
        sLog = Mid(sLog, 1, Len(sLog) - 1) & vbCrLf
    End If
    
    For Each fld In pValues
        On Error Resume Next
        sLog = sLog & fld.Value & ";"
    Next
    sLog = Mid(sLog, 1, Len(sLog) - 1)
    oTexto.WriteLine sLog
    
    oTexto.Close
    
    Set oTexto = Nothing
    Set oFSO = Nothing
End Sub

Private Sub GravarLogTabelaCorrompida(ByVal pNomeTabela As String)
'    Open Mid(sFileD, 1, Len(sFileD) - 4) & "_INFO.TXT" For Output As #1
'        Print #1, txtResumo.Text
'    Close #1

    Dim oFSO        As Scripting.FileSystemObject
    Dim oTexto      As Scripting.TextStream
    Dim sPathFile   As String
    Dim sLog        As String
    
    Set oFSO = New Scripting.FileSystemObject
            
    If Not oFSO.FolderExists(m_sPastaErro) Then
        oFSO.CreateFolder m_sPastaErro
    End If
    
    sPathFile = m_sPastaErro & "\_TABELAS_TOTALMENTE_CORROMPIDAS.csv"
    If oFSO.FileExists(sPathFile) Then
        Set oTexto = oFSO.OpenTextFile(sPathFile, ForAppending, True)
    Else
        Set oTexto = oFSO.OpenTextFile(sPathFile, ForWriting, True)
        sLog = "TABELAS QUE NÃO FORAM POSSÍVEIS RECUPERAR DADOS" & vbCrLf
        sLog = sLog & "-----------------------------------------------" & vbCrLf
    End If
    
    sLog = sLog & ">> " & pNomeTabela
    oTexto.WriteLine sLog
    
    oTexto.Close
    
    Set oTexto = Nothing
    Set oFSO = Nothing
End Sub


Private Sub SleepCPU()
    Select Case True
    Case optProcCPU(0).Value
        Sleep 30
    Case optProcCPU(1).Value
        Sleep 15
    Case optProcCPU(2).Value
        Sleep 1
    End Select
End Sub
