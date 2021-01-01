VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'cria enumerador público que poderá
'ser referenciado de fora da classe
Public Enum eDSConexao
    cnSucesso = 1       '= Existem registros
    cnVazio = 2         '= Não existem registros
    cnErroProcesso = 3  '= xiiiii! algo saiu errado
End Enum

Public Enum eDSOperacao
   opVisualizacao = 0
   opInclusao = 1
   opAlteracao = 2
End Enum

Public Enum eDSBotoesExtras
    beNone
    bePrinter_Search
    bePrinter_Requery
    beSearch_Requery
    bePrinter
    beSearch
    beRequery
    beAllButtons
End Enum

Public Enum eDSPermissoes
    peTodos
    peIncluir
    peAlterar
    peExcluir
    peIncluir_Excluir
    peIncluir_Alterar
    peAlterar_Excluir
    peNenhuma
End Enum

Public Enum eDSModelo
    mdMaster = 0
    mdSimples = 1
End Enum
