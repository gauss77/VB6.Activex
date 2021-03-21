VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FOpcoes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caixa Opções"
   ClientHeight    =   1785
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   2550
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LView 
      Height          =   1605
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   2831
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   15859453
      BorderStyle     =   1
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
      NumItems        =   0
   End
End
Attribute VB_Name = "FOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msKey       As String
Private msDescricao As String

Public Sub ShowLOpcoes(ByVal pArrayKeyDesc As String, _
                       ByRef psRetKey As String, ByRef psRetDescricao As String)
                    
  Dim aList() As String
  Dim aKeyDesc() As String
  Dim i As Integer, sK As String
  
  aList = Split(pArrayKeyDesc, "|")
  
  For i = 0 To UBound(aList)
    aKeyDesc = Split(aList(i), ",")
    
    On Error Resume Next
    sK = aKeyDesc(0)
    On Error GoTo 0
    
    If sK <> "" Then
      LView.ListItems.Add , sK, aKeyDesc(1)
    Else
      LView.ListItems.Add , , aKeyDesc(1)
    End If
  Next
  
  Me.Show vbModal
  
  psRetKey = msKey
  psRetDescricao = msDescricao
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set FOpcoes = Nothing
End Sub

Private Sub LView_DblClick()
  msKey = LView.SelectedItem.Key
  msDescricao = LView.SelectedItem.Text
  Me.Hide
End Sub

Private Sub LView_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    LView_DblClick
    KeyCode = 0
  End If
End Sub
