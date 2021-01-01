VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmObjetos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object Monitor"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmObjetos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSempreVisivel 
      Caption         =   "Sempre visível"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   5835
      Width           =   1350
   End
   Begin VB.Timer tmrAtualizar 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3690
      Top             =   5685
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   5595
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjetos.frx":058A
            Key             =   "Processo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjetos.frx":15DC
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjetos.frx":1976
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjetos.frx":1D10
            Key             =   "UserControl"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   360
      Left            =   5520
      TabIndex        =   4
      Top             =   5745
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   6750
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monitora a Criação e Destruição de Objetos (Classes, Forms e UserControls)"
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   6
         Top             =   435
         Width           =   5490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Object Monitor"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   810
         TabIndex        =   5
         Top             =   135
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "frmObjetos.frx":20AA
         Top             =   165
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4905
      Left            =   75
      TabIndex        =   1
      Top             =   750
      Width           =   6750
      Begin MSComctlLib.ListView lstObjetos 
         Height          =   4485
         Left            =   30
         TabIndex        =   3
         Top             =   375
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   7911
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qtd"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Projeto"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Objeto"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Objetos existentes na memória."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2055
         TabIndex        =   2
         Top             =   165
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000001&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H80000001&
         Height          =   255
         Left            =   45
         Top             =   135
         Width           =   6645
      End
   End
End
Attribute VB_Name = "frmObjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Atualizar(pArrayObjetos As Variant)
   Dim i As Integer
   Dim s As String
   Dim lsItem As ListItem
   Dim nBoundArray As Integer
   Dim nPosPonto As Integer
   Dim sProjeto As String
   Dim sObjeto As String
   Dim nObjCount As Integer
   Dim sTipo As String
   
   On Error Resume Next
   nBoundArray = UBound(pArrayObjetos, 2)
   If Err.Number <> 0 Then Exit Sub
   On Error GoTo 0
   
   lstObjetos.ListItems.Clear
   For i = 0 To nBoundArray
      nObjCount = pArrayObjetos(2, i)
      
      If nObjCount > 0 Then
         s = pArrayObjetos(1, i)
         sProjeto = Mid(s, 1, (InStr(1, s, ".")) - 1)
         sObjeto = Mid(s, (InStr(1, s, ".")) + 1)
      
         Select Case UCase$(Mid(sObjeto, 1, 3))
            Case "FRM"
               sTipo = "Form"
               Set lsItem = lstObjetos.ListItems.Add(, , nObjCount, "Form", "Form")
            Case "CTR"
               sTipo = "UserControl"
               Set lsItem = lstObjetos.ListItems.Add(, , nObjCount, "UserControl", "UserControl")
            Case Else
               sTipo = "Class Module"
               Set lsItem = lstObjetos.ListItems.Add(, , nObjCount, "Class", "Class")
         End Select
      
         
         lsItem.SubItems(1) = sProjeto
         lsItem.SubItems(2) = sObjeto
         lsItem.SubItems(3) = sTipo
      End If
   Next
End Sub

Private Sub chkSempreVisivel_Click()
  If chkSempreVisivel.Value Then
    Modulo.AlwaysOnTop Me, True
  Else
    Modulo.AlwaysOnTop Me, False
  End If
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub lstObjetos_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   lstObjetos.SortKey = ColumnHeader.Index - 1
   lstObjetos.Sorted = True
End Sub

Private Sub Timer1_Timer()
'   If Alterou Then
'      Call Atualizar(ArrayObjetos)
'   End If
End Sub
