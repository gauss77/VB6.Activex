VERSION 5.00
Object = "{EADE62FD-5B6B-444E-A6C6-26CFE520CF78}#1.0#0"; "ideToolBar.ocx"
Begin VB.Form FPrinterRel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FPrinterRel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFCFC&
      Height          =   1605
      ItemData        =   "FPrinterRel.frx":058A
      Left            =   60
      List            =   "FPrinterRel.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   375
      Width           =   4710
   End
   Begin Insignia_Toolbar.ideToolbar asxToolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   582
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   9
      HotTrackingColor=   255
      ButtonStyle1    =   2
      ButtonStyle2    =   2
      ButtonCaption3  =   "Impressora   "
      ButtonKey3      =   "Imprimir"
      ButtonPicture3  =   "FPrinterRel.frx":058E
      ButtonToolTipText3=   "Manda direto para a Impressora"
      ButtonCaption4  =   "Video   "
      ButtonKey4      =   "View"
      ButtonPicture4  =   "FPrinterRel.frx":08E0
      ButtonToolTipText4=   "Visualizar Impressão"
      ButtonStyle5    =   2
      ButtonKey6      =   "Check"
      ButtonPicture6  =   "FPrinterRel.frx":0C32
      ButtonToolTipText6=   "Marcar todos"
      ButtonKey7      =   "UnCheck"
      ButtonPicture7  =   "FPrinterRel.frx":0F84
      ButtonToolTipText7=   "Desmarcar todos"
      ButtonWidth8    =   120
      ButtonStyle8    =   0
      ButtonCaption9  =   "Fechar   "
      ButtonKey9      =   "Fechar"
      ButtonPicture9  =   "FPrinterRel.frx":12D6
      ButtonToolTipText9=   "Fechar Janela"
      ButtonVisible9  =   0   'False
   End
End
Attribute VB_Name = "FPrinterRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private maPrinter() As String

'Public Event Printer(ByVal NameReport As String, ByVal View As Boolean)
Public Event Printer(ByVal View As Boolean, ByVal ListReports As String)

'Public Sub ShowLRelatorios(ParamArrayargs() As String)
Public Sub ShowLRelatorios(ByVal ListArray As String)
  Dim aTemp() As String
  Dim aList() As String
  Dim i As Integer
  
  aTemp = Split(ListArray, "|")
  
  lstPrint.Clear

  For i = 0 To UBound(aTemp)
    aList = Split(aTemp(i), ",")
    
    lstPrint.AddItem aList(0)
    ReDim Preserve maPrinter(i)
    maPrinter(i) = aList(1)
  Next
  
  If lstPrint.ListCount = 0 Then
    MsgBox "Erro ao carregar lista", vbInformation, App.FileDescription
    Unload Me
  Else
    Me.Show vbModal
  End If
End Sub

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
  Dim i As Integer

  With lstPrint
    Select Case ButtonKey
      Case Is = "Imprimir"
        Call Printer(True)
        
      Case Is = "View"
        Call Printer(False)
        
      Case Is = "Fechar"
         Unload Me
      Case Is = "Check"
        For i = 0 To .ListCount - 1
          .Selected(i) = True
        Next
      Case Is = "UnCheck"
        For i = 0 To .ListCount - 1
          .Selected(i) = False
        Next
    End Select
  End With
End Sub

Private Sub Printer(ByVal bView As Boolean)
  Dim i As Integer
  Dim rpt As String
  
  For i = 0 To lstPrint.ListCount - 1
    If lstPrint.Selected(i) = True Then
      rpt = rpt & "|" & maPrinter(i)
    End If
  Next
  
  If rpt <> "" Then RaiseEvent Printer(bView, Mid(rpt, 2))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set FPrinterRel = Nothing
End Sub
