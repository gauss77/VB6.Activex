VERSION 5.00
Begin VB.UserControl ideFrame 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   PropertyPages   =   "ideFrame.ctx":0000
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   131
   ToolboxBitmap   =   "ideFrame.ctx":003D
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ideFrame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   0
      Top             =   285
      Width           =   645
   End
End
Attribute VB_Name = "ideFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private mbReadProperty As Boolean

'Declaração da API
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long

'Declaração da Type Rect
Private Type Rect
    Left    As Long '<- Lado Esquerdo
    Top     As Long '<- Cima
    Right   As Long '<- Lado Direito
    Bottom  As Long '<- Baixo
End Type
Private mRectExter As Rect
Private mRectInter As Rect

'Declaração das Constantes
'Constantes de estilo
Const BDR_INNER As Long = &HC
Const BDR_OUTER As Long = &H3
Const BDR_RAISED As Long = &H5
Const BDR_RAISEDINNER As Long = &H4
Const BDR_RAISEDOUTER As Long = &H1
Const BDR_SUNKEN As Long = &HA
Const BDR_SUNKENOUTER As Long = &H8
Const BDR_SUNKENINNER As Long = &H2
Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
Const EDGE_ETCHED As Long = (BDR_SUNKENINNER Or BDR_RAISEDINNER)

'Constantes de bordas
Const BF_BOTTOM As Long = &H8
Const BF_LEFT As Long = &H1
Const BF_RIGHT As Long = &H4
Const BF_TOP As Long = &H2
Const BF_TOPLEFT As Long = (BF_TOP Or BF_LEFT)
Const BF_TOPRIGHT As Long = (BF_TOP Or BF_RIGHT)
Const BF_TOPBOTTOM As Long = (BF_TOP Or BF_BOTTOM)
Const BF_LEFTRIGHT As Long = (BF_LEFT Or BF_RIGHT)
Const BF_LEFTBOTTOM As Long = (BF_LEFT Or BF_BOTTOM)
Const BF_BOTTOMRIGHT As Long = (BF_BOTTOM Or BF_RIGHT)
Const BF_RECT As Long = (BF_TOP Or BF_RIGHT Or BF_BOTTOM Or BF_LEFT)
'FIM DECLARACOES API

Public Enum eFRABorderStyle
    bs_None = 0
    bs_Raised = BDR_RAISED
    bs_RaisedInner = BDR_RAISEDINNER
    bs_Inset = BDR_SUNKEN
    bs_InsetInner = BDR_SUNKENINNER
    bs_Frame = EDGE_ETCHED
    bs_Bump = EDGE_BUMP
End Enum
Private meBorderExt As eFRABorderStyle  'Bordas Externas
Private meBorderInt As eFRABorderStyle  'Bordas Internas
Private meBorderTemp As eFRABorderStyle  'Bordas Internas para retorno de simulaçao de botao

Public Enum eFRABorderPaint
    bp_Rect = BF_RECT
    bp_Top = BF_TOP
    bp_Left = BF_LEFT
    bp_Right = BF_RIGHT
    bp_Bottom = BF_BOTTOM
    bp_TopLeft = BF_TOPLEFT
    bp_TopRight = BF_TOPRIGHT
    bp_TopBottom = BF_TOPBOTTOM
    bp_LeftRight = BF_LEFTRIGHT
    bp_LeftBottom = BF_LEFTBOTTOM
    bp_BottomRight = BF_BOTTOMRIGHT
End Enum
Private meBorderPaint As eFRABorderPaint

Public Enum eFRAGradientStyle
    gsNone
    gsVertical
    gsHorizontal
    gsBox
    gsButtonRight
End Enum
Private meGradStyle As eFRAGradientStyle

Enum eFRACaptionAlignPos
    casBorderExt
    casBorderInt
End Enum
Private meCaptionAlignPos As eFRACaptionAlignPos

Enum eFRACaptionAlign
    caCenter = 0
    caCenterTop = 1
    caCenterBottom = 2
    caLeftCenter = 3
    caLeftTop = 4
    caLeftBottom = 5
    caRightCenter = 6
    caRightTop = 7
    caRightBottom = 8
End Enum
Private meCapAling    As eFRACaptionAlign

Enum eFRACaptionBackStyle
    cbsTransparent = 0
    cbsOpaque = 1
End Enum

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Default Property Values:
Private Const m_def_BorderWidth As Integer = 3
Private Const m_def_BackColor   As Long = vbButtonFace
Private Const m_def_BackColorB  As Long = vbButtonShadow
Private Const m_def_ForeColor   As Long = vbWindowText
Private Const m_def_CaptionBackColor As Long = vbButtonFace
Private Const m_def_CaptionBackStyle As Long = vbTransparent

'Property Variables:
Private moBackColorB        As OLE_COLOR

Private mnBorderWidth       As Integer
Private msRGB               As String
Private msToolTip           As String

Public Sub About()
Attribute About.VB_UserMemId = -552
    FormSplash.Show vbModal
End Sub

Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
  hwnd = UserControl.hwnd
End Property

Private Sub UserControl_InitProperties()
    meBorderExt = bs_RaisedInner
    meBorderInt = bs_None
    meBorderPaint = bp_Rect
    mnBorderWidth = m_def_BorderWidth
    
    meGradStyle = gsNone
    moBackColorB = m_def_BackColorB
    
    Call Draw
    
    Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    msToolTip = PropBag.ReadProperty("ToolTipText", "")
    meBorderExt = PropBag.ReadProperty("BorderExt", bs_RaisedInner)
    meBorderInt = PropBag.ReadProperty("BorderInt", bs_None)
    meBorderPaint = PropBag.ReadProperty("BorderPaint", bp_Rect)
    mnBorderWidth = PropBag.ReadProperty("BorderWidth", m_def_BorderWidth)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    moBackColorB = PropBag.ReadProperty("BackColorB", m_def_BackColorB)
    
    GradientStyle = PropBag.ReadProperty("GradientStyle", 0)
    
    lblCaption.Caption = PropBag.ReadProperty("Caption", "")
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    lblCaption.BackColor = PropBag.ReadProperty("CaptionBackColor", m_def_CaptionBackColor)
    lblCaption.BackStyle = PropBag.ReadProperty("CaptionBackStyle", m_def_CaptionBackStyle)
    meCapAling = PropBag.ReadProperty("CaptionAlign", 0)
    meCaptionAlignPos = PropBag.ReadProperty("CaptionAlignPos", 0)
    
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ToolTipText", msToolTip, "")
    Call PropBag.WriteProperty("BorderExt", meBorderExt, bs_RaisedInner)
    Call PropBag.WriteProperty("BorderInt", meBorderInt, bs_None)
    Call PropBag.WriteProperty("BorderPaint", meBorderPaint, bp_Rect)
    Call PropBag.WriteProperty("BorderWidth", mnBorderWidth, m_def_BorderWidth)
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackColorB", moBackColorB, m_def_BackColorB)
    
    Call PropBag.WriteProperty("GradientStyle", GradientStyle, 0)
    
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("CaptionBackColor", lblCaption.BackColor, m_def_CaptionBackColor)
    Call PropBag.WriteProperty("CaptionBackStyle", lblCaption.BackStyle, m_def_CaptionBackStyle)
    Call PropBag.WriteProperty("CaptionAlign", meCapAling, 0)
    Call PropBag.WriteProperty("CaptionAlignPos", meCaptionAlignPos, 0)
    
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    
    Call PropBag.WriteProperty("Enabled", Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Painel3D
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
   Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   lblCaption.Caption = New_Caption
   PropertyChanged "Caption"
   Call AlinharCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,2
Public Property Get CaptionAlign() As eFRACaptionAlign
   CaptionAlign = meCapAling
End Property

Public Property Let CaptionAlign(ByVal New_AlinharCaption As eFRACaptionAlign)
   meCapAling = New_AlinharCaption
   PropertyChanged "CaptionAlign"
   Call AlinharCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,2
Public Property Get BorderExt() As eFRABorderStyle
   BorderExt = meBorderExt
End Property

Public Property Let BorderExt(ByVal New_BordasExternas As eFRABorderStyle)
    meBorderExt = New_BordasExternas
    PropertyChanged "BorderExt"
     
    Call Draw
End Property

Public Property Get BorderInt() As eFRABorderStyle
    BorderInt = meBorderInt
End Property

Public Property Let BorderInt(ByVal New_BordasInternas As eFRABorderStyle)
    meBorderInt = New_BordasInternas
    PropertyChanged "BorderInt"
    
    UserControl_Resize
End Property

Public Property Get BorderPaint() As eFRABorderPaint
  BorderPaint = meBorderPaint
End Property

Public Property Let BorderPaint(ByVal vNewValue As eFRABorderPaint)
    meBorderPaint = vNewValue
    PropertyChanged "BorderPaint"
    
    Call Draw
End Property

Public Property Get BorderWidth() As Integer
    BorderWidth = mnBorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    mnBorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Call Draw
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Font"
Attribute ForeColor.VB_UserMemId = -513
   ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   lblCaption.ForeColor = New_ForeColor
   PropertyChanged "ForeColor"
End Property

Public Property Get ToolTipText() As String
    ToolTipText = msToolTip
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    msToolTip = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Enabled.VB_UserMemId = -514
    Enabled = lblCaption.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If Ambient.UserMode Then
        UserControl.Enabled = New_Enabled
    End If
    lblCaption.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get CaptionBackColor() As OLE_COLOR
   CaptionBackColor = lblCaption.BackColor
End Property

Public Property Let CaptionBackColor(ByVal New_Value As OLE_COLOR)
   lblCaption.BackColor = New_Value
   PropertyChanged "CaptionBackColor"
End Property

Public Property Get CaptionBackStyle() As eFRACaptionBackStyle
   CaptionBackStyle = lblCaption.BackStyle
End Property

Public Property Let CaptionBackStyle(ByVal New_Value As eFRACaptionBackStyle)
   lblCaption.BackStyle = New_Value
   PropertyChanged "CaptionBackStyle"
End Property

Public Property Get CaptionAlignPos() As eFRACaptionAlignPos
  CaptionAlignPos = meCaptionAlignPos
End Property

Public Property Let CaptionAlignPos(vNewValue As eFRACaptionAlignPos)
  meCaptionAlignPos = vNewValue
  
  PropertyChanged "CaptionAlignPos"
  Call AlinharCaption
End Property

Private Sub lblCaption_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_Resize()
    Call Draw
    Call AlinharCaption 'Ajusta o Label
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If mbClickEffect Then
'        If BorderInt <> bs_None Then
'            meBorderTemp = meBorderInt
'            BorderInt = bs_InsetInner
'        Else
'            meBorderTemp = meBorderExt
'            BorderExt = bs_InsetInner
'        End If
'    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If mbClickEffect Then
'        If BorderInt <> bs_None Then
'            BorderInt = meBorderTemp
'        Else
'            BorderExt = meBorderTemp
'        End If
'    End If
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   UserControl_MouseUp Button, Shift, X, Y
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  Set UserControl.Picture = New_Picture
  PropertyChanged "Picture"
End Property

Public Property Get GradientStyle() As eFRAGradientStyle
   GradientStyle = meGradStyle
End Property

Public Property Let GradientStyle(ByVal New_GradientStyle As eFRAGradientStyle)
   meGradStyle = New_GradientStyle
   PropertyChanged "GradientStyle"
   
   Call Draw
End Property

Public Property Get BackColorB() As OLE_COLOR
   BackColorB = moBackColorB
End Property

Public Property Let BackColorB(ByVal New_Value As OLE_COLOR)
   moBackColorB = New_Value
   PropertyChanged "BackColorB"
   
   Call Draw
End Property

Private Sub AlinharCaption()
  Dim nTop As Integer
  
  If meCaptionAlignPos = casBorderInt Then
    nTop = 2
  Else
    nTop = mnBorderWidth + 2
  End If
   
  With lblCaption
    Select Case meCapAling
      Case Is = caCenterBottom
        .Left = (ScaleWidth - .Width) / 2
        .Top = (ScaleHeight - .Height) - nTop

      Case Is = caCenter
        .Left = (ScaleWidth - .Width) / 2
        .Top = (ScaleHeight - .Height) / 2
         
      Case Is = caCenterTop
        .Left = (ScaleWidth - .Width) / 2
        .Top = nTop
      
      Case Is = caLeftBottom, caLeftCenter, caLeftTop
        .Left = nTop
        
        If meCapAling = caLeftCenter Then
          .Top = (ScaleHeight - .Height) / 2
        ElseIf meCapAling = caLeftBottom Then
          .Top = (ScaleHeight - .Height) - nTop
        Else
          .Top = nTop
        End If
      
      Case Is = caRightBottom, caRightCenter, caRightTop
        .Left = (ScaleWidth - .Width) - nTop
      
        If meCapAling = caRightCenter Then
          .Top = (ScaleHeight - .Height) / 2
        ElseIf meCapAling = caRightBottom Then
          .Top = ScaleHeight - .Height - nTop
        Else
          .Top = nTop
        End If
    End Select
  End With
End Sub

Private Sub Draw()
    Cls
    If meGradStyle <> gsNone Then Call DrawGradiente
    
    With mRectExter
        .Bottom = ScaleHeight
        .Right = ScaleWidth
    End With
    Call DrawEdge(hDC, mRectExter, meBorderExt, meBorderPaint)
    
    If mnBorderWidth > 0 Then
        With mRectInter
            .Top = mnBorderWidth
            .Left = mnBorderWidth
            .Bottom = ScaleHeight - mnBorderWidth
            .Right = ScaleWidth - mnBorderWidth
        End With
        Call DrawEdge(hDC, mRectInter, meBorderInt, meBorderPaint)
    End If
End Sub

Private Sub DrawGradiente()
    Dim X, X2 As Integer
    
    UserControl.Line (0, 0)-(ScaleWidth, ScaleHeight), QBColor(1), BF
    
    For X = 0 To 128
        Select Case meGradStyle
            Case Is = gsVertical
                UserControl.Line (0, X * ScaleHeight / 128)- _
                (ScaleWidth, ScaleHeight), ColorBlend(UserControl.BackColor, moBackColorB, (X + 1) / 130), BF
            Case Is = gsHorizontal
                UserControl.Line (X * ScaleWidth / 128, 0)- _
                (ScaleWidth, ScaleHeight), ColorBlend(UserControl.BackColor, moBackColorB, (X + 1) / 130), BF
            
            Case Is = gsBox
                UserControl.Line (X * ScaleWidth / 256, X * ScaleHeight / 256)- _
                (ScaleWidth - (X * ScaleWidth / 256), _
                ScaleHeight - (X * ScaleHeight / 256)), ColorBlend(UserControl.BackColor, moBackColorB, (X + 1) / 130), BF
            
            Case Is = gsButtonRight
                UserControl.Line (X * ScaleWidth / 128, X * ScaleHeight / 128)- _
                (ScaleWidth, ScaleHeight), ColorBlend(UserControl.BackColor, moBackColorB, (X + 1) / 130), BF
        End Select
    Next
End Sub

Public Sub RGBPaint(ByVal pColorR As Long, ByVal pColorG As Long, ByVal pColorB As Long, Optional pbBorder As Boolean)
               
  Dim nLC, nLF As Long, nLang As Long
  Dim nR, nG, nB As Long
  Dim nW As Integer, nH As Integer
  
  msRGB = ""
  nLang = mnBorderWidth
  nLF = nLang / 2
  
  If nLF = 0 Then Exit Sub
  
  msRGB = pColorR & "," & pColorG & "," & pColorB
  UserControl.Cls
  
  nW = ScaleWidth
  nH = ScaleHeight
  
  
  nR = pColorR / nLF
  nG = pColorG / nLF
  nB = pColorB / nLF
  
  While nLang >= nLF
    If Not pbBorder Then
      UserControl.Line (0 + nLC, 0 + nLC)- _
                       (nW - nLC, nH - nLC), _
                        RGB(pColorR, pColorG, pColorB), BF
    Else
      UserControl.Line (0 + nLC, 0 + nLC)- _
                       (nW - nLC, nH - nLC), _
                        RGB(pColorR, pColorG, pColorB), B
    
      UserControl.Line (0 + nLang, 0 + nLang)- _
                       (nW - nLang, nH - nLang), _
                        RGB(pColorR, pColorG, pColorB), B
    
    End If
    
    nLang = nLang - 1
    nLC = nLC + 1
    
    pColorR = pColorR + nR
    pColorG = pColorG + nG
    pColorB = pColorB + nB
  Wend
End Sub


