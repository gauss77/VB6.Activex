VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
DefInt A-Z

Public Caption$, Tag$, ToolTipText$
Attribute Tag.VB_VarUserMemId = 1073938432
Attribute ToolTipText.VB_VarUserMemId = 1073938432
Public Enabled As Boolean, Visible As Boolean
Attribute Enabled.VB_VarUserMemId = 1073938435
Attribute Visible.VB_VarUserMemId = 1073938435
Public UseMaskColor As Boolean
Attribute UseMaskColor.VB_VarUserMemId = 1073938437
Public Checked As Boolean
Attribute Checked.VB_VarUserMemId = 1073938438
Public MaskColor As OLE_COLOR
Attribute MaskColor.VB_VarUserMemId = 1073938439
Public ClientLeft As Single, ClientHeight As Single
Attribute ClientLeft.VB_VarUserMemId = 1073938440
Attribute ClientHeight.VB_VarUserMemId = 1073938440
Public ClientWidth As Single, ClientTop As Single
Attribute ClientWidth.VB_VarUserMemId = 1073938442
Attribute ClientTop.VB_VarUserMemId = 1073938442
Public TextClientLeft As Single, TextClientHeight As Single
Attribute TextClientLeft.VB_VarUserMemId = 1073938444
Attribute TextClientHeight.VB_VarUserMemId = 1073938444
Public TextClientWidth As Single, TextClientTop As Single
Attribute TextClientWidth.VB_VarUserMemId = 1073938446
Attribute TextClientTop.VB_VarUserMemId = 1073938446
Public Key$, Description$
Attribute Key.VB_VarUserMemId = 1073938448
Attribute Description.VB_VarUserMemId = 1073938448
Public Style As eTBButtonStyles
Attribute Style.VB_VarUserMemId = 1073938450
Public PlaceholderSize As Single
Attribute PlaceholderSize.VB_VarUserMemId = 1073938451
Public AlwaysShowCaption As Boolean
Attribute AlwaysShowCaption.VB_VarUserMemId = 1073938452
Public GroupID
Attribute GroupID.VB_VarUserMemId = 1073938453

Public TemporaryPicture As StdPicture
Attribute TemporaryPicture.VB_VarUserMemId = 1073938454

Dim m_Picture(2) As StdPicture
Attribute m_Picture.VB_VarUserMemId = 1073938455
'----------------------------------------------------------------------
'Name        : APicture
'Created     : 12/07/1999 14:30
'Modified    :
'Modified By :
'----------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: Ariad Software
'----------------------------------------------------------------------
Public Property Get APicture(WhichPicture As eTBPictures) As StdPicture
  Set APicture = m_Picture(WhichPicture)
End Property  '(Public) Property Get APicture () As StdPicture

Property Set APicture(WhichPicture As eTBPictures, ByVal APicture As StdPicture)
  Set m_Picture(WhichPicture) = APicture
End Property  ' Property Set APicture

Private Sub Class_Terminate()
  Dim I
  For I = 0 To 2
    Set m_Picture(I) = Nothing
  Next
  Set TemporaryPicture = Nothing
End Sub


