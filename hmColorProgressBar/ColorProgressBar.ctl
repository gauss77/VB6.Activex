VERSION 5.00
Begin VB.UserControl ColorProgressBar 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   ClipControls    =   0   'False
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
   ToolboxBitmap   =   "ColorProgressBar.ctx":0000
End
Attribute VB_Name = "ColorProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Control enum types for color progress bar
Enum ProgressDirection
    cpbUp = 1
    cpbDown = 2
    cpbLeft = 3
    cpbRight = 4
End Enum
Enum BorderStyle
    cpbNone = 0
    cpbFixedSingle = 1
End Enum
'Default Property Values:
Const m_def_RedrawAtColorChange = True
Const m_def_Max = 100
Const m_def_Min = 0
Const m_def_ProgressDirection = cpbRight
Const m_def_BarColor = &HFF&
Const m_def_Value = 0
'Property Variables:
Dim m_RedrawAtColorChange As Boolean
Dim m_Max As Double
Dim m_Min As Double
Dim m_ProgressDirection As ProgressDirection
Dim m_BarColor As OLE_COLOR
Dim m_Value As Double
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event ValueChanged()

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If UserControl.BackColor() = New_BackColor Then Exit Property
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If UserControl.Enabled() = New_Enabled Then Exit Property
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    If UserControl.BorderStyle() = New_BorderStyle Then Exit Property
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    ScaleBar
    RedrawBar
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    If UserControl.MousePointer() = New_MousePointer Then Exit Property
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 330)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 2655)
    m_BarColor = PropBag.ReadProperty("BarColor", m_def_BarColor)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_ProgressDirection = PropBag.ReadProperty("ProgressDirection", m_def_ProgressDirection)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    m_RedrawAtColorChange = PropBag.ReadProperty("RedrawAtColorChange", m_def_RedrawAtColorChange)
End Sub

Private Sub UserControl_Resize()
    ScaleBar
    RedrawBar
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 330)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 2655)
    Call PropBag.WriteProperty("BarColor", m_BarColor, m_def_BarColor)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("ProgressDirection", m_ProgressDirection, m_def_ProgressDirection)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, True)
    Call PropBag.WriteProperty("RedrawAtColorChange", m_RedrawAtColorChange, m_def_RedrawAtColorChange)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Returns/sets the OLE_COLOR  the progress bar will use to draw progress."
    BarColor = m_BarColor
End Property

Public Property Let BarColor(ByVal New_BarColor As OLE_COLOR)
    If m_BarColor = New_BarColor Then Exit Property
    m_BarColor = New_BarColor
    PropertyChanged "BarColor"
    If m_RedrawAtColorChange Then RedrawBar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Value() As Double
Attribute Value.VB_Description = "Returns/sets the value of progress within the progress bar."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Double)
    Static OldValue  As Double
    If New_Value < m_Min Or New_Value > m_Max Or m_Value = New_Value Then Exit Property
    OldValue = m_Value
    m_Value = New_Value
    PropertyChanged "Value"
    RedrawBar OldValue
    RaiseEvent ValueChanged
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BarColor = m_def_BarColor
    m_Value = m_def_Value
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_ProgressDirection = m_def_ProgressDirection
    UserControl.AutoRedraw = False 'required to make it visible!
    m_RedrawAtColorChange = m_def_RedrawAtColorChange
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,0,100
Public Property Get Max() As Double
Attribute Max.VB_Description = "Returns/sets the maximum value the progress bar will have."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Double)
    If m_Max = New_Max Or New_Max = m_Min Then Exit Property
    If Not ScaleBar(New_Max) Then Exit Property
    m_Max = New_Max
    If m_Max < m_Value Then m_Value = m_Max
    RedrawBar
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=4,0,0,0
Public Property Get Min() As Double
Attribute Min.VB_Description = "Returns/sets the minimum value the progress bar will have."
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Double)
    If m_Min = New_Min Or New_Min - m_Max = 0 Then Exit Property
    If Not ScaleBar(, New_Min) Then Exit Property
    m_Min = New_Min
    If m_Value < m_Min Then m_Value = m_Min
    RedrawBar
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,4
Public Property Get ProgressDirection() As ProgressDirection
Attribute ProgressDirection.VB_Description = "Returns/sets the direction progress will grow."
    ProgressDirection = m_ProgressDirection
End Property

Public Property Let ProgressDirection(ByVal New_ProgressDirection As ProgressDirection)
    If m_ProgressDirection = New_ProgressDirection Then Exit Property
    m_ProgressDirection = New_ProgressDirection
    PropertyChanged "ProgressDirection"
    ScaleBar
    RedrawBar
End Property

Private Sub RedrawBar(Optional OldValue As Double)
    If m_Value = m_Min Then
        UserControl.Cls
        Exit Sub
    End If
    AddDeleteChunk OldValue
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_Value As Boolean)
    UserControl.AutoRedraw = New_Value
End Property

Private Function ScaleBar(Optional NewMax As Variant, Optional NewMin As Variant) As Boolean
    On Error GoTo cancel 'this will solve the division by zero problem
    Static TempMax As Double
    Static TempMin As Double
    If IsMissing(NewMax) Then
        TempMax = m_Max
    Else
        TempMax = NewMax
    End If
    If IsMissing(NewMin) Then
        TempMin = m_Min
    Else
        TempMin = NewMin
    End If
    UserControl.Cls
    Select Case m_ProgressDirection
        Case cpbUp
            UserControl.Scale (0, TempMax)-(UserControl.Width, TempMin)
        Case cpbDown
            UserControl.Scale (0, TempMin)-(UserControl.Width, TempMax)
        Case cpbLeft
            UserControl.Scale (TempMax, 0)-(TempMin, UserControl.Height)
        Case cpbRight
            UserControl.Scale (TempMin, 0)-(TempMax, UserControl.Height)
    End Select
    ScaleBar = True
    Exit Function
cancel:
    ScaleBar = False
End Function

Private Sub AddDeleteChunk(OldValue As Double)
    Select Case m_ProgressDirection
        Case cpbUp
            If m_Value < OldValue Then
                UserControl.Line (0, m_Value)-(UserControl.Width, OldValue), UserControl.BackColor, BF
            Else
                UserControl.Line (0, OldValue)-(UserControl.Width, m_Value), m_BarColor, BF
            End If
        Case cpbDown
            If m_Value < OldValue Then
                UserControl.Line (0, m_Value)-(UserControl.Width, OldValue), UserControl.BackColor, BF
            Else
                UserControl.Line (0, OldValue)-(UserControl.Width, m_Value), m_BarColor, BF
            End If
        Case cpbLeft
            If m_Value < OldValue Then
                UserControl.Line (OldValue, 0)-(m_Value, UserControl.Height), UserControl.BackColor, BF
            Else
                UserControl.Line (m_Value, 0)-(OldValue, UserControl.Height), m_BarColor, BF
            End If
        Case cpbRight
            If m_Value < OldValue Then
                UserControl.Line (m_Value, 0)-(OldValue, UserControl.Height), UserControl.BackColor, BF
            Else
                UserControl.Line (OldValue, 0)-(m_Value, UserControl.Height), m_BarColor, BF
            End If
    End Select
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get RedrawAtColorChange() As Boolean
Attribute RedrawAtColorChange.VB_Description = "Returns/sets whether or not  the progress bar should be redrawn when the bar color is changed.  If false, the bar can show bands."
    RedrawAtColorChange = m_RedrawAtColorChange
End Property

Public Property Let RedrawAtColorChange(ByVal New_RedrawAtColorChange As Boolean)
    If m_RedrawAtColorChange = New_RedrawAtColorChange Then Exit Property
    m_RedrawAtColorChange = New_RedrawAtColorChange
    PropertyChanged "RedrawAtColorChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

