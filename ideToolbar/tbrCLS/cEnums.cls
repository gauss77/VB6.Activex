VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefInt A-Z

Public Enum eTBClickStyles
  '##EN The <B>eTBClickStyles</B> enum determines how the buttons on an <B>ideToolbar</B> react when clicked.
  csToolbar
  '##EM csToolbar (Default) Buttons stay down when pushed
  csNoStick
  '##EM csNoStick Buttons do not stay down when pushed
End Enum

Public Enum eTBBorderStyles
  '##EN The <B>eTBBorderStyles</B> enum determines how borders are drawn.
  bsNone
  '##EM bsNone No border
  bsInset
  '##EM bsInset A inset or lowered border
  bsRaised
  '##EM bsRaised (Default) A raised border
  bsFrame
  '##EM bsFrame A combination raised and lowered border, similar to the standard Frame control
  bsInsetButton
  '##EM bsInsetButton A double width lowered border
  bsRaisedButton
  '##EM bsRaisedButton A double width raised border
End Enum

Public Enum eTBButtonStyles
  '##EN The <B>eTBButtonStyles</B> enum determines how a button on an <B>ideToolbar</B> control is drawn and behaves
  bsPlaceholder
  '##EM bsPlaceholder A holder for another control - no visible button elements
  bsButton
  '##EM bsButton (Default) A standard clickable button
  bsSeparator
  '##EM bsSeparator A separator between buttons
End Enum

Public Enum eTBSizeStyles
  '##EN The <B>eTBSizeStyles</B> enum determines how buttons on an <B>ideToolbar</B> control are sized
  ssVariable
  '##EM ssVariable (Default) Button size varies depending on caption text and picutre settings
  ssFixed
  '##EM ssFixed Button size is fixed according to object settings
End Enum

Public Enum eTBAppearances
  '##EN The <B>eTBAppearances</B> enum determines how borders are drawn on a control
  apStandard
  '##EM apStandard (Default) Objects are drawn using the default 3D style
  apFlat
  '##EM apFlat Objects are drawn using a flat 3D style, generaly only when the mouse is hovered over the object
End Enum

Public Enum eTBOrientations
  '##EN The <B>IFCBorderStyles</B> enum determines how objects are drawn on a control
  orHorizontal
  '##EM orHorizontal (Default) Drawn objects are positioned horizontally.
  orVertical
  '##EM orVertical Drawn objects are positioned vertically.
End Enum

Public Enum eTBPictures
  '##EN The <B>eTBPictures</B> enum determines which pictures are used in a function or method.
  piNormal
  '##EM piNormal (Default) Images are displayed from the Picture property
  piOver
  '##EM piOver Images are displayed from the PictureOver property
  piDown
  '##EM piDown Images are displayed from the PictureDown property
End Enum

Public Enum eTBCaptionOptions
  coNoLabels
  '##EM coNoLabels No labels are displayed
  coShowLabels
  '##EM coShowLabels (Default) All labels are displayed
  coSelectedLabels
  '##EM coSelectedLabels Only certain labels are displayed
End Enum

Public Enum eTBCaptionAlignments
  caOnTop
  '##EM caOnTop Caption is displayed above button graphic
  caOnBottom
  '##EM caOnBottom (Default) Caption is displayed below button graphic
  caOnLeft
  '##EM caOnLeft Caption is displayed to left of button graphic
  caOnRight
  '##EM caOnRight Caption is displayed to the right of button graphic
End Enum

Public Enum eTBBackStyles
  bsTransparent
  '##EM bsTransparent Background color and any graphics are visible behind the control.
  bsOpaque
  '##EM bsOpaque (Default) The control's BackColor property setting fills the control and obscures any color or graphics behind it.
End Enum

