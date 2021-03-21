VERSION 5.00
Object = "{EADE62FD-5B6B-444E-A6C6-26CFE520CF78}#1.0#0"; "ideToolBar.ocx"
Begin VB.Form frmToolbarDemonstration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toolbar Demonstration"
   ClientHeight    =   4650
   ClientLeft      =   1485
   ClientTop       =   1455
   ClientWidth     =   6870
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
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkBorderTop 
      Caption         =   "Border&Top"
      Height          =   285
      Left            =   2700
      TabIndex        =   19
      Top             =   630
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.CheckBox chkBorderLeft 
      Caption         =   "Border&Left"
      Height          =   285
      Left            =   2700
      TabIndex        =   18
      Top             =   945
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.CheckBox chkBorderRight 
      Caption         =   "Border&Right"
      Height          =   285
      Left            =   2700
      TabIndex        =   17
      Top             =   1260
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.CheckBox chkBorderBottom 
      Caption         =   "Border&Bottom"
      Height          =   285
      Left            =   2700
      TabIndex        =   16
      Top             =   1575
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox cboBorderStyle 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1530
      Width           =   1725
   End
   Begin VB.CheckBox chkDoubleTopBorder 
      Caption         =   "Dou&bleTopBorder"
      Height          =   285
      Left            =   4770
      TabIndex        =   14
      Top             =   630
      Width           =   1965
   End
   Begin VB.CheckBox chkDoubleBottomBorder 
      Caption         =   "DoubleBotto&mBorder"
      Height          =   285
      Left            =   4770
      TabIndex        =   13
      Top             =   945
      Width           =   1965
   End
   Begin VB.ComboBox cboAppearance 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   855
      Width           =   1725
   End
   Begin VB.ComboBox cboCaption 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2205
      Width           =   1725
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "BoldOn&Checked"
      Height          =   285
      Left            =   2700
      TabIndex        =   10
      Top             =   2025
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.CheckBox chkShowSeparators 
      Caption         =   "Sho&wSeparators"
      Height          =   285
      Left            =   2700
      TabIndex        =   9
      Top             =   2340
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Left            =   630
      TabIndex        =   8
      Top             =   3555
      Width           =   600
   End
   Begin VB.ComboBox cboAlignment 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2880
      Width           =   1725
   End
   Begin VB.CheckBox chkAutoSize 
      Caption         =   "AutoSi&ze"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2700
      TabIndex        =   6
      Top             =   3105
      Width           =   1965
   End
   Begin VB.CheckBox chkDisabled 
      Caption         =   "DisabledText&3D"
      Height          =   285
      Left            =   2700
      TabIndex        =   5
      Top             =   3510
      Width           =   1965
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "&Enabled"
      Height          =   285
      Left            =   2700
      TabIndex        =   4
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4230
      Width           =   1725
   End
   Begin VB.CheckBox chkSolidChecked 
      Caption         =   "SolidChec&ked"
      Height          =   285
      Left            =   2700
      TabIndex        =   2
      Top             =   2655
      Width           =   1965
   End
   Begin VB.CheckBox chkShowToolTips 
      Caption         =   "Show&ToolTips"
      Height          =   240
      Left            =   4770
      TabIndex        =   1
      Top             =   2340
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox chkSounds 
      Caption         =   "Play&Sounds"
      Height          =   240
      Left            =   4770
      TabIndex        =   0
      Top             =   2025
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin Insignia_Toolbar.ideToolbar tbrFontDemo 
      Height          =   465
      Left            =   5355
      Top             =   4050
      Width           =   1410
      _extentx        =   2487
      _extenty        =   820
      tooltipfont     =   "Test2.frx":0000
      solidchecked    =   -1
      buttongap       =   6
      borderstyle     =   3
      font            =   "Test2.frx":002E
      buttoncount     =   3
      showseparators  =   -1
      buttonchecked1  =   -1
      buttonkey1      =   "Left"
      buttontooltiptext1=   "Left"
      buttongroupid1  =   1
      buttonkey2      =   "Center"
      buttontooltiptext2=   "Center"
      buttongroupid2  =   1
      buttonkey3      =   "Right"
      buttontooltiptext3=   "Right"
      buttongroupid3  =   1
   End
   Begin Insignia_Toolbar.ideToolbar tbrDemo 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   6870
      _extentx        =   12118
      _extenty        =   900
      fixedsize       =   32
      tooltipfont     =   "Test2.frx":0056
      textcolor       =   -2147483641
      font            =   "Test2.frx":0084
      buttoncount     =   15
      captionoptions  =   2
      buttoncaption1  =   "Back"
      buttondescription1=   "Display previous page from History"
      buttontooltiptext1=   "Back"
      buttoncaption2  =   "Next"
      buttondescription2=   "Display next page from history"
      buttontooltiptext2=   "Next"
      buttondescription3=   "Stop loading a page"
      buttontooltiptext3=   "Stop"
      buttoncaption4  =   "Refresh"
      buttondescription4=   "Refresh the current page"
      buttontooltiptext4=   "Refresh"
      buttonalwaysshowcaption4=   -1
      buttondescription5=   "Displays your home page"
      buttontooltiptext5=   "Home"
      buttonstyle6    =   2
      buttonenabled7  =   0
      buttondescription7=   "Displays a search engine"
      buttonkey7      =   "Search"
      buttontooltiptext7=   "Search"
      buttondescription8=   "Displays your favourites menu"
      buttonkey8      =   "Fav"
      buttontooltiptext8=   "Favourites"
      buttondescription9=   "Displays your history list"
      buttonkey9      =   "History"
      buttontooltiptext9=   "History"
      buttonstyle10   =   2
      buttondescription11=   "Allows you to set options"
      buttontooltiptext11=   "Options"
      buttonstyle12   =   2
      buttondescription13=   "Displays the page full screen"
      buttontooltiptext13=   "Full Screen"
      buttondescription14=   "Allows the page to be edited"
      buttontooltiptext14=   "Edit"
      buttonstyle15   =   2
   End
   Begin Insignia_Toolbar.ideToolbar tbrHeader 
      Align           =   3  'Align Left
      Height          =   4140
      Left            =   0
      Top             =   510
      Width           =   510
      _extentx        =   900
      _extenty        =   7303
      tooltipfont     =   "Test2.frx":00AC
      textcolor       =   0
      textdisabledcolor=   49344
      buttongap       =   6
      backcolor       =   16761024
      highlightcolor  =   16777215
      shadowcolor     =   4210752
      highlightdarkcolor=   0
      shadowdarkcolor =   0
      font            =   "Test2.frx":00DA
      buttoncount     =   3
      showseparators  =   -1
      hottracking     =   -1
      hottrackingcolor=   65535
      buttoncaption1  =   "Code"
      buttonkey1      =   "Code"
      buttontooltiptext1=   "Display Code Module"
      buttongroupid1  =   1
      buttoncaption2  =   "Notes"
      buttonkey2      =   "Notes"
      buttontooltiptext2=   "Display Code Description"
      buttongroupid2  =   1
      buttonkey3      =   "Example"
      buttontooltiptext3=   "Display examples on using the code"
      buttongroupid3  =   1
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "Border&Style:"
      Height          =   195
      Index           =   0
      Left            =   630
      TabIndex        =   25
      Top             =   1305
      Width           =   855
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "&Appearance:"
      Height          =   195
      Index           =   1
      Left            =   630
      TabIndex        =   24
      Top             =   630
      Width           =   915
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "Caption&Options:"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   23
      Top             =   1980
      Width           =   1125
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "Button&Gap:"
      Height          =   195
      Index           =   3
      Left            =   630
      TabIndex        =   22
      Top             =   3330
      Width           =   810
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "CaptionAlignm&ent:"
      Height          =   195
      Index           =   4
      Left            =   630
      TabIndex        =   21
      Top             =   2655
      Width           =   1275
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      Caption         =   "St&yle:"
      Height          =   195
      Index           =   5
      Left            =   630
      TabIndex        =   20
      Top             =   4005
      Width           =   390
   End
End
Attribute VB_Name = "frmToolbarDemonstration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z


Private Sub cboAlignment_Click()
 tbrDemo.CaptionAlignment = cboAlignment.ListIndex
End Sub


Private Sub cboAppearance_Click()
 tbrDemo.Appearance = cboAppearance.ListIndex
 tbrFontDemo.Appearance = cboAppearance.ListIndex
 tbrHeader.Appearance = cboAppearance.ListIndex
End Sub


Private Sub cboBorderStyle_Click()
 tbrDemo.BorderStyle = cboBorderStyle.ListIndex
End Sub

Private Sub cboCaption_Click()
 tbrDemo.CaptionOptions = cboCaption.ListIndex
End Sub


Private Sub cboStyle_Click()
 tbrDemo.Style = cboStyle.ListIndex
End Sub


Private Sub chkBold_Click()
 tbrHeader.BoldOnChecked = chkBold
End Sub

Private Sub chkBorderBottom_Click()
 tbrDemo.BorderBottom = chkBorderBottom
End Sub

Private Sub chkBorderLeft_Click()
 tbrDemo.BorderLeft = chkBorderLeft
End Sub

Private Sub chkBorderRight_Click()
 tbrDemo.BorderRight = chkBorderRight
End Sub


Private Sub chkBorderTop_Click()
 tbrDemo.BorderTop = chkBorderTop
End Sub


Private Sub chkDisabled_Click()
 tbrDemo.DisabledText3D = chkDisabled
End Sub

Private Sub chkDoubleBottomBorder_Click()
 tbrDemo.DoubleBottomBorder = chkDoubleBottomBorder
End Sub

Private Sub chkDoubleTopBorder_Click()
 tbrDemo.DoubleTopBorder = chkDoubleTopBorder
End Sub


Private Sub chkEnabled_Click()
 tbrDemo.Enabled = chkEnabled
 tbrFontDemo.Enabled = chkEnabled
 tbrHeader.Enabled = chkEnabled
End Sub


Private Sub chkShowSeparators_Click()
 tbrHeader.ShowSeparators = chkShowSeparators
End Sub

Private Sub chkShowToolTips_Click()
 tbrDemo.ShowToolTips = chkShowToolTips
End Sub

Private Sub chkSolidChecked_Click()
 tbrHeader.SolidChecked = chkSolidChecked
End Sub


Private Sub chkSounds_Click()
 tbrDemo.PlaySounds = chkSounds
End Sub

Private Sub Form_Load()
 cboAppearance.ListIndex = apStandard
 cboBorderStyle.ListIndex = bsRaised
 cboCaption.ListIndex = coSelectedLabels
 cboAlignment.ListIndex = caOnRight
 cboStyle.ListIndex = ssVariable
 txtField = tbrDemo.ButtonGap

End Sub


Private Sub tbrDemo_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
 If ButtonKey = "Search" Or ButtonKey = "Fav" Or ButtonKey = "History" Then
  tbrDemo.ButtonChecked(ButtonIndex) = Not tbrDemo.ButtonChecked(ButtonIndex)
 End If
End Sub


Private Sub txtField_Change()
 tbrDemo.ButtonGap = Val(txtField)
End Sub


