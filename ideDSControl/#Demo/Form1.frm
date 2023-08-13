VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   16755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7680
      Left            =   60
      TabIndex        =   0
      Top             =   165
      Width           =   16155
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   12570
         TabIndex        =   66
         Text            =   "HELIOMAR 0123456789"
         Top             =   5910
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   12570
         TabIndex        =   65
         Text            =   "Heliomar"
         Top             =   5430
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   2835
         TabIndex        =   64
         Text            =   "HELIOMAR 0123456789"
         Top             =   5925
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   10125
         TabIndex        =   63
         Text            =   "HELIOMAR 0123456789"
         Top             =   5925
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   7695
         TabIndex        =   62
         Text            =   "HELIOMAR 0123456789"
         Top             =   5925
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   5265
         TabIndex        =   61
         Text            =   "HELIOMAR 0123456789"
         Top             =   5925
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   405
         TabIndex        =   60
         Text            =   "HELIOMAR 0123456789"
         Top             =   5925
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2835
         TabIndex        =   59
         Text            =   "Heliomar"
         Top             =   5445
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   10125
         TabIndex        =   58
         Text            =   "Heliomar"
         Top             =   5445
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7695
         TabIndex        =   57
         Text            =   "Heliomar"
         Top             =   5445
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   5265
         TabIndex        =   56
         Text            =   "Heliomar"
         Top             =   5445
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   390
         TabIndex        =   55
         Text            =   "Heliomar"
         Top             =   5445
         Width           =   2280
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   53
         Left            =   375
         TabIndex        =   54
         Top             =   2865
         Width           =   540
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   52
         Left            =   375
         TabIndex        =   53
         Top             =   2415
         Width           =   1185
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   51
         Left            =   375
         TabIndex        =   52
         Top             =   1515
         Width           =   585
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   50
         Left            =   375
         TabIndex        =   51
         Top             =   5130
         Width           =   1110
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG / IE"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   49
         Left            =   375
         TabIndex        =   50
         Top             =   4665
         Width           =   510
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   375
         TabIndex        =   49
         Top             =   3765
         Width           =   495
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   375
         TabIndex        =   48
         Top             =   3315
         Width           =   765
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/R. Social:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   46
         Left            =   375
         TabIndex        =   47
         Top             =   1965
         Width           =   1200
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   375
         TabIndex        =   46
         Top             =   4215
         Width           =   585
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   44
         Left            =   2850
         TabIndex        =   45
         Top             =   2835
         Width           =   570
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   43
         Left            =   2850
         TabIndex        =   44
         Top             =   2385
         Width           =   1260
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   240
         Index           =   42
         Left            =   2850
         TabIndex        =   43
         Top             =   1485
         Width           =   600
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   41
         Left            =   2850
         TabIndex        =   42
         Top             =   5100
         Width           =   1200
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG / IE"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   40
         Left            =   2850
         TabIndex        =   41
         Top             =   4635
         Width           =   555
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   39
         Left            =   2850
         TabIndex        =   40
         Top             =   3735
         Width           =   480
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   38
         Left            =   2850
         TabIndex        =   39
         Top             =   3285
         Width           =   825
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/R. Social:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   240
         Index           =   37
         Left            =   2850
         TabIndex        =   38
         Top             =   1935
         Width           =   1275
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   36
         Left            =   2850
         TabIndex        =   37
         Top             =   4185
         Width           =   675
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   35
         Left            =   5265
         TabIndex        =   36
         Top             =   2865
         Width           =   600
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   34
         Left            =   5265
         TabIndex        =   35
         Top             =   2415
         Width           =   1335
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   33
         Left            =   5265
         TabIndex        =   34
         Top             =   1515
         Width           =   675
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   5265
         TabIndex        =   33
         Top             =   5160
         Width           =   1440
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG / IE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   5265
         TabIndex        =   32
         Top             =   4665
         Width           =   630
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   5265
         TabIndex        =   31
         Top             =   3765
         Width           =   600
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   5265
         TabIndex        =   30
         Top             =   3315
         Width           =   870
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/R. Social:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   28
         Left            =   5265
         TabIndex        =   29
         Top             =   1965
         Width           =   1395
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   5265
         TabIndex        =   28
         Top             =   4215
         Width           =   675
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   7695
         TabIndex        =   27
         Top             =   2895
         Width           =   495
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   7695
         TabIndex        =   26
         Top             =   2445
         Width           =   1125
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   24
         Left            =   7695
         TabIndex        =   25
         Top             =   1545
         Width           =   570
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   7695
         TabIndex        =   24
         Top             =   5145
         Width           =   1230
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG / IE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   7695
         TabIndex        =   23
         Top             =   4695
         Width           =   510
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   7695
         TabIndex        =   22
         Top             =   3795
         Width           =   480
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   7695
         TabIndex        =   21
         Top             =   3345
         Width           =   735
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/R. Social:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   19
         Left            =   7695
         TabIndex        =   20
         Top             =   1995
         Width           =   1140
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   7695
         TabIndex        =   19
         Top             =   4245
         Width           =   555
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   10110
         TabIndex        =   18
         Top             =   2835
         Width           =   630
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   10110
         TabIndex        =   17
         Top             =   2385
         Width           =   1470
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   12
         Left            =   10110
         TabIndex        =   16
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações..."
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   10125
         TabIndex        =   15
         Top             =   5130
         Width           =   1470
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG / IE"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   10110
         TabIndex        =   14
         Top             =   4635
         Width           =   735
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   10110
         TabIndex        =   13
         Top             =   3735
         Width           =   735
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   10110
         TabIndex        =   12
         Top             =   3285
         Width           =   945
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/R. Social:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   1
         Left            =   10110
         TabIndex        =   11
         Top             =   1935
         Width           =   1575
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Hack"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10110
         TabIndex        =   10
         Top             =   4185
         Width           =   735
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   12615
         TabIndex        =   9
         Top             =   2910
         Width           =   480
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   12615
         TabIndex        =   8
         Top             =   2460
         Width           =   1110
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   4
         Left            =   12600
         TabIndex        =   7
         Top             =   1545
         Width           =   570
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações..."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   12615
         TabIndex        =   6
         Top             =   5145
         Width           =   1290
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG / IE"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   12615
         TabIndex        =   5
         Top             =   4710
         Width           =   555
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   12615
         TabIndex        =   4
         Top             =   3810
         Width           =   450
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   12615
         TabIndex        =   3
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome/R. Social:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   195
         Index           =   2
         Left            =   12615
         TabIndex        =   2
         Top             =   2010
         Width           =   1185
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   12615
         TabIndex        =   1
         Top             =   4260
         Width           =   540
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

