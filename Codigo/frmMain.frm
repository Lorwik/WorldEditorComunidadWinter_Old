VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor"
   ClientHeight    =   10755
   ClientLeft      =   390
   ClientTop       =   840
   ClientWidth     =   19200
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   717
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1280
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   91
      Top             =   0
      Width           =   4575
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   0
         Left            =   70
         TabIndex        =   92
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":24A16
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   93
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":25668
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   2
         Left            =   890
         TabIndex        =   94
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":262BA
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   3
         Left            =   1290
         TabIndex        =   95
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":26F0C
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   4
         Left            =   2100
         TabIndex        =   96
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   5
         Left            =   2500
         TabIndex        =   97
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   6
         Left            =   2900
         TabIndex        =   98
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   7
         Left            =   3300
         TabIndex        =   99
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   8
         Left            =   3700
         TabIndex        =   100
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "G"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   9
         Left            =   1700
         TabIndex        =   114
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":27B5E
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   10
         Left            =   4110
         TabIndex        =   119
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "Tr"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.PictureBox PreviewGrh 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   120
      ScaleHeight     =   3510
      ScaleWidth      =   4425
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   7140
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   4680
      ScaleHeight     =   617
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   965
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1440
      Width           =   14475
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   20400
      Top             =   0
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      Picture         =   "frmMain.frx":281E0
      ScaleHeight     =   6225
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin VB.Frame FraRellenar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Rellenar"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   101
         Top             =   4320
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox DX1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   105
            Text            =   "1"
            Top             =   220
            Width           =   495
         End
         Begin VB.TextBox DX2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   104
            Text            =   "5"
            Top             =   220
            Width           =   495
         End
         Begin VB.TextBox DY1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   103
            Text            =   "1"
            Top             =   220
            Width           =   495
         End
         Begin VB.TextBox DY2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   102
            Text            =   "5"
            Top             =   220
            Width           =   495
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   106
            Top             =   560
            Width           =   1215
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Pintar Area"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   107
            Top             =   950
            Width           =   1455
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Quitar Bloqueos"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   108
            Top             =   560
            Width           =   1455
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Insertar Bloqueos"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   109
            Top             =   950
            Width           =   1215
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Quitar Area"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   4
            Left            =   3000
            TabIndex        =   115
            Top             =   560
            Width           =   1095
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "Ins. Trigger"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   5
            Left            =   3000
            TabIndex        =   116
            Top             =   950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Quit. Trigger"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   117
            Top             =   1335
            Width           =   1935
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Insertar Zona"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   7
            Left            =   2160
            TabIndex        =   118
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "Quitar Zona"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label lblY2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y2:"
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
            Left            =   3195
            TabIndex        =   113
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblY1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y1:"
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
            Left            =   2160
            TabIndex        =   112
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblX2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X2:"
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
            Index           =   1
            Left            =   1275
            TabIndex        =   111
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblX2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X1:"
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
            Index           =   0
            Left            =   240
            TabIndex        =   110
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   56
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   55
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   54
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":821AA
         Left            =   1080
         List            =   "frmMain.frx":821BA
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   43
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   42
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":821CA
         Left            =   120
         List            =   "frmMain.frx":821CC
         Sorted          =   -1  'True
         TabIndex        =   41
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   46
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":821CE
         Left            =   3360
         List            =   "frmMain.frx":821D0
         TabIndex        =   37
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":821D2
         Left            =   840
         List            =   "frmMain.frx":821D4
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":821D6
         Left            =   120
         List            =   "frmMain.frx":821D8
         TabIndex        =   36
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":821DA
         Left            =   840
         List            =   "frmMain.frx":821DC
         TabIndex        =   28
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":821DE
         Left            =   3360
         List            =   "frmMain.frx":821E0
         TabIndex        =   27
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   26
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":821E2
         Left            =   120
         List            =   "frmMain.frx":821E4
         TabIndex        =   25
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":821E6
         Left            =   120
         List            =   "frmMain.frx":821E8
         TabIndex        =   24
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   3
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   4
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   20
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   735
         Left            =   2400
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":821EA
         Left            =   840
         List            =   "frmMain.frx":821EC
         TabIndex        =   47
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   48
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":821EE
         Left            =   120
         List            =   "frmMain.frx":821F0
         TabIndex        =   49
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":821F2
         Left            =   3360
         List            =   "frmMain.frx":821F4
         TabIndex        =   50
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   64
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   63
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   62
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   17
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   19920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox StatTxt 
      Height          =   1425
      Left            =   14640
      TabIndex        =   67
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   0
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   2514
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":821F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   10245
      TabIndex        =   81
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":82273
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   9135
      TabIndex        =   82
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":82839
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8025
      TabIndex        =   83
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":82D3A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   6915
      TabIndex        =   84
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":830EE
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5805
      TabIndex        =   85
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":8346F
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1050
      Index           =   0
      Left            =   4725
      TabIndex        =   86
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":86ACF
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   8040
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":8A015
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   11340
      TabIndex        =   88
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&Copiar Bordes"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":8A3C9
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   12450
      TabIndex        =   89
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "&Particulas"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":8AA0A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   9
      Left            =   13560
      TabIndex        =   90
      Top             =   0
      Width           =   1050
      _ExtentX        =   132
      _ExtentY        =   1826
      Caption         =   "Luces "
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":8B08C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   13860
      TabIndex        =   68
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   13095
      TabIndex        =   69
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5445
      TabIndex        =   70
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6210
      TabIndex        =   71
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6975
      TabIndex        =   72
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7740
      TabIndex        =   73
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8505
      TabIndex        =   74
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9270
      TabIndex        =   75
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10035
      TabIndex        =   76
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   10800
      TabIndex        =   77
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   11565
      TabIndex        =   78
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   79
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   12330
      TabIndex        =   80
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa (Int)"
         Index           =   1
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "Abrir Mapas IAO"
         Index           =   2
         Begin VB.Menu mnuAbrirMapaIAO 
            Caption         =   "Abrir Mapa IAO 1.4"
            Index           =   0
         End
         Begin VB.Menu mnuAbrirMapaIAO 
            Caption         =   "Abrir Mapa IAO 1.3"
            Index           =   1
         End
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarRender 
         Caption         =   "Guardar Render"
      End
      Begin VB.Menu mnuConversor 
         Caption         =   "Conversor de Mapas"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuComo 
         Caption         =   "¿ Como seleccionar ? ---- Mantener SHIFT y arrastrar el cursor."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarParticulas 
            Caption         =   "&Todas las particulas"
         End
         Begin VB.Menu mnuQuitarBloqueosBorde 
            Caption         =   "Bloqueos de los borde"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLine10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuVerParticulas 
         Caption         =   "...Parti&culas"
      End
      Begin VB.Menu mnuZonas 
         Caption         =   "...Zo&nas"
         Begin VB.Menu mnuVerZonas 
            Caption         =   "Zona Actual"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuVerZonas 
            Caption         =   "Todas las Zonas"
            Index           =   1
         End
         Begin VB.Menu mnuVerZonas 
            Caption         =   "Ocultar Zonas"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarCabezas 
            Caption         =   "C&abezas"
         End
         Begin VB.Menu mnuActualizarCuerpos 
            Caption         =   "C&uerpos"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu Minimap_ 
      Caption         =   "Minimap"
      Begin VB.Menu Minimap_capa1 
         Caption         =   "Capa 1"
      End
      Begin VB.Menu Minimap_capa2 
         Caption         =   "Capa 2"
      End
      Begin VB.Menu Minimap_capa3 
         Caption         =   "Capa 3"
      End
      Begin VB.Menu Minimap_capa4 
         Caption         =   "Capa 4"
      End
      Begin VB.Menu Minimap_npcs 
         Caption         =   "Npcs"
      End
      Begin VB.Menu Minimap_objetos 
         Caption         =   "Objetos"
      End
      Begin VB.Menu Minimap_bloqueos 
         Caption         =   "Bloqueos"
      End
      Begin VB.Menu Minimap_particulas 
         Caption         =   "Particulas"
      End
      Begin VB.Menu Minimap_ndemapa 
         Caption         =   "Nº de mapa"
      End
      Begin VB.Menu Dibujarmini 
         Caption         =   "Dibujar"
      End
      Begin VB.Menu mnuLineMinimap 
         Caption         =   "-"
      End
      Begin VB.Menu minimapSave 
         Caption         =   "Guardar Minimapa"
      End
      Begin VB.Menu SaveAllMiniMap 
         Caption         =   "Guardar todos los Minimapas"
      End
   End
   Begin VB.Menu mnuVentanas 
      Caption         =   "Ventanas"
      Begin VB.Menu mnuMapa 
         Caption         =   "Mapa"
      End
      Begin VB.Menu mnuMapInfo 
         Caption         =   "Información de Mapa / Zona"
      End
      Begin VB.Menu mnuConfigZonas 
         Caption         =   "Configuración de Zonas"
      End
      Begin VB.Menu mnuAdvSup 
         Caption         =   "Configuracion Avanzada de Superficies"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuLineAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private clicX              As Long
Private clicY              As Long

Public Resolution As Boolean

Private Sub PonerAlAzar(ByVal n As Integer, T As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 by GS
'*************************************************
    Dim ObjIndex As Long
    Dim NPCIndex As Long
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long
    Dim Head As Integer
    Dim Body As Integer
    Dim Heading As Byte
    Dim Leer As New clsIniReader
    i = n
    
    Do While i > 0
        X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
        Y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
        
        Select Case T
            Case 0
                If MapData(X, Y).OBJInfo.ObjIndex = 0 Then
                      i = i - 1
                      If cInsertarBloqueo.value = True Then
                        MapData(X, Y).Blocked = 1
                      Else
                        MapData(X, Y).Blocked = 0
                      End If
                      If cNumFunc(2).Text > 0 Then
                          ObjIndex = cNumFunc(2).Text
                          InitGrh MapData(X, Y).ObjGrh, ObjData(ObjIndex).GrhIndex
                          MapData(X, Y).OBJInfo.ObjIndex = ObjIndex
                          MapData(X, Y).OBJInfo.Amount = Val(cCantFunc(2).Text)
                          Select Case ObjData(ObjIndex).ObjType ' GS
                                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                    MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                          End Select
                      End If
                End If
            Case 1
               If MapData(X, Y).Blocked = 0 Then
                      i = i - 1
                      If cNumFunc(T - 1).Text > 0 Then
                            NPCIndex = cNumFunc(T - 1).Text
                            Body = NpcData(NPCIndex).Body
                            Head = NpcData(NPCIndex).Head
                            Heading = NpcData(NPCIndex).Heading
                            Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                            MapData(X, Y).NPCIndex = NPCIndex
                      End If
                End If
            Case 2
               If MapData(X, Y).Blocked = 0 Then
                      i = i - 1
                      If cNumFunc(T - 1).Text >= 0 Then
                            NPCIndex = cNumFunc(T - 1).Text
                            Body = NpcData(NPCIndex).Body
                            Head = NpcData(NPCIndex).Head
                            Heading = NpcData(NPCIndex).Heading
                            Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                            MapData(X, Y).NPCIndex = NPCIndex
                      End If
               End If
            End Select
            DoEvents
    Loop
End Sub

Private Sub cAgregarFuncalAzar_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    On Error Resume Next
    If IsNumeric(cCantFunc(Index).Text) = False Or cCantFunc(Index).Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub
    End If
    
    cAgregarFuncalAzar(Index).Enabled = False
    Call PonerAlAzar(CInt(cCantFunc(Index).Text), 1 + (IIf(Index = 2, -1, Index)))
    cAgregarFuncalAzar(Index).Enabled = True
End Sub

Private Sub cCantFunc_Change(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Val(cCantFunc(Index)) < 1 Then
      cCantFunc(Index).Text = 1
    End If
    If Val(cCantFunc(Index)) > 10000 Then
      cCantFunc(Index).Text = 10000
    End If
End Sub

Private Sub cCapas_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 31/05/06
'*************************************************
    If Val(cCapas.Text) < 1 Then
      cCapas.Text = 1
    End If
    If Val(cCapas.Text) > 4 Then
      cCapas.Text = 4
    End If
    cCapas.Tag = vbNullString
End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub cFiltro_GotFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    HotKeysAllow = False
End Sub

Private Sub cFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If KeyAscii = 13 Then
        Call Filtrar(Index)
    End If
End Sub

Private Sub cFiltro_LostFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    HotKeysAllow = True
End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If KeyAscii = 13 Then
        Call fPreviewGrh(cGrh.Text)
        If frmMain.cGrh.ListCount > 5 Then
            frmMain.cGrh.RemoveItem 0
        End If
        frmMain.cGrh.AddItem frmMain.cGrh.Text
    End If
    Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub cInsertarFunc_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cInsertarFunc(Index).value = True Then
        cQuitarFunc(Index).Enabled = False
        cAgregarFuncalAzar(Index).Enabled = False
        If Index <> 2 Then cCantFunc(Index).Enabled = False
        Call modPaneles.EstSelectPanel((Index) + 3, True)
    Else
        cQuitarFunc(Index).Enabled = True
        cAgregarFuncalAzar(Index).Enabled = True
        If Index <> 2 Then cCantFunc(Index).Enabled = True
        Call modPaneles.EstSelectPanel((Index) + 3, False)
    End If
End Sub

Private Sub cInsertarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
    If cInsertarTrans.value = True Then
        cQuitarTrans.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        cQuitarTrans.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)
    End If
End Sub

Private Sub cInsertarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cInsertarTrigger.value = True Then
        cQuitarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        cQuitarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)
    End If
End Sub

Private Sub cmdInformacionDelMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmMapInfo.Show
    frmMapInfo.Visible = True
End Sub

Private Sub cUnionManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    cInsertarTrans.value = (cUnionManual.value = True)
    Call cInsertarTrans_Click
End Sub

Private Sub cverBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuVerBloqueos.Checked = cVerBloqueos.value
End Sub

Private Sub cverTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub cNumFunc_KeyPress(Index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

    If KeyAscii = 13 Then
        Dim Cont As String
        Cont = frmMain.cNumFunc(Index).Text
        Call cNumFunc_LostFocus(Index)
        If Cont <> frmMain.cNumFunc(Index).Text Then Exit Sub
        If frmMain.cNumFunc(Index).ListCount > 5 Then
            frmMain.cNumFunc(Index).RemoveItem 0
        End If
        frmMain.cNumFunc(Index).AddItem frmMain.cNumFunc(Index).Text
        Exit Sub
    ElseIf KeyAscii = 8 Then
        
    ElseIf IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
        Exit Sub
    End If

End Sub

Private Sub cNumFunc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If cNumFunc(Index).Text = vbNullString Then
        frmMain.cNumFunc(Index).Text = IIf(Index = 1, 500, 1)
    End If
End Sub

Private Sub cNumFunc_LostFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    Select Case Index
    
        Case 0
            If frmMain.cNumFunc(Index).Text > 499 Or frmMain.cNumFunc(Index).Text < 1 Then
                frmMain.cNumFunc(Index).Text = 1
            End If
        
        Case 1
            If frmMain.cNumFunc(Index).Text < 500 Or frmMain.cNumFunc(Index).Text > 32000 Then
                frmMain.cNumFunc(Index).Text = 500
            End If
        
        Case 2
            If frmMain.cNumFunc(Index).Text < 1 Or frmMain.cNumFunc(Index).Text > 32000 Then
                frmMain.cNumFunc(Index).Text = 1
            End If
    End Select
End Sub

Private Sub cInsertarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    cInsertarBloqueo.Tag = vbNullString
    If cInsertarBloqueo.value = True Then
        cQuitarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
        
    Else
        cQuitarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)
        
    End If
End Sub

Private Sub cQuitarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    cInsertarBloqueo.Tag = vbNullString
    If cQuitarBloqueo.value = True Then
        cInsertarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
        
    Else
        cInsertarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)
        
    End If
End Sub

Private Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarEnEstaCapa.value = True Then
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarEnTodasLasCapas.value = True Then
        cCapas.Enabled = False
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        cCapas.Enabled = True
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If
End Sub


Private Sub cQuitarFunc_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarFunc(Index).value = True Then
        cInsertarFunc(Index).Enabled = False
        cAgregarFuncalAzar(Index).Enabled = False
        cCantFunc(Index).Enabled = False
        cNumFunc(Index).Enabled = False
        cFiltro((Index) + 1).Enabled = False
        lListado((Index) + 1).Enabled = False
        Call modPaneles.EstSelectPanel((Index) + 3, True)
        
    Else
        cInsertarFunc(Index).Enabled = True
        cAgregarFuncalAzar(Index).Enabled = True
        cCantFunc(Index).Enabled = True
        cNumFunc(Index).Enabled = True
        cFiltro((Index) + 1).Enabled = True
        lListado((Index) + 1).Enabled = True
        Call modPaneles.EstSelectPanel((Index) + 3, False)
        
    End If
End Sub

Private Sub cQuitarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarTrans.value = True Then
        cInsertarTransOBJ.Enabled = False
        cInsertarTrans.Enabled = False
        cUnionManual.Enabled = False
        cUnionAuto.Enabled = False
        tTMapa.Enabled = False
        tTX.Enabled = False
        tTY.Enabled = False
        mnuInsertarTransladosAdyasentes.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
        
    Else
        tTMapa.Enabled = True
        tTX.Enabled = True
        tTY.Enabled = True
        cUnionAuto.Enabled = True
        cUnionManual.Enabled = True
        cInsertarTrans.Enabled = True
        cInsertarTransOBJ.Enabled = True
        mnuInsertarTransladosAdyasentes.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)
        
    End If
End Sub

Private Sub cQuitarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cQuitarTrigger.value = True Then
        lListado(4).Enabled = False
        cInsertarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
        
    Else
        lListado(4).Enabled = True
        cInsertarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)
        
    End If
End Sub

Private Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If cSeleccionarSuperficie.value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
        
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)
        
    End If
End Sub

Private Sub cUnionAuto_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmUnionAdyacente.Show
End Sub

Private Sub Form_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Me.SetFocus

End Sub

Private Sub Form_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
    Dim tX As Integer
    Dim tY As Integer
    
    If Not MapaCargado Then Exit Sub
    
    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)
    End If
End Sub

Private Sub Form_Load()
'*************************************************
'Author: Lorwik
'Last modified: 23/03/2021
'*************************************************

    If ClientSetup.WeMode = eWeMode.ImperiumClasico Then
        mnuAbrirMapa(2).Visible = True
        frmMapInfo.LvBOptX(1).Enabled = False
        Call frmMapInfo.OptTam_Click(0)
        frmMapInfo.OptTam(1).Enabled = False
        
        With frmMapa.picMapa
            .Height = 100
            .Width = 100
        End With
        
        With frmMapa
            .Height = 1935
            .Width = 1590
        End With
        
        mnuConfigZonas.Visible = False
        frmMapInfo.LvBGuardar.Visible = False
        frmMapInfo.LvBGuardar.Visible = False
        LvBAreas(6).Visible = False
        LvBAreas(7).Visible = False
        
    ElseIf ClientSetup.WeMode = eWeMode.WinterAO Then
        mnuAbrirMapa(2).Visible = False
        
        If Resolution Then
            frmMapInfo.OptTam(1).value = True
            
        Else
            frmMapInfo.OptTam(0).value = True
            
        End If
        
    End If
    
    If ClientSetup.Preview = True Then
        frmConfigSup.chkActivarPreview.value = Unchecked
        
    Else
        frmConfigSup.chkActivarPreview.value = Checked
        
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
    ' HotKeys
    If HotKeysAllow = False Then Exit Sub
    
    Select Case UCase(Chr(KeyAscii))
        Case "S" ' Activa/Desactiva Insertar Superficie
            cSeleccionarSuperficie.value = (cSeleccionarSuperficie.value = False)
            Call cSeleccionarSuperficie_Click
            
        Case "T" ' Activa/Desactiva Insertar Translados
            cInsertarTrans.value = (cInsertarTrans.value = False)
            Call cInsertarTrans_Click
            
        Case "B" ' Activa/Desactiva Insertar Bloqueos
            cInsertarBloqueo.value = (cInsertarBloqueo.value = False)
            Call cInsertarBloqueo_Click
            
        Case "N" ' Activa/Desactiva Insertar NPCs
            cInsertarFunc(0).value = (cInsertarFunc(0).value = False)
            Call cInsertarFunc_Click(0)
            
       ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
       '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
       '     Call cInsertarFunc_Click(1)
       
        Case "O" ' Activa/Desactiva Insertar Objetos
            cInsertarFunc(2).value = (cInsertarFunc(2).value = False)
            Call cInsertarFunc_Click(2)
            
            
        Case "G" ' Activa/Desactiva Insertar Triggers
            cInsertarTrigger.value = (cInsertarTrigger.value = False)
            Call cInsertarTrigger_Click
            
        Case "Q" ' Quitar Funciones
            Call mnuQuitarFunciones_Click
            
    End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_Resize()
    
    With MainViewPic
    
        .Height = Me.ScaleHeight - 100
        .Width = Me.ScaleWidth - 315
    
    End With
    
    With MainScreenRect
        .Bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With
    
    Call ChangeView
    
End Sub

Private Sub lListado_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    On Error Resume Next
    If HotKeysAllow = False Then
        lListado(Index).Tag = lListado(Index).ListIndex
        Select Case Index
            Case 0
                cGrh.Text = DameGrhIndex(ReadField(2, lListado(Index).Text, Asc("#")))
                
                If SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa <> 0 Then
                    If LenB(ReadField(2, lListado(Index).Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
                    cCapas.Text = SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa
                Else
                    If LenB(cCapas.Tag) <> 0 Then
                        cCapas.Text = cCapas.Tag
                        cCapas.Tag = vbNullString
                    End If
                End If
                
                If SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Block = True Then
                    If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                    cInsertarBloqueo.value = True
                    Call cInsertarBloqueo_Click
                Else
                    If LenB(cInsertarBloqueo.Tag) <> 0 Then
                        cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                        cInsertarBloqueo.Tag = vbNullString
                        Call cInsertarBloqueo_Click
                    End If
                End If
                Call fPreviewGrh(cGrh.Text)
                Call RenderPreview
                
            Case 1
                cNumFunc(0).Text = ReadField(1, ReadField(2, lListado(Index).Text, Asc("#")), Asc(" "))
                
            Case 2
                cNumFunc(1).Text = ReadField(2, lListado(Index).Text, Asc("#"))
                
            Case 3
                cNumFunc(2).Text = ReadField(2, lListado(Index).Text, Asc("#"))
                Call fPreviewGrh(ObjData(ReadField(2, lListado(Index).Text, Asc("#"))).GrhIndex)
                Call RenderPreview
                
        End Select
    Else
        lListado(Index).ListIndex = lListado(Index).Tag
    End If

End Sub

Private Sub lListado_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    If Index = 3 And Button = 2 Then
        If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc
    End If
End Sub

Private Sub lListado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
On Error Resume Next
    HotKeysAllow = False
End Sub

Private Sub LvBAreas_Click(Index As Integer)
    If IsNumeric(DX1.Text) = False Or _
       IsNumeric(DX2.Text) = False Or _
       IsNumeric(DY1.Text) = False Or _
       IsNumeric(DY2.Text) = False Then
    
        Call MsgBox("Debes introducir valores nï¿½mericos. Estos pueden tener un mï¿½nimo de 1 y un mï¿½ximo de " & (YMinMapSize + XMinMapSize) / 2 & ".")
    
       Exit Sub
    End If
    
    Select Case Index
        Case 0
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 1
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 2
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 3
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 4
            Call Triggers_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 5
            Call Triggers_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
        Case 6
            Call Zonas_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
            
        Case 7
            Call Zonas_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
            
    End Select
End Sub

Private Sub LvBOpcion_Click(Index As Integer)

    Select Case Index
        Case 0
            cVerBloqueos.value = (cVerBloqueos.value = False)
            mnuVerBloqueos.Checked = cVerBloqueos.value
            
        Case 1
            mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
            
        Case 2
            mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
            
        Case 3
            cVerTriggers.value = (cVerTriggers.value = False)
            mnuVerTriggers.Checked = cVerTriggers.value
            
        Case 4
            mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
            
        Case 5
            mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
            
        Case 6
            mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
            
        Case 7
            mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
            
        Case 8
            VerGrilla = (VerGrilla = False)
            mnuVerGrilla.Checked = (mnuVerGrilla.Checked = False)
            
        Case 9
            VerParticulas = (VerParticulas = False)
            mnuVerParticulas.Checked = (mnuVerParticulas.Checked = False)
            
        Case 10
            AlphaTecho = Not AlphaTecho
            
    End Select
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
                                  
    Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
                                
    Call Form_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub MainViewPic_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

    If Not MapaCargado Then Exit Sub
    
    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)
        
    End If
End Sub

Private Sub MapPest_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/03/2021
'Lorwik> Ahora distingue entre csm y map
'*************************************************
    Dim Formato As String
    
    Select Case frmMain.Dialog.FilterIndex
    
        Case 1
            Formato = ".csm"
            
        Case 2
            Formato = ".map"
            
    End Select
    
    
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then _
            Call modMapIO.GuardarMapa(Dialog.filename)

    End If
        
    If (Index + NumMap_Save - 4) <> NumMap_Save Then
        Dialog.CancelError = True

        On Error GoTo ErrHandler

        Dialog.filename = PATH_Save & NameMap_Save & (Index + NumMap_Save - 7) & Formato
        
        Call modMapIO.NuevoMapa
        
        DoEvents
        Select Case frmMain.Dialog.FilterIndex
        
            Case 1
                If TipoMapaCargado = eTipoMapa.tWinter Then
                    Call modMapWinter.Cargar_CSM(Dialog.filename)
                
                #If Privado = 0 Then
                ElseIf TipoMapaCargado = eTipoMapa.tIAOClasico Then
                    Call modMapImpC.Cargar_MapImpClasico(Dialog.filename)
                
                ElseIf TipoMapaCargado = eTipoMapa.tIAOnew Or TipoMapaCargado = eTipoMapa.tIAOold Then
                    Call modMapImpC.Cargar_MapIAO(Dialog.filename, TipoMapaCargado)
                    
                #End If
                End If
                
            Case 2
                If TipoMapaCargado = eTipoMapa.tInt Then
                    Call modMapIO.MapaV2_Cargar(Dialog.filename, True)
                    
                Else
                    Call modMapIO.MapaV2_Cargar(Dialog.filename)
                    
                End If
            
        End Select
        
        EngineRun = True
        
    End If
    
        Exit Sub
    
ErrHandler:
        Call MsgBox(Err.Description)
End Sub

Private Sub MiniMap_Bloqueos_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_Bloqueos_Click_Err
    
    Minimap_bloqueos.Checked = (Minimap_bloqueos.Checked = False)
    MMiniMap_Bloqueos = Not MMiniMap_Bloqueos

    Exit Sub

MiniMap_Bloqueos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_Bloqueos_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa1_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa1_Click_Err
    
    Minimap_capa1.Checked = (Minimap_capa1.Checked = False)
    MMiniMap_capa1 = Not MMiniMap_capa1
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_capa1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa1_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa2_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa2_Click_Err
    
    Minimap_capa2.Checked = (Minimap_capa2.Checked = False)
    MMiniMap_capa2 = Not MMiniMap_capa2
    Call DibujarMinimapa

    Exit Sub

MiniMap_capa2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa2_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa3_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa3_Click_Err
    
    Minimap_capa3.Checked = (Minimap_capa3.Checked = False)
    MMiniMap_capa3 = Not MMiniMap_capa3
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_capa3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa3_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_capa4_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_capa4_Click_Err
    
    Minimap_capa4.Checked = (Minimap_capa4.Checked = False)
    MMiniMap_capa4 = Not MMiniMap_capa4
    Call DibujarMinimapa

    Exit Sub

MiniMap_capa4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_capa4_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_ndemapa_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_ndemapa_Click_Err
    
    Minimap_ndemapa.Checked = (Minimap_ndemapa.Checked = False)
    MMiniMap_Nombre = Not MMiniMap_Nombre
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_ndemapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_ndemapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_Npcs_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_Npcs_Click_Err
    
    Minimap_npcs.Checked = (Minimap_npcs.Checked = False)
    MMiniMap_Npcs = Not MMiniMap_Npcs
    Call DibujarMinimapa
    
    Exit Sub

MiniMap_Npcs_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_Npcs_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_objetos_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_objetos_Click_Err
    
    Minimap_objetos.Checked = (Minimap_objetos.Checked = False)
    MMiniMap_objetos = Not MMiniMap_objetos
    Call DibujarMinimapa

    
    Exit Sub

MiniMap_objetos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_objetos_Click", Erl)
    Resume Next
    
End Sub

Private Sub MiniMap_particulas_Click()
'*************************************************
'Author: ???
'Last modified: ???
'*************************************************
    On Error GoTo MiniMap_particulas_Click_Err
    
    Minimap_particulas.Checked = (Minimap_particulas.Checked = False)
    MMiniMap_particulas = Not MMiniMap_particulas
    Call DibujarMinimapa

    
    Exit Sub

MiniMap_particulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.MiniMap_particulas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAbrirMapa_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************
    Select Case Index
    
        Case 0
            Call AbrirMapa(False)
            
        Case 1
            Call AbrirMapa(True)
            
    End Select
    
End Sub

Private Sub mnuAbrirMapaIAO_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************
#If Privado = 0 Then
    Select Case Index
    
        Case 0
            Call modMapImpC.AbrirMapaIAO(eTipoMapa.tIAOnew)
        
        Case 1
            Call modMapImpC.AbrirMapaIAO(eTipoMapa.tIAOold)
            
    End Select
#End If
End Sub

Private Sub mnuActualizarCabezas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    Call modCarga.CargarCabezas
End Sub

Private Sub mnuActualizarCuerpos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    Call modCarga.CargarCuerpos
End Sub

Private Sub mnuActualizarGraficos_Click()
    Call modCarga.LoadGrhData
End Sub

Private Sub mnuActualizarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modCarga.CargarIndicesSuperficie
End Sub

Private Sub mnuacercade_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmAbout.Show
End Sub



Private Sub mnuActualizarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modCarga.CargarIndicesNPC
End Sub

Private Sub mnuActualizarObjs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modCarga.CargarIndicesOBJ
End Sub

Private Sub mnuActualizarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modCarga.CargarIndicesTriggers
End Sub

Private Sub mnuAdvSup_Click()
    frmConfigSup.Show , frmMain
End Sub

Private Sub mnuAutoCapturarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
    mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
    mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmAutoGuardarMapa.Show
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub

Private Sub mnuBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    For i = 0 To 6
        If i <> 2 Then
            frmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    
    modPaneles.VerFuncion 2, True
End Sub

Private Sub mnuBloquearBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Bloquear_Bordes
End Sub

Private Sub mnuBloquearMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Bloqueo_Todo(1)
End Sub

Private Sub mnuBloquearS_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    Call BlockearSeleccion
End Sub

Private Sub mnuConfigObjTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    Cfg_TrOBJ = cNumFunc(2).Text
End Sub

Private Sub mnuConfigZonas_Click()
    frmZonas.Show , frmMain
End Sub

Private Sub mnuConversor_Click()
'*************************************************
'Author: Lorwik
'Last modified: 22/03/2021
'*************************************************
    frmConvert.Show
End Sub

Private Sub mnuCopiar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    Call CopiarSeleccion
End Sub

Private Sub mnuCortar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    Call CortarSeleccion
End Sub

Private Sub mnuDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo mnuDeshacer_Click_Err
    
    Call modEdicion.Deshacer_Recover

    Exit Sub

mnuDeshacer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuDeshacer_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuDeshacerPegado_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    
    On Error GoTo mnuDeshacerPegado_Click_Err
    
    Call modEdicion.Deshacer_Add("Deshacer Pegado de Selección")
    Call DePegar
    
    Exit Sub

mnuDeshacerPegado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuDeshacerPegado_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGRHaBMP_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    frmGRHaBMP.Show
End Sub

Public Sub mnuGuardarMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error GoTo mnuGuardarMapa_Click_Err

    modMapIO.GuardarMapa Dialog.filename
    
    Exit Sub

mnuGuardarMapa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    
    On Error GoTo mnuGuardarMapaComo_Click_Err
    
    modMapIO.GuardarMapa
    
    Exit Sub

mnuGuardarMapaComo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarMapaComo_Click", Erl)
    Resume Next
End Sub

Private Sub mnuGuardarRender_Click()
'*************************************************
'Author: Lorwik
'Last modified: 22/03/2021
'*************************************************

On Error GoTo mnuGuardarRender_Click_Err

    Radio = Val(InputBox("Escriba la escala de 1 a 16 en la que generemos su mapa", "la escala se multiplica x 32"))

    If Radio = 0 Then Radio = 1
    If Radio >= 16 Then Radio = 16

    frmRenderer.picMap.Width = (Radio * 1000)
    frmRenderer.picMap.Height = (Radio * 1000)
    
    frmRenderer.Show , frmMain
    
    Exit Sub

mnuGuardarRender_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuGuardarRender_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGuardarUltimaConfig_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 23/05/06
'*************************************************
    mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInfoMap_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmMapInfo.Show
    frmMapInfo.Visible = True
End Sub

Private Sub mnuInformes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmInformes.Show
End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Superficie_Azar
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Superficie_Todo
End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmUnionAdyacente.Show
End Sub

Private Sub mnuMapa_Click()
    frmMapa.Show , frmMain
    Call DibujarMinimapa
End Sub

Private Sub mnuMapInfo_Click()
    frmMapInfo.Show , frmMain
End Sub

Private Sub mnuModoCaminata_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
    ToggleWalkMode
End Sub

Private Sub mnuNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    For i = 0 To 6
        If i <> 3 Then
            frmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    modPaneles.VerFuncion 3, True
End Sub

'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    Dim loopc As Integer
    
    DeseaGuardarMapa Dialog.filename
    
    For loopc = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(loopc).Visible = False
    Next
    
    frmMain.Dialog.filename = Empty
    
    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If
    
    Call modMapIO.NuevoMapa
    
    Call cmdInformacionDelMapa_Click

End Sub

Private Sub mnuObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    For i = 0 To 6
        If i <> 5 Then
            frmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    modPaneles.VerFuncion 5, True
End Sub


Private Sub mnuOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
    frmOptimizar.Show
End Sub

Private Sub mnuPegar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    Call PegarSeleccion
End Sub

Private Sub mnuQBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    modPaneles.VerFuncion 2, False
End Sub

Private Sub mnuQNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    modPaneles.VerFuncion 3, False
End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    modPaneles.VerFuncion 5, False
End Sub

Private Sub mnuQSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    modPaneles.VerFuncion 0, False
End Sub

Private Sub mnuQTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    modPaneles.VerFuncion 1, False
End Sub

Private Sub mnuQTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    modPaneles.VerFuncion 6, False
End Sub

Private Sub mnuQuitarBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Bloqueo_Todo(0)
End Sub

Private Sub mnuQuitarBloqueosBorde_Click()
    Call Resolucion
    Call Bloquear_Bordes(True)
End Sub

Private Sub mnuQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    ' Superficies
    cSeleccionarSuperficie.value = False
    Call cSeleccionarSuperficie_Click
    cQuitarEnEstaCapa.value = False
    Call cQuitarEnEstaCapa_Click
    cQuitarEnTodasLasCapas.value = False
    Call cQuitarEnTodasLasCapas_Click
    
    ' Translados
    cQuitarTrans.value = False
    Call cQuitarTrans_Click
    cInsertarTrans.value = False
    Call cInsertarTrans_Click
    
    ' Bloqueos
    cQuitarBloqueo.value = False
    Call cQuitarBloqueo_Click
    cInsertarBloqueo.value = False
    Call cInsertarBloqueo_Click
    
    ' Otras funciones
    cInsertarFunc(0).value = False
    Call cInsertarFunc_Click(0)
    cInsertarFunc(1).value = False
    Call cInsertarFunc_Click(1)
    cInsertarFunc(2).value = False
    Call cInsertarFunc_Click(2)
    cQuitarFunc(0).value = False
    Call cQuitarFunc_Click(0)
    cQuitarFunc(1).value = False
    Call cQuitarFunc_Click(1)
    cQuitarFunc(2).value = False
    Call cQuitarFunc_Click(2)
    
    ' Triggers
    cInsertarTrigger.value = False
    Call cInsertarTrigger_Click
    cQuitarTrigger.value = False
    Call cQuitarTrigger_Click
    
    'Luces
    frmLuces.AccionLuces
End Sub

Private Sub mnuQuitarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Quitar_NPCs(False)
End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Quitar_Objetos
End Sub

Private Sub mnuQuitarParticulas_Click()
    On Error GoTo mnuQuitarParticulas_Click_Err
    
    Dim X As Long
    Dim Y As Long
    Dim i As Long

    For X = 1 To XMaxMapSize
        For Y = 1 To YMaxMapSize
            MapData(X, Y).Particle_Index = 0
        Next Y
    Next X

    Particle_Group_Remove_All
    
    Exit Sub

mnuQuitarParticulas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmMain.mnuQuitarParticulas_Click", Erl)
    Resume Next
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Quitar_Bordes
End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Quitar_Capa(cCapas.Text)
End Sub

Private Sub mnuQuitarTODO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Borrar_Mapa
End Sub

Private Sub mnuQuitarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
    Call modEdicion.Quitar_Translados
End Sub

Private Sub mnuQuitarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Call modEdicion.Quitar_Triggers
End Sub

Private Sub mnuReAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error GoTo ErrHandler
    If FileExist(Dialog.filename, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.filename
        End If
    End If
    Call modMapIO.NuevoMapa
    
    If frmMain.Dialog.FilterIndex = 0 Then
        modMapIO.MapaV2_Cargar Dialog.filename
    Else
        modMapWinter.Cargar_CSM Dialog.filename
    End If
    
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
Exit Sub
ErrHandler:
End Sub

Private Sub mnuRealizarOperacion_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    Call AccionSeleccion
End Sub

Private Sub mnuSalir_Click()
'*************************************************
'Author: Lorwik
'Last modified: 20/03/2021
'*************************************************
    Call CloseClient
End Sub

Private Sub mnuSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    
    For i = 0 To 6
        If i <> 0 Then
            frmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    
    modPaneles.VerFuncion 0, True
End Sub

Private Sub mnuTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    
    For i = 0 To 6
        If i <> 1 Then
            frmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    
    modPaneles.VerFuncion 1, True
End Sub

Private Sub mnuTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    
    For i = 0 To 6
        If i <> 6 Then
            frmMain.SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    
    modPaneles.VerFuncion 6, True
End Sub

Private Sub mnuUtilizarDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
    mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)
End Sub


Private Sub mnuVerAutomatico_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)
End Sub

Private Sub mnuVerBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    cVerBloqueos.value = (cVerBloqueos.value = False)
    mnuVerBloqueos.Checked = cVerBloqueos.value

End Sub

Private Sub mnuVerCapa1_Click()
mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
End Sub

Private Sub mnuVerCapa2_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
End Sub

Private Sub mnuVerCapa3_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
End Sub

Private Sub mnuVerCapa4_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
End Sub


Private Sub mnuVerGrilla_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
    VerGrilla = (VerGrilla = False)
    mnuVerGrilla.Checked = VerGrilla
End Sub

Private Sub mnuVerParticulas_Click()
'*************************************************
'Author: Lorwik
'Last modified: 25/03/2021
'*************************************************
    VerParticulas = (VerParticulas = False)
    mnuVerParticulas.Checked = VerParticulas
End Sub

Private Sub mnuVerNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

End Sub

Private Sub mnuVerObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
    
End Sub

Private Sub mnuVerTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

End Sub

Private Sub mnuVerTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    cVerTriggers.value = (cVerTriggers.value = False)
    mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub mnuVerZonas_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 01/04/2021
'*************************************************

    mnuVerZonas(Index).Checked = True

    Select Case Index
    
        Case 0
            mnuVerZonas(1).Checked = False
            mnuVerZonas(2).Checked = False
        
        Case 1
            mnuVerZonas(0).Checked = False
            mnuVerZonas(2).Checked = False
        
        Case 2
            mnuVerZonas(0).Checked = False
            mnuVerZonas(1).Checked = False
    
    End Select

End Sub

Private Sub PicColorMap_Click()
    
    frmColorPicker.Show

End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    If X < 11 Then X = 11
    If X > 89 Then X = 89
    If Y < 10 Then Y = 10
    If Y > 92 Then Y = 92
    
    UserPos.X = X
    UserPos.Y = Y
    
    Call DibujarMinimapa(True)
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
    MiRadarX = X
    MiRadarY = Y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'Last modified: 20/11/07 - Loopzer
'*************************************************

    Dim tX As Integer
    Dim tY As Integer
    
    If Not MapaCargado Then Exit Sub
    
    'If X <= MainViewPic.Left Or X >= MainViewPic.Left + MainViewWidth Or Y <= MainViewPic.Top Or Y >= MainViewPic.Top + MainViewHeight Then
    '    Exit Sub
    'End If
    
    Call ConvertCPtoTP(X, Y, tX, tY)
    
    'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
        DX1.Text = tX
        DY1.Text = tY
    Else
        ClickEdit Button, tX, tY
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'*************************************************
    Dim tX As Integer
    Dim tY As Integer
    
    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True

    'Make sure click is in view window
    'If X <= MainViewPic.Left Or X >= MainViewPic.Left + MainViewWidth Or Y <= MainViewPic.Top Or Y >= MainViewPic.Top + MainViewHeight Then
    '    Exit Sub
    'End If
    
    Call ConvertCPtoTP(X, Y, tX, tY)
    
    MousePos = "X: " & tX & " - Y: " & tY
    
     If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
        DX2.Text = tX
        DY2.Text = tY
    Else
        ClickEdit Button, tX, tY
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************

    ' Guardar configuración
    WriteVar WEConfigDir, "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")
    If frmMain.mnuGuardarUltimaConfig.Checked = True Then
        WriteVar WEConfigDir, "PATH" & ClientSetup.WeMode, "UltimoMapa", Dialog.filename
        WriteVar WEConfigDir, "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Particulas", IIf(frmMain.mnuVerParticulas.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y
        WriteVar WEConfigDir, "CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
        WriteVar WEConfigDir, "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
        WriteVar WEConfigDir, "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
        WriteVar WEConfigDir, "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)
        
        WriteVar WEConfigDir, "MINIMAP", "Capa1", IIf(frmMain.Minimap_capa1.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Capa2", IIf(frmMain.Minimap_capa2.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Capa3", IIf(frmMain.Minimap_capa3.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Capa4", IIf(frmMain.Minimap_capa4.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Obj", IIf(frmMain.Minimap_objetos.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "NPC", IIf(frmMain.Minimap_npcs.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Particulas", IIf(frmMain.Minimap_particulas.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Bloqueos", IIf(frmMain.Minimap_bloqueos.Checked = True, "1", "0")
        WriteVar WEConfigDir, "MINIMAP", "Nombre", IIf(frmMain.Minimap_ndemapa.Checked = True, "1", "0")
    End If
    
    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If

End Sub

Private Sub SaveAllMiniMap_Click()
    MsgBox "No disponible"
End Sub

Private Sub SelectPanel_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim i As Byte
    
    For i = 0 To 9
        If i <> Index Then
            SelectPanel(i).value = False
            Call VerFuncion(i, False)
        End If
    Next
    
    If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
    Call VerFuncion(Index, SelectPanel(Index).value)
End Sub

Private Sub TimAutoGuardarMapa_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If mnuAutoGuardarMapas.Checked = True Then
        bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1
        If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
            If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
                modMapIO.GuardarMapa Dialog.filename
                
            End If
            bAutoGuardarMapaCount = 0
            
        End If
    End If
End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    With Dialog
        .Filter = "Mapas del nuevo formato (*.csm)|*.csm|Mapas clasicos de Argentum Online (*.map)|*.map"
        If Guardar Then
                .DialogTitle = "Guardar"
                .DefaultExt = ".txt"
                .filename = vbNullString
                .flags = cdlOFNPathMustExist
                .ShowSave
        Else
            .DialogTitle = "Cargar"
            .filename = vbNullString
            .flags = cdlOFNFileMustExist
            .ShowOpen
        End If
    End With
End Sub
