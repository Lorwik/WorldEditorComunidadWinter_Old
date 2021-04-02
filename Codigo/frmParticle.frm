VERSION 5.00
Begin VB.Form frmParticle 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Particulas"
   ClientHeight    =   3120
   ClientLeft      =   8280
   ClientTop       =   7845
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox ParticlePic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1800
      ScaleHeight     =   2265
      ScaleWidth      =   2145
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin WorldEditor.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Agregar"
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
   Begin VB.TextBox Life 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Text            =   "-1"
      Top             =   2400
      Width           =   390
   End
   Begin VB.ListBox lstParticle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin WorldEditor.lvButtons_H cmdDel 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Quitar"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LiveCounter:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frmParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If cmdAdd.value = True Then
        'lstParticle.Enabled = False
        cmdDel.Enabled = False
        Call modPaneles.EstSelectPanel(8, True)
        
    Else
        lstParticle.Enabled = True
        cmdDel.Enabled = True
        Call modPaneles.EstSelectPanel(8, False)
        
    End If
End Sub

Private Sub cmdDel_Click()
    If cmdDel.value = True Then
        lstParticle.Enabled = False
        cmdAdd.Enabled = False
        Call modPaneles.EstSelectPanel(8, True)
        
    Else
        lstParticle.Enabled = True
        cmdAdd.Enabled = True
        Call modPaneles.EstSelectPanel(8, False)
        
    End If
End Sub

Public Sub AccionParticulas()
    cmdAdd.value = False
    Call cmdAdd_Click
    cmdDel.value = False
    Call cmdDel_Click
End Sub

Private Sub lstParticle_Click()
    Dim Index As Integer
    
    Index = lstParticle.ListIndex + 1
    
    ParticlePreview = General_Particle_Create(Index, -1, -1)
    Debug.Print ParticlePreview
End Sub
