VERSION 5.00
Begin VB.Form frmZonas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Zonas"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2310
   ControlBox      =   0   'False
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
   ScaleHeight     =   3390
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WorldEditor.lvButtons_H LvBCerrar 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Cerrar"
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
   Begin VB.Frame FraAreas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Zonas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.ListBox LstZona 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin WorldEditor.lvButtons_H LvBResetear 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Resetear"
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
      Begin WorldEditor.lvButtons_H LvBZona 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Nueva"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
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
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBZona 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Eliminar"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
      Begin WorldEditor.lvButtons_H LvBPintar 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
      Begin WorldEditor.lvButtons_H LvBQuitar 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   2280
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Quitar"
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
   End
End
Attribute VB_Name = "frmZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LstZona_Click()
 
    Call MapZona_Actualizar(LstZona.ListIndex + 1)
    
End Sub

Private Sub LvBCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub LvBPintar_Click()
    If LvBPintar.value = True Then
        LvBQuitar.Enabled = False
        
    Else
        LvBQuitar.Enabled = True
        
    End If
End Sub

Private Sub LvBQuitar_Click()
    If LvBQuitar.value = True Then
        LvBPintar.Enabled = False
        
    Else
        LvBPintar.Enabled = True
        
    End If
End Sub

Private Sub LvBResetear_Click()
    If MsgBox("¿¡Estas seguro que deseas resetear las propiedades de la zona!?", vbExclamation + vbYesNo) = vbYes Then
        Call ResetearZona(LstZona.ListIndex + 1)
        Call ActualizarZonaList
    End If
End Sub

Private Sub LvBZona_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    
        Case 0
            Call NuevaZona(Index + 1)
            
        Case 1
            If LstZona.ListIndex + 1 <> CantZonas Then
                MsgBox "Solo puedes eliminar la ultima zona de la lista. Si no vas a utilizar mas esa zona, reseteala para reutilizarla en el futuro."
                Exit Sub
            End If
            
            Call EliminarZona
            
    End Select
End Sub
