VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conversor de Mapas"
   ClientHeight    =   8265
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   6270
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
   ScaleHeight     =   8265
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraConversorDe 
      Caption         =   "Conversor de Formatos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtMin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "Automatizar proceso"
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
         Left            =   3600
         TabIndex        =   31
         Top             =   1850
         Width           =   1815
      End
      Begin VB.TextBox txtMax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   30
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox ComOpracion 
         Height          =   315
         ItemData        =   "frmConvert.frx":0000
         Left            =   600
         List            =   "frmConvert.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1800
         Width           =   2295
      End
      Begin WorldEditor.lvButtons_H LvBConversion 
         Height          =   495
         Left            =   1920
         TabIndex        =   36
         Top             =   3360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Caption         =   "Conversión"
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
      Begin VB.Label Info 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Esperando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   2880
         Width           =   5535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   33
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblOperacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
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
         Left            =   1080
         TabIndex        =   29
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Integer, Long, CSM, ImpC, IAO1.3, IAO1.4"
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Top             =   960
         Width           =   3840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carpetas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrucciones:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Metes el mapa en su carpeta de origen, en la conversion, aparecera en su carpeta de destino."
         Height          =   435
         Left            =   1440
         TabIndex        =   24
         Top             =   360
         Width           =   4680
      End
   End
   Begin WorldEditor.lvButtons_H LvBCerrar 
      Height          =   495
      Left            =   1560
      TabIndex        =   22
      Top             =   7680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
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
   Begin VB.Frame FraTransformarA 
      Caption         =   "Transformar a mundo continuo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   6015
      Begin VB.Frame FraMapaPequeño 
         Caption         =   "Mapa Pequeño"
         Height          =   855
         Left            =   3360
         TabIndex        =   15
         Top             =   2280
         Width           =   2535
         Begin VB.TextBox txtminiY 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Text            =   "100"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtminiX 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   17
            Text            =   "100"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            Height          =   195
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   150
         End
         Begin VB.Label lblX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   150
         End
      End
      Begin VB.Frame FraMapaGigante 
         Caption         =   "Mapa Gigante"
         Height          =   1575
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtAncho 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   21
            Text            =   "1000"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   13
            Text            =   "1"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtTam2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            Text            =   "1100"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtTam1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   9
            Text            =   "1100"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblAncho 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho:"
            Height          =   195
            Left            =   600
            TabIndex        =   20
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblCsm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".csm"
            Height          =   195
            Left            =   1800
            TabIndex        =   14
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblNombreMapa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre: Mapa"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            Height          =   195
            Left            =   1320
            TabIndex        =   10
            Top             =   360
            Width           =   150
         End
         Begin VB.Label lblTamaño1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   150
         End
      End
      Begin VB.TextBox txtHasta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtDesde 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "110"
         Top             =   720
         Width           =   615
      End
      Begin WorldEditor.lvButtons_H LvBFusionarMapas 
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Fusionar mapas a Mundo Continuo"
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
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblSoloValido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solo valido para el mundo de 110 mapas de Winter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4275
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Automatico As Boolean

Private Sub chkAuto_Click()

    If chkAuto.value = False Then
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        txtMax.Visible = False
        Automatico = False
        
    Else
    
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        txtMax.Visible = True
        Automatico = True
        
    End If
    
End Sub

Private Sub Form_Load()
    ComOpracion.ListIndex = 0
End Sub

Private Sub LvBCerrar_Click()
    Unload Me
    
End Sub

Private Sub ConvertirInteger()

    Dim i As Integer
    
    Call frmMapInfo.LvBOptX_Click(0)
    
    If Automatico = False Then
        If FileExist(App.Path & "\Conversor\Integer\Mapa" & txtMin.Text & ".map", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call MapaV2_Cargar(App.Path & "\Conversor\Integer\Mapa" & txtMin.Text & ".map", True)
            Call MapaV2_Guardar(App.Path & "\Conversor\Long\Mapa" & txtMin.Text & ".map")
            
            Info.Caption = "Conversion realizada correctamente!"
                    
        Else
            Info.Caption = "Mapa" & txtMin.Text & ".map no existe!"
        End If
    Else
        For i = txtMin.Text To txtMax.Text
            If FileExist(App.Path & "\Conversor\Integer\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Integer\Mapa" & i & ".map", True)
                Call MapaV2_Guardar(App.Path & "\Conversor\Long\Mapa" & i & ".map")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            Else
                Info.Caption = "Mapa" & i & ".map no existe!"
                
            End If
            
        Next i
    End If
    
End Sub

Private Sub ConvertirLong()

    Dim i As Integer
    
    Call frmMapInfo.LvBOptX_Click(0)
    
    If Automatico = False Then
        If FileExist(App.Path & "\Conversor\Long\Mapa" & txtMin.Text & ".map", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call MapaV2_Cargar(App.Path & "\Conversor\Long\Mapa" & txtMin.Text & ".map")
            
            Call Save_CSM(App.Path & "\Conversor\CSM\Mapa" & txtMin.Text & ".csm")
            
            Info.Caption = "Conversion realizada correctamente!"
            
        Else
            Info.Caption = "Mapa" & txtMin.Text & ".map no existe!"
            
        End If
        
    Else
        For i = txtMin.Text To txtMax.Text
            
            If FileExist(App.Path & "\Conversor\Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Long\Mapa" & i & ".map")
                
                Call Save_CSM(App.Path & "\Conversor\CSM\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            Else
                Info.Caption = "Mapa" & i & ".map no existe!"
                
            End If
        Next i
    End If
    
End Sub

#If Privado = 0 Then
Private Sub ConvertirIAO()
    Dim i As Integer
    
    Call frmMapInfo.LvBOptX_Click(0)
    
    If Automatico = False Then
        If FileExist(App.Path & "\Conversor\IAO 1.3\Mapa" & txtMin.Text & ".csm", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call modMapImpC.Cargar_MapIAO(App.Path & "\Conversor\IAO 1.3\Mapa" & txtMin.Text & ".csm", tIAOold)
            
            Call modMapImpC.Save_MapIAO(App.Path & "\Conversor\IAO 1.4\Mapa" & txtMin.Text & ".csm")
            
            Info.Caption = "Conversion realizada correctamente!"
            
        Else
            Info.Caption = "Mapa" & txtMin.Text & " no existe!"
            
        End If
        
    Else
        For i = txtMin.Text To txtMax.Text
            
            If FileExist(App.Path & "\Conversor\IAO 1.3\Mapa" & i & ".csm", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call modMapImpC.Cargar_MapIAO(App.Path & "\Conversor\IAO 1.3\Mapa" & i & ".csm", tIAOold)
                
                Call modMapImpC.Save_MapIAO(App.Path & "\Conversor\IAO 1.4\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
            
            Else
                Info.Caption = "Mapa" & i & " no existe!"
            
            End If
        Next i
    End If
    
End Sub

Private Sub ConvertirImpC()

    Dim i As Integer
    
    Call frmMapInfo.LvBOptX_Click(0)
    
    If Automatico = False Then
        If FileExist(App.Path & "\Conversor\IAO 1.4\Mapa" & txtMin.Text & ".csm", vbNormal) = True Then
            Call modMapIO.NuevoMapa
            Call modMapImpC.Cargar_MapIAO(App.Path & "\Conversor\IAO 1.4\Mapa" & txtMin.Text & ".csm", tIAOnew)
            
            Call modMapImpC.Save_MapImpClasico(App.Path & "\Conversor\ImpC\Mapa" & txtMin.Text & ".csm")
            
            Info.Caption = "Conversion realizada correctamente!"
                    
        Else
            Info.Caption = "Mapa" & txtMin.Text & ".csm no existe!"
        End If
        
    Else
        For i = txtMin.Text To txtMax.Text
            
            If FileExist(App.Path & "\Conversor\IAO 1.4\Mapa" & i & ".csm", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call modMapImpC.Cargar_MapIAO(App.Path & "\Conversor\IAO 1.4\Mapa" & i & ".csm", tIAOnew)
                
                Call modMapImpC.Save_MapImpClasico(App.Path & "\Conversor\ImpC\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            End If
        Next i
    End If
    
End Sub

#End If

Private Sub LvBConversion_Click()

    Select Case ComOpracion.ListIndex
    
        Case 0 'Int > Long
            Call ConvertirInteger
            
        Case 1 'Long > CSM
            Call ConvertirLong
        
        #If Privado = 0 Then
        Case 2 'IAO 1.3 > IAO 1.4
            Call ConvertirIAO
            
        Case 3 'IAO 1.4 > ImpC
            Call ConvertirImpC
        #End If
    End Select
End Sub

Private Sub LvBFusionarMapas_Click()
    Dim i As Integer
    Dim tX As Integer
    Dim tY As Integer
    Dim sX As Integer
    Dim sY As Integer
    Dim Columnas As Integer
    Dim Fila As Integer
    
    Dim SuperMapX As Integer
    Dim SuperMapY As Integer
    
    Dim MiniMapX As Integer
    Dim MiniMapY As Integer
    
    SuperMapX = CInt(Val(txtTam1.Text))
    SuperMapY = CInt(Val(txtTam2.Text))
    
    MiniMapX = CInt(Val(txtminiX.Text))
    MiniMapY = CInt(Val(txtminiY.Text))
    
    ReDim SuperMapData(1 To SuperMapX, 1 To SuperMapY) As MapBlock
    
    Fila = 0
    Columnas = 0
    
    i = CInt(Val(txtDesde.Text))
    
    Do While i > Val(txtHasta.Text) - 1
        If i <> 0 Then
            Call modMapIO.NuevoMapa
            Call modMapWinter.Cargar_CSM_Old(App.Path & "\Conversor\Mapas\" & "Mapa" & i & ".csm")
            
            For tX = 1 To MiniMapX
                For tY = 1 To MiniMapY
                
                    sX = (Val(txtAncho.Text) - (Columnas * MiniMapX)) - 100
                    sY = (Fila * MiniMapY)
                
                    SuperMapData(sX + tX, sY + tY) = MapData(tX, tY)
    
                Next tY
            Next tX
            
            Columnas = Columnas + 1
            
            If Columnas = 10 Then
                Columnas = 0
                Fila = Fila + 1
            End If
            
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando mapa " & i, 255, 255, 255)
        End If
        
        If Val(txtDesde.Text) > Val(txtHasta.Text) Then
            i = i - Val(txtHasta.Text)
            
        Else
            i = i + Val(txtHasta.Text)
            
        End If
        
    Loop
    
    Call modMapWinter.Save_CSM(App.Path & "\Conversor\" & "Mapa" & Val(txtName.Text) & ".csm", True)
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa gigante guardado!", 1, 255, 1)
    
    'Destruyo el array para que no joda
    ReDim SuperMapData(1) As MapBlock
    
End Sub
