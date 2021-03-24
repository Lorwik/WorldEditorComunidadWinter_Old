VERSION 5.00
Begin VB.Form frmCopiarBordes 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Translados Automaticos"
   ClientHeight    =   7020
   ClientLeft      =   7185
   ClientTop       =   7845
   ClientWidth     =   4695
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
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraManual 
      BackColor       =   &H00404040&
      Caption         =   "Manual"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   4455
      Begin WorldEditor.lvButtons_H LvBPegar 
         Height          =   975
         Left            =   1740
         TabIndex        =   17
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Caption         =   "Pegar"
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
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBCopiarAl 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Copiar al Oeste"
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
      Begin WorldEditor.lvButtons_H LvBCopiarAl 
         Height          =   375
         Index           =   0
         Left            =   1500
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Copiar al Norte"
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
      Begin WorldEditor.lvButtons_H LvBCopiarAl 
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Copiar al Este"
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
      Begin WorldEditor.lvButtons_H LvBCopiarAl 
         Height          =   375
         Index           =   3
         Left            =   1500
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Copiar al Sur"
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
   End
   Begin VB.Frame FraCopiarBordes 
      BackColor       =   &H00404040&
      Caption         =   "Copiar bordes automatico"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Superior 
         BackColor       =   &H00404040&
         Caption         =   "Norte"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Inferior 
         BackColor       =   &H00404040&
         Caption         =   "Sur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Derecho 
         BackColor       =   &H00404040&
         Caption         =   "Este"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Izquierdo 
         BackColor       =   &H00404040&
         Caption         =   "Oeste"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin WorldEditor.lvButtons_H LvBComenzar 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Comenzar"
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
      Begin VB.Label lblMapaOeste 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Traslados"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   945
      End
      Begin VB.Label lblMapaNorte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Traslados"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblMapaEste 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Traslados"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   960
         Width           =   945
      End
      Begin VB.Label lblMapaSur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Traslados"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label lblMapActual 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1320
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblMapaActual 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa Actual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1485
         TabIndex        =   5
         Top             =   840
         Width           =   1065
      End
   End
   Begin WorldEditor.lvButtons_H LvBCopiarY 
      Height          =   375
      Left            =   720
      TabIndex        =   18
      Top             =   3120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Caption         =   "&Magic Mapas"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label lblCopiaY 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copia y optimiza todos los mapas alrededor del actual. No funcionan si no estan todos los mapas conectados."
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   4485
   End
End
Attribute VB_Name = "frmCopiarBordes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Izq As Integer
Private Arr As Integer
Private Der As Integer
Private Abj As Integer

Private Copiando As Byte
Private EnCopia As Boolean

Private MapaNorte As Integer
Private MapaOeste As Integer
Private MapaEste As Integer
Private MapaSur As Integer

Public Sub HacerTranslados()
    
    On Error GoTo HacerTranslados_Err
    
    lblMapActual.Caption = MapaActual

    Dim X As Integer
    Dim y As Integer

    Call VerMapaTraslado

    If Superior.value = vbChecked Then
        If MapaNorte = 0 Then
            MapaNorte = MapData(49, 10).TileExit.Map
            If MapaNorte > 0 Then lblMapaNorte.Caption = MapaNorte

            If MapaNorte = 0 Then
                Call SimpleLogError("Mapa " & lblMapActual.Caption & " sin translado")
                MsgBox "Arriba cancelado con dos intentos"
                Exit Sub

            End If

        End If

        ' ver ReyarB
        SeleccionIX = 1
        SeleccionFX = 100
        SeleccionIY = 10
        SeleccionFY = 21
        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaActual.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapaActual.Caption) & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If
    
        SobreX = 1
        SobreY = 90
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        NoSobreescribir = True
        frmMain.mnuGuardarMapa_Click

        'Call Form_Load
    
        If MapaSur = 0 Then
            MapaSur = MapData(49, 91).TileExit.Map
            If MapaSur > 0 Then lblMapaSur.Caption = MapaSur

            If MapaSur = 0 Then
                Call SimpleLogError("Mapa " & lblMapActual.Caption & " sin translado")
                Exit Sub

            End If

        End If
    
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If

    Call VerMapaTraslado

    If Inferior.value = vbChecked Then
        If MapaSur = 0 Then
            MapaSur = MapData(49, 91).TileExit.Map
            If MapaSur > 0 Then lblMapaSur.Caption = MapaSur

            If lblMapaSur = 0 Then
                Call SimpleLogError("Mapa " & lblMapActual.Caption & " sin translado")
                MsgBox "Abajo cancelado con dos intentos"
                Exit Sub

            End If

        End If

        ' ver ReyarB
        SeleccionIX = 1
        SeleccionFX = 100
        SeleccionIY = 81
        SeleccionFY = 89

        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapaActual.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapaActual.Caption) & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If
                
        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        NoSobreescribir = True
        frmMain.mnuGuardarMapa_Click
            
        Call Inicializar
                
        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If

    Call VerMapaTraslado

    If Derecho.value = vbChecked Then
        If MapaEste = 0 Then
            MapaEste = MapData(88, 49).TileExit.Map
            If MapaEste > 0 Then lblMapaEste.Caption = MapaEste
            
            If MapaEste = 0 Then
                Call SimpleLogError("Mapa " & lblMapActual.Caption & " sin translado")
                MsgBox "Derecha cancelado con dos intentos"
                Exit Sub

            End If

        End If

        ' ver ReyarB
        SeleccionIX = 76
        SeleccionFX = 87
        SeleccionIY = 1
        SeleccionFY = 100
        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & MapaEste & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & MapaEste & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If
                
        SobreX = 1
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        NoSobreescribir = True
        frmMain.mnuGuardarMapa_Click

        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If
            
    Call VerMapaTraslado

    If Izquierdo.value = vbChecked Then
        If MapaOeste = 0 Then
            MapaOeste = MapData(12, 49).TileExit.Map
            If MapaOeste > 0 Then lblMapaOeste.Caption = MapaOeste

            If MapaOeste = 0 Then
                Call SimpleLogError("Mapa " & lblMapActual.Caption & " sin translado")
                MsgBox "Izquierda cancelado con dos intentos"
                Exit Sub

            End If

        End If

        SeleccionIX = 13
        SeleccionFX = 25
        SeleccionIY = 1
        SeleccionFY = 100
        Call CopiarSeleccion

        If FileExist(PATH_Save & NameMap_Save & MapaOeste & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & MapaOeste & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If
                
        SobreX = 88
        SobreY = 1
        Call PegarSeleccion
        Call modEdicion.Bloquear_Bordes
        NoSobreescribir = True
        frmMain.mnuGuardarMapa_Click

        If FileExist(PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm"
            
            If ClientSetup.WeMode = eWeMode.WinterAO Then
                modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
            ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
            End If
                
            frmMain.mnuReAbrirMapa.Enabled = True

        End If

    End If
            
    Debug.Print "TERMINADO"
    Unload Me

    
    Exit Sub

HacerTranslados_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCopiarBordes.HacerTranslados", Erl)
    Resume Next
    
End Sub

Public Sub Inicializar()

    On Error GoTo Inicializar_Err
    
    Dim X           As Integer
    Dim y           As Integer
    Dim i           As Byte
    
    Izq = 9
    Arr = 7
    Der = 92
    Abj = 94
    
    If Not EnCopia Then
        For i = 0 To 3
            LvBCopiarAl(i).Enabled = True
        Next i
                
        LvBPegar.Enabled = False
    End If
    
    lblMapActual.Caption = MapaActual
    
    lblMapaNorte.Caption = "Sin Traslados"
    lblMapaSur.Caption = "Sin Traslados"
    lblMapaEste.Caption = "Sin Traslados"
    lblMapaOeste.Caption = "Sin Traslados"
    
    MapaNorte = 0
    MapaOeste = 0
    MapaEste = 0
    MapaSur = 0
    
    Call VerMapaTraslado
    
    Exit Sub

Inicializar_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCopiarBordes.Inicializar", Erl)
    Resume Next
    
End Sub

Private Sub VerMapaTraslado()
    
    On Error GoTo VerMapaTraslado_Err
    
    Dim X As Integer
    Dim y As Integer

    'Izquierda
    X = Izq

    For y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, y).TileExit.Map > 0 Then
            MapaOeste = MapData(X, y).TileExit.Map
            If MapaOeste > 0 Then lblMapaOeste.Caption = MapaOeste
            Exit For

        End If

    Next
    
    'arriba
    y = Arr

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, y).TileExit.Map > 0 Then
            MapaNorte = MapData(X, y).TileExit.Map
            If MapaNorte > 0 Then lblMapaNorte.Caption = MapaNorte
            Exit For

        End If

    Next
    
    'Derecha
    X = Der

    For y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, y).TileExit.Map > 0 Then
            MapaEste = MapData(X, y).TileExit.Map
            If MapaEste > 0 Then lblMapaEste.Caption = MapaEste
            Exit For

        End If

    Next
    
    'Abajo
    y = Abj

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, y).TileExit.Map > 0 Then
            MapaSur = MapData(X, y).TileExit.Map
            If MapaSur > 0 Then lblMapaSur.Caption = MapaSur
            Exit For

        End If

    Next

    
    Exit Sub

VerMapaTraslado_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.VerMapaTraslado", Erl)
    Resume Next
    
End Sub

Private Sub LvBComenzar_Click()
    'FrmMain.Timer4.Enabled = True
    
    On Error GoTo LvBComenzar_Click_Err
    
    If Superior.value = Unchecked And Inferior.value = Unchecked And _
        Izquierdo.value = Unchecked And Derecho.value = Unchecked Then
        
        MsgBox "No has seleccionado una dirección."
        Exit Sub
    End If
    
    frmMain.Dialog.filename = PATH_Save & NameMap_Save & CLng(lblMapActual.Caption) & ".csm"
    NoSobreescribir = True
    frmMain.mnuGuardarMapa_Click
    lblMapActual.Caption = MapaActual
    Call HacerTranslados
    lblMapActual.Caption = 0

    Exit Sub

LvBComenzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCopiarBordes.LvBComenzar_Click", Erl)
    Resume Next
End Sub

Private Sub LvBCopiarAl_Click(index As Integer)

    Call CopiarMapa(index)
    
End Sub

Private Sub LvBCopiarY_Click()

    Call frmOptimizar.Optimizar
    
    'Index de botones:
    '0: Norte
    '1: Oeste
    '2: Este
    '3: Sur
    
    'Norte
    If CopiarMapa(0) Then
        Copiando = eDireccion.NORTH
        Call LvBPegar_Click
        
    End If
    
    ' copio el de arriba al oeste
    If CopiarMapa(1) Then
        Copiando = eDireccion.WEST
        Call LvBPegar_Click
        
    End If
    
    ' vuelvo
    If CopiarMapa(2) Then
        Copiando = eDireccion.EAST
        Call LvBPegar_Click
        
    End If
    
    'copio al sur
    If CopiarMapa(3) Then
        Copiando = eDireccion.SOUTH
        Call LvBPegar_Click
        
    End If
    
    'Oeste
    If CopiarMapa(1) Then
        Copiando = eDireccion.WEST
        Call LvBPegar_Click
        
    End If
    
    'copio sur y vuelvo
    If CopiarMapa(3) Then
        Copiando = eDireccion.SOUTH
        Call LvBPegar_Click
        
        If CopiarMapa(0) Then
            Copiando = eDireccion.NORTH
            Call LvBPegar_Click
            
        End If
    End If
    
    If CopiarMapa(2) Then
        Copiando = eDireccion.EAST
        Call LvBPegar_Click
        
    End If
    
    'Este
    If CopiarMapa(2) Then
        Copiando = eDireccion.EAST
        Call LvBPegar_Click
        
    End If
    
    ' copio y vuelvo al norte
    If CopiarMapa(0) Then
        Copiando = eDireccion.NORTH
        Call LvBPegar_Click
        
        If CopiarMapa(3) Then
            Copiando = eDireccion.SOUTH
            Call LvBPegar_Click
            
        End If
    End If
    
    If CopiarMapa(1) Then
        Copiando = eDireccion.WEST
        Call LvBPegar_Click
    End If
    
    'Sur
    If CopiarMapa(3) Then
        Copiando = eDireccion.SOUTH
        Call LvBPegar_Click
    End If

    'copio este y vuelvo
    If CopiarMapa(2) Then
        Copiando = eDireccion.EAST
        Call LvBPegar_Click
        
        If CopiarMapa(1) Then
            Copiando = eDireccion.WEST
            Call LvBPegar_Click
            
        End If
    End If
    
    If CopiarMapa(0) Then
        Copiando = eDireccion.NORTH
        Call LvBPegar_Click
    End If

End Sub

Private Sub LvBPegar_Click()

    On Error GoTo LvBPegar_Click_Err
    
    Dim Nombre As Long
    
    If Copiando = eDireccion.NORTH Then
        Nombre = MapaNorte
        UserPos.y = 14
        
    ElseIf Copiando = eDireccion.WEST Then
        Nombre = MapaOeste
        UserPos.X = 19
        
    ElseIf Copiando = eDireccion.EAST Then
        Nombre = MapaEste
        UserPos.X = 83
        
    ElseIf Copiando = eDireccion.SOUTH Then
        Nombre = MapaSur
        UserPos.y = 87
    End If
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Pegando en el mapa " & Nombre & ".", 255, 204, 153, False, True)
    Debug.Print "Pegando en el mapa " & Nombre & "."
        
    If FileExist(PATH_Save & NameMap_Save & Nombre & ".csm", vbArchive) = True Then
        Call modMapIO.NuevoMapa
        frmMain.Dialog.filename = PATH_Save & NameMap_Save & Nombre & ".csm"
        
        If ClientSetup.WeMode = eWeMode.WinterAO Then
            modMapIO.AbrirunMapa (frmMain.Dialog.filename)
                
        ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
            modMapImpC.AbrirunMapaIAO frmMain.Dialog.filename, TipoMapaCargado
                
        End If
    
        frmMain.mnuReAbrirMapa.Enabled = True

    End If

    SobreX = 1
    SobreY = 1
    Call PegarSeleccion
    Call modEdicion.Bloquear_Bordes
    Call frmOptimizar.Optimizar
    MapInfo.Changed = 1
    
    EnCopia = False
    
    Exit Sub

LvBPegar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "Form2.Command6_Click", Erl)
    Resume Next
    
End Sub

Private Function CopiarMapa(ByVal index As Integer) As Boolean
    Dim i As Byte

    Call VerMapaTraslado

    Select Case index
    
        Case 0
            If MapaNorte = 0 Then
                CopiarMapa = False
                Exit Function
            End If
            
            SeleccionIX = 1
            SeleccionFX = 100
            SeleccionIY = 11
            SeleccionFY = 22
            Copiando = eDireccion.NORTH
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Norte del mapa " & lblMapActual.Caption & ".", 255, 204, 153, True, False)
            Debug.Print "Copiando el Norte del mapa " & lblMapActual.Caption & "."
            
        Case 1
            If MapaOeste = 0 Then
                CopiarMapa = False
                Exit Function
            End If
            
            SeleccionIX = 14
            SeleccionFX = 27
            SeleccionIY = 1
            SeleccionFY = 100
            Copiando = eDireccion.WEST
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Oeste del mapa " & lblMapActual.Caption & ".", 255, 204, 153, True, False)
            Debug.Print "Copiando el Oeste del mapa " & lblMapActual.Caption & "."
            
        Case 2
            If MapaEste = 0 Then
                CopiarMapa = False
                Exit Function
            End If
            
            SeleccionIX = 75
            SeleccionFX = 87
            SeleccionIY = 1
            SeleccionFY = 100
            Copiando = eDireccion.EAST
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Este del mapa " & lblMapActual.Caption & ".", 255, 204, 153, True, False)
            Debug.Print "Copiando el Este del mapa " & lblMapActual.Caption & "."
            
        Case 3
            If MapaSur = 0 Then
                CopiarMapa = False
                Exit Function
            End If
            
            SeleccionIX = 1
            SeleccionFX = 100
            SeleccionIY = 81
            SeleccionFY = 90
            Copiando = eDireccion.SOUTH
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Sur del mapa " & lblMapActual.Caption & ".", 255, 204, 153, True, False)
            Debug.Print "Copiando el Sur del mapa " & lblMapActual.Caption & "."
            
    End Select
    
    For i = 0 To 3
        LvBCopiarAl(i).Enabled = False
    Next i
            
    LvBPegar.Enabled = True
    EnCopia = True
    
    Call CopiarSeleccion
    MapInfo.Changed = 1
    NoSobreescribir = True
    frmMain.mnuGuardarMapa_Click
    
    CopiarMapa = True
    
End Function
