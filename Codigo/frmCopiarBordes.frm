VERSION 5.00
Begin VB.Form frmCopiarBordes 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copiar Translados Automaticos"
   ClientHeight    =   6765
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
   ScaleHeight     =   451
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
      Top             =   3600
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton OptMapa 
         BackColor       =   &H00404040&
         Caption         =   "Perimetro"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   20
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton OptMapa 
         BackColor       =   &H00404040&
         Caption         =   "En Linea"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton OptMapa 
         BackColor       =   &H00404040&
         Caption         =   "Mapa Actual"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoCopy 
         BackColor       =   &H00404040&
         Caption         =   "Norte"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoCopy 
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
         Index           =   3
         Left            =   1800
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoCopy 
         BackColor       =   &H00404040&
         Caption         =   "Este"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoCopy 
         BackColor       =   &H00404040&
         Caption         =   "Oeste"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin WorldEditor.lvButtons_H LvBComenzar 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   2760
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
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   480
         X2              =   3840
         Y1              =   2280
         Y2              =   2280
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
         Left            =   1800
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
         Left            =   3240
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
         Left            =   1680
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
         Left            =   1560
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
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmCopiarBordes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EnCopia As Boolean

Private MapaNorte As Integer
Private MapaOeste As Integer
Private MapaEste As Integer
Private MapaSur As Integer

Private NuevoMapa As Integer
Private DireccionCopia As Byte

Public Sub Inicializar()

    On Error GoTo Inicializar_Err
    
    Dim X           As Integer
    Dim Y           As Integer
    Dim i           As Byte
    
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

Private Sub CopiarBorde(ByVal Direccion As Byte)
'****************************************
'Lorwik
'Fecha: 24/03/2021
'Descripción: Selecciona y copia el borde del mapa deseado en funciona a la resoluccion seleccionada
'****************************************

    Select Case Direccion
    
        Case eDireccion.NORTH
            SeleccionIX = XMinMapSize                                'Minimo
            SeleccionFX = XMaxMapSize                                'Maximo
            SeleccionIY = MinYBorder                                 'Minimo
            SeleccionFY = (MinYBorder) + (MinYBorder - YMinMapSize)  'Maximo
            
            DireccionCopia = eDireccion.SOUTH
            
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Norte del mapa " & MapaActual & ".", 0, 255, 0)
        
        Case eDireccion.WEST
            SeleccionIX = MinXBorder                                 'Minimo
            SeleccionFX = (MinXBorder) + (MinXBorder - XMinMapSize) 'Maximo
            SeleccionIY = YMinMapSize                                'Minimo
            SeleccionFY = YMaxMapSize                                'Maximo
            
            DireccionCopia = eDireccion.EAST
            
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Oeste del mapa " & MapaActual & ".", 0, 255, 0)
            
        Case eDireccion.EAST
            SeleccionIX = MinXBorder                                'Minimo
            SeleccionFX = (MinXBorder) - (MinXBorder - XMinMapSize) 'Maximo
            SeleccionIY = YMinMapSize                               'Minimo
            SeleccionFY = YMaxMapSize                               'Maximo
            
            DireccionCopia = eDireccion.WEST
            
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Este del mapa " & MapaActual & ".", 0, 255, 0)
            
        Case eDireccion.SOUTH
            SeleccionIX = XMinMapSize                               'Minimo
            SeleccionFX = XMaxMapSize                               'Maximo
            SeleccionIY = MaxYBorder                                'Minimo
            SeleccionFY = (MaxYBorder) - (MaxYBorder - YMaxMapSize) 'Maximo
            
            DireccionCopia = eDireccion.NORTH
            
            Call AddtoRichTextBox(frmMain.StatTxt, "Copiando el Sur del mapa " & MapaActual & ".", 0, 255, 0)
    
    End Select
    
    Call CopiarSeleccion(True)
    
    EnCopia = True

End Sub

Private Sub PegarBorde(Optional ByVal Direccion As Byte = 255)
'****************************************
'Lorwik
'Fecha: 24/03/2021
'Descripción: Pega el borde del mapa
'****************************************

    'Quizas queremos pegar en otra direccion
    If Direccion = 255 Then Direccion = DireccionCopia

    Select Case Direccion
    
        Case eDireccion.NORTH
            SobreX = 1
            SobreY = 1
            
        Case eDireccion.WEST
            SobreX = 1
            SobreY = 1
            
        Case eDireccion.EAST
            SobreX = MaxXBorder
            SobreY = 1
        
        Case eDireccion.SOUTH
            SobreX = 1
            SobreY = MaxYBorder
    
    End Select

    Call PegarSeleccion
    Call modEdicion.Bloquear_Bordes
    Call frmOptimizar.Optimizar
    
    EnCopia = False
End Sub

Private Sub VerMapaTraslado()
'****************************************
'Lorwik
'Fecha: 24/03/2021
'Descripción: Averigua el mapa al que dirige los traslados de las 4 direcciones
'****************************************

    On Error GoTo VerMapaTraslado_Err
    
    Dim X As Integer
    Dim Y As Integer

    'Izquierda
    X = MinXBorder

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map > 0 Then
            MapaOeste = MapData(X, Y).TileExit.Map
            If MapaOeste > 0 Then lblMapaOeste.Caption = MapaOeste
            Exit For

        End If

    Next
    
    'arriba
    Y = MinYBorder

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map > 0 Then
            MapaNorte = MapData(X, Y).TileExit.Map
            If MapaNorte > 0 Then lblMapaNorte.Caption = MapaNorte
            Exit For

        End If

    Next
    
    'Derecha
    X = MaxXBorder

    For Y = (MinYBorder + 1) To (MaxYBorder - 1)

        If MapData(X, Y).TileExit.Map > 0 Then
            MapaEste = MapData(X, Y).TileExit.Map
            If MapaEste > 0 Then lblMapaEste.Caption = MapaEste
            Exit For

        End If

    Next
    
    'Abajo
    Y = MaxYBorder

    For X = (MinXBorder + 1) To (MaxXBorder - 1)

        If MapData(X, Y).TileExit.Map > 0 Then
            MapaSur = MapData(X, Y).TileExit.Map
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
'****************************************
'Lorwik
'Fecha: 24/03/2021
'****************************************
    
    '¿No selecciono ninguna direccion?
    If OptMapa(2).value = False And chkAutoCopy(0).value = Unchecked And chkAutoCopy(1).value = Unchecked _
        And chkAutoCopy(2).value = Unchecked And chkAutoCopy(3).value = Unchecked Then
        
        MsgBox "No has seleccionado ninguna dirección."
        Exit Sub
        
    End If
    
    'Copia el mapa actual, pasa al siguiente y lo pega
    If OptMapa(0).value Then
    
        'Norte
        If chkAutoCopy(0).value = Checked Then
            If MapaNorte = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.NORTH)
            If GuardaryAbrir(MapaNorte, False) Then Call PegarBorde
            Exit Sub
        End If
        
        'Norte
        If chkAutoCopy(1).value = Checked Then
            If MapaOeste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.WEST)
            If GuardaryAbrir(MapaOeste, False) Then Call PegarBorde
            Exit Sub
        End If
        
        'Norte
        If chkAutoCopy(2).value = Checked Then
            If MapaEste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.EAST)
            If GuardaryAbrir(MapaEste, False) Then Call PegarBorde
            Exit Sub
        End If
        
        'Sur
        If chkAutoCopy(3).value = Checked Then
            If MapaSur = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.SOUTH)
            If GuardaryAbrir(MapaSur, False) Then Call PegarBorde
            Exit Sub
        End If
    
    ElseIf OptMapa(1).value Then 'Copia el mapa actual, pasa al siguiente, lo pega y asi hasta que no hay mas mapas siguientes.
    
        'Norte
        If chkAutoCopy(0).value = Checked Then
            Do While MapaNorte <> 0
                Call CopiarBorde(eDireccion.NORTH)
                If GuardaryAbrir(MapaNorte, False) Then Call PegarBorde
            Loop
            Exit Sub
        End If
        
        'Norte
        If chkAutoCopy(1).value = Checked Then
            Do While MapaOeste <> 0
                Call CopiarBorde(eDireccion.WEST)
                If GuardaryAbrir(MapaOeste, False) Then Call PegarBorde
            Loop
            Exit Sub
        End If
        
        'Norte
        If chkAutoCopy(2).value = Checked Then
            Do While MapaEste <> 0
                Call CopiarBorde(eDireccion.EAST)
                If GuardaryAbrir(MapaEste, False) Then Call PegarBorde
            Loop
            Exit Sub
        End If
        
        'Sur
        If chkAutoCopy(3).value = Checked Then
            Do While MapaSur <> 0
                Call CopiarBorde(eDireccion.SOUTH)
                If GuardaryAbrir(MapaSur, False) Then Call PegarBorde
            Loop
        End If
    
    ElseIf OptMapa(2).value Then 'Copia el mapa actual y todos los mapas que componen su perimetro
    
        'Norte
            If MapaNorte = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.NORTH)
            If GuardaryAbrir(MapaNorte, True) Then Call PegarBorde
            
        'Oeste
            If MapaOeste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.WEST)
            If GuardaryAbrir(MapaOeste, True) Then Call PegarBorde
        
        'Este
            If MapaEste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.EAST)
            If GuardaryAbrir(MapaEste, True) Then Call PegarBorde
        
        'Sur
            If MapaSur = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.SOUTH)
            If GuardaryAbrir(MapaSur, True) Then Call PegarBorde
        
        'Oeste
            If MapaOeste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.WEST)
            If GuardaryAbrir(MapaOeste, True) Then Call PegarBorde
            
        'Sur
            If MapaSur = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.SOUTH)
            If GuardaryAbrir(MapaSur, True) Then Call PegarBorde
            
        'Norte
            If MapaNorte = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.NORTH)
            If GuardaryAbrir(MapaNorte, True) Then Call PegarBorde
            
        'Este
            If MapaEste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.EAST)
            If GuardaryAbrir(MapaEste, True) Then Call PegarBorde
            
        'Este
            If MapaEste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.EAST)
            If GuardaryAbrir(MapaEste, True) Then Call PegarBorde
            
        'Norte
            If MapaNorte = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.NORTH)
            If GuardaryAbrir(MapaNorte, True) Then Call PegarBorde
            
        'Sur
            If MapaSur = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.SOUTH)
            If GuardaryAbrir(MapaSur, True) Then Call PegarBorde
            
        'Oeste
            If MapaOeste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.WEST)
            If GuardaryAbrir(MapaOeste, True) Then Call PegarBorde
            
        'Sur
            If MapaSur = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.SOUTH)
            If GuardaryAbrir(MapaSur, True) Then Call PegarBorde
            
        'Este
            If MapaEste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.EAST)
            If GuardaryAbrir(MapaEste, True) Then Call PegarBorde
            
        'Oeste
            If MapaOeste = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.WEST)
            If GuardaryAbrir(MapaOeste, True) Then Call PegarBorde
            
        'Morte
            If MapaNorte = 0 Then Exit Sub
            Call CopiarBorde(eDireccion.NORTH)
            If GuardaryAbrir(MapaNorte, True) Then Call PegarBorde

    
    End If

End Sub

Private Sub LvBCopiarAl_Click(Index As Integer)
'****************************************
'Lorwik
'Fecha: 24/03/2021
'****************************************
    Dim i As Byte

    Select Case Index
    
        Case 0
            Call CopiarBorde(eDireccion.NORTH)
            
        Case 1
            Call CopiarBorde(eDireccion.WEST)
            
        Case 2
            Call CopiarBorde(eDireccion.EAST)
            
        Case 3
            Call CopiarBorde(eDireccion.SOUTH)
    
    End Select
    
    For i = 0 To 3
        LvBCopiarAl(i).Enabled = False
    Next i
                
    LvBPegar.Enabled = True
    
End Sub

Private Sub LvBPegar_Click()
'****************************************
'Lorwik
'Fecha: 24/03/2021
'****************************************
    Dim i As Byte
    
    Call PegarBorde
    
    For i = 0 To 3
        LvBCopiarAl(i).Enabled = True
    Next i
                
    LvBPegar.Enabled = False
    
End Sub

Private Function GuardaryAbrir(ByVal newMap As Integer, ByVal Guardar As Boolean) As Boolean

    If Guardar Then
        NoSobreescribir = True
        frmMain.mnuGuardarMapa_Click
    End If

    '¿El mapa al que nos dirigimos existe?
    If FileExist(PATH_Save & NameMap_Save & newMap & ".csm", vbArchive) = True Then
        Call modMapIO.NuevoMapa
        frmMain.Dialog.filename = PATH_Save & NameMap_Save & newMap & ".csm"
        modMapIO.AbrirunMapa frmMain.Dialog.filename
        
        GuardaryAbrir = True
        
        Exit Function
    End If
End Function
