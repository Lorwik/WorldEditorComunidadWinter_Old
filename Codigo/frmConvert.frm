VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conversor de Mapas"
   ClientHeight    =   6795
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   6735
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
   ScaleHeight     =   6795
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin WorldEditor.lvButtons_H LvBCerrar 
      Height          =   375
      Left            =   1800
      TabIndex        =   34
      Top             =   6240
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      caption         =   "Cerrar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmConvert.frx":0000
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
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
      Left            =   360
      TabIndex        =   12
      Top             =   2880
      Width           =   6015
      Begin VB.Frame FraMapaPequeño 
         Caption         =   "Mapa Pequeño"
         Height          =   855
         Left            =   3360
         TabIndex        =   27
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
            TabIndex        =   31
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
            TabIndex        =   29
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
            TabIndex        =   30
            Top             =   360
            Width           =   150
         End
         Begin VB.Label lblX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   150
         End
      End
      Begin VB.Frame FraMapaGigante 
         Caption         =   "Mapa Gigante"
         Height          =   1575
         Left            =   3360
         TabIndex        =   19
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txt 
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
            TabIndex        =   33
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
            TabIndex        =   25
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
            TabIndex        =   23
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
            TabIndex        =   21
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
            TabIndex        =   32
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblCsm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".csm"
            Height          =   195
            Left            =   1800
            TabIndex        =   26
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblNombreMapa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre: Mapa"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            Height          =   195
            Left            =   1320
            TabIndex        =   22
            Top             =   360
            Width           =   150
         End
         Begin VB.Label lblTamaño1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
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
         TabIndex        =   16
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
         TabIndex        =   15
         Text            =   "110"
         Top             =   720
         Width           =   615
      End
      Begin WorldEditor.lvButtons_H LvBFusionarMapas 
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
         _extentx        =   4260
         _extenty        =   873
         caption         =   "Fusionar mapas a Mundo Continuo"
         capalign        =   2
         backstyle       =   2
         font            =   "frmConvert.frx":0028
         cgradient       =   0
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin VB.Label lblHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   1800
         TabIndex        =   18
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   195
         Left            =   360
         TabIndex        =   17
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
         TabIndex        =   14
         Top             =   240
         Width           =   4275
      End
   End
   Begin WorldEditor.lvButtons_H LvBConvertirLong 
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   2160
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      caption         =   "Convertir Long > CSM"
      capalign        =   2
      backstyle       =   2
      font            =   "frmConvert.frx":0054
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H LvBConvertirInteger 
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
      _extentx        =   4048
      _extenty        =   873
      caption         =   "Convertir Integer > Long"
      capalign        =   2
      backstyle       =   2
      font            =   "frmConvert.frx":0080
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
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
      Left            =   5040
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
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
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label7 
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
      Left            =   4320
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   $"frmConvert.frx":00AC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "Instrucciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Numero del mapa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
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

Private Sub LvBCerrar_Click()
    Unload Me
    
End Sub

Private Sub LvBConvertirInteger_Click()

    Dim i As Integer
    
    If Automatico = False Then
        Call modMapIO.NuevoMapa
        Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Integer\Mapa" & txtMin.Text & ".map", True)
        Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Long\Mapa" & txtMin.Text & ".map")
        
        Info.Caption = "Conversion realizada correctamente!"
        
    Else
        For i = txtMin.Text To txtMax.Text
            If FileExist(App.Path & "\Conversor\Mapas Integer\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Integer\Mapa" & i & ".map", True)
                Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            End If
            
        Next i
    End If
    
End Sub

Private Sub LvBConvertirLong_Click()

    Dim i As Integer
    
    If Automatico = False Then
        Call modMapIO.NuevoMapa
        Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & txtMin.Text & ".map")
        
        Call Save_CSM(App.Path & "\Conversor\Mapas CSM\Mapa" & txtMin.Text & ".csm")
        
        Info.Caption = "Conversion realizada correctamente!"
        
    Else
        For i = txtMin.Text To txtMax.Text
            
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                
                Call Save_CSM(App.Path & "\Conversor\Mapas CSM\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
                
            End If
        Next i
    End If
    
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
    
    Do While i > 0
        If i <> 0 Then
            Call modMapIO.NuevoMapa
            Call modMapWinter.Cargar_CSM(App.Path & "\Conversor\Mapas\" & "Mapa" & i & ".csm")
            
            For tX = 1 To MiniMapX
                For tY = 1 To MiniMapY
                
                    sX = 1000 - (Columnas * MiniMapX)
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
            i = i - txtHasta.Text
            
        Else
            i = i + txtHasta.Text
            
        End If
        
    Loop
    
    Call modMapWinter.Save_CSM(App.Path & "\Conversor\" & "Mapa" & Val(txtName.Text) & ".csm", True)
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa gigante guardado!", 1, 255, 1)
    
    'Destruyo el array para que no joda
    ReDim SuperMapData(1) As MapBlock
    
End Sub
