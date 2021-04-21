VERSION 5.00
Begin VB.Form frmMapInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del Mapa / Zona"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
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
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraInformacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox txtMapZona 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmMapInfo.frx":628A
         Left            =   1680
         List            =   "frmMapInfo.frx":6294
         TabIndex        =   38
         Text            =   "txtMapZona"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Frame FraTamanoDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tamaño del Screen"
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   5520
         Width           =   4455
         Begin VB.OptionButton OptTam 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tam. Winter"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptTam 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tam. Clasico"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame FraLuzBase 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Luz base"
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   4440
         Width           =   2175
         Begin VB.TextBox LuzMapa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   31
            Top             =   580
            Width           =   1335
         End
         Begin VB.PictureBox PicColorMap 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   375
         End
         Begin VB.CheckBox chkLuzClimatica 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Luz climatica"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            MaskColor       =   &H00404040&
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
         Begin WorldEditor.lvButtons_H LvBActualizarLuces 
            Height          =   375
            Left            =   1570
            TabIndex        =   32
            Top             =   120
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
            Image           =   "frmMapInfo.frx":62A8
            cBack           =   -2147483633
         End
      End
      Begin VB.Frame FraFormatoDel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tamaño del Mapa"
         Height          =   975
         Left            =   2400
         TabIndex        =   25
         Top             =   4440
         Width           =   2175
         Begin VB.OptionButton LvBOptX 
            BackColor       =   &H00FFFFFF&
            Caption         =   "100 x 100 (Clasico)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton LvBOptX 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1100 x 1100 (Winter)"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CheckBox chkMapMagiaSinEfecto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Magia Sin Efecto"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CheckBox chkMapBackup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Backup"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   3600
         Value           =   2  'Grayed
         Width           =   1575
      End
      Begin VB.TextBox txtMapNombre 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "Mapa Desconocido"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtMapMusica 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Text            =   "0"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox txtMapTerreno 
         Height          =   315
         ItemData        =   "frmMapInfo.frx":8326
         Left            =   1680
         List            =   "frmMapInfo.frx":8330
         TabIndex        =   12
         Text            =   "txtMapTerreno"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox chkMapPK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PK (inseguro)"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   1575
      End
      Begin VB.ComboBox txtMapRestringir 
         Height          =   315
         ItemData        =   "frmMapInfo.frx":8346
         Left            =   1680
         List            =   "frmMapInfo.frx":835C
         TabIndex        =   10
         Text            =   "txtMapRestringir"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtMapVersion 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkMapInviSinEfecto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "InviSinEfecto"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox chkMapResuSinEfecto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ResuSinEfecto"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CheckBox ChkMapNpc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Robo de NPC Permitido"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox TxtlvlMinimo 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "0"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox TxtAmbient 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "0"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.CheckBox chkInvocarSin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invocar sin efecto"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CheckBox chkOcultarSin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ocultar sin Efecto"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   8
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   4080
         Width           =   1935
      End
      Begin WorldEditor.lvButtons_H cmdMusica 
         Height          =   330
         Left            =   3600
         TabIndex        =   9
         Top             =   1050
         Width           =   735
         _extentx        =   1296
         _extenty        =   582
         caption         =   "&Más"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapInfo.frx":8392
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBGuardar 
         Height          =   375
         Left            =   2400
         TabIndex        =   33
         Top             =   6360
         Width           =   2055
         _extentx        =   2990
         _extenty        =   661
         caption         =   "&Guardar"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapInfo.frx":83BE
         mode            =   0
         value           =   0   'False
         cback           =   12632319
      End
      Begin WorldEditor.lvButtons_H cmdCerrar 
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   6360
         Width           =   1815
         _extentx        =   2990
         _extenty        =   661
         caption         =   "&Cerrar"
         capalign        =   2
         backstyle       =   2
         cgradient       =   0
         font            =   "frmMapInfo.frx":83EA
         mode            =   0
         value           =   0   'False
         cback           =   -2147483633
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   4315
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Mapa:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Musica:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terreno:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restringir:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión del Mapa:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Minimo:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido Ambiental:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmMapInfo"
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

Private Sub chkInvocarSin_LostFocus()
'*************************************************
'Author: Hardoz
'Last modified: 28/08/2010
'*************************************************
    MapInfo.InvocarSinEfecto = ChkMapNpc.value
    MapInfo.Changed = 1
 
End Sub

Private Sub chkLuzClimatica_Click()

    If chkLuzClimatica.value = Unchecked Then
        PicColorMap.BackColor = 0
        
        If ClientSetup.WeMode = eWeMode.WinterAO Then
            MapZonas(frmZonas.LstZona.ListIndex + 1).LuzBase = 0
            
        Else
            MapInfo.LuzBase = 0
            
        End If
        
        Call Actualizar_Estado
    End If
    
    frmMapInfo.chkLuzClimatica.value = chkLuzClimatica.value
    
End Sub

Private Sub chkMapBackup_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.BackUp = chkMapBackup.value
    MapInfo.Changed = 1
End Sub

Private Sub chkMapMagiaSinEfecto_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.MagiaSinEfecto = chkMapMagiaSinEfecto.value
    MapInfo.Changed = 1
    
End Sub

Private Sub chkMapInviSinEfecto_LostFocus()
'*************************************************
'Author:
'Last modified:
'*************************************************
    MapInfo.InviSinEfecto = chkMapInviSinEfecto.value
    MapInfo.Changed = 1

End Sub

Private Sub chkMapnpc_LostFocus()
'*************************************************
'Author: Hardoz
'Last modified: 28/08/2010
'*************************************************
    MapInfo.RoboNpcsPermitido = ChkMapNpc.value
    MapInfo.Changed = 1
 
End Sub

Private Sub chkMapResuSinEfecto_LostFocus()
'*************************************************
'Author:
'Last modified:
'*************************************************
    MapInfo.ResuSinEfecto = chkMapResuSinEfecto.value
    MapInfo.Changed = 1

End Sub

Private Sub chkMapPK_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.PK = chkMapPK.value
    MapInfo.Changed = 1
    
End Sub

Private Sub chkOcultarSin_Click()
'*************************************************
'Author: Lorwik
'Last modified: 26/04/2020
'*************************************************
    MapInfo.OcultarSinEfecto = ChkMapNpc.value
    MapInfo.Changed = 1
    
End Sub

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Me.Hide
    
End Sub

Private Sub cmdMusica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    frmMusica.Show
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
    
End Sub


Private Sub LvBActualizarLuces_Click()
    Call Actualizar_Estado
End Sub

Private Sub LvBGuardar_Click()
    Call guardarInfoZona(frmZonas.LstZona.ListIndex + 1)
End Sub

Public Sub guardarInfoZona(ByVal id As Integer)
    Dim i As Integer
    
    With MapZonas(id)
        .MapVersion = txtMapVersion.Text
        .name = txtMapNombre.Text
        .Music = txtMapMusica.Text
        .ambient = TxtAmbient.Text
        .PK = chkMapPK.value
        .MagiaSinEfecto = chkMapMagiaSinEfecto.value
        .InviSinEfecto = chkMapInviSinEfecto.value
        .ResuSinEfecto = chkMapResuSinEfecto.value
        .Terreno = txtMapTerreno.Text
        .Zona = txtMapZona.Text
        .Restringir = txtMapRestringir.Text
        .NoEncriptarMP = 0
        .LuzBase = .LuzBase
        
    End With
    
    Call ActualizarZonaList

End Sub

Public Sub LvBOptX_Click(Index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************
'Nota: Hay que cambiar muchas cosas, el engine cuando inicia hace calculos con el tamaï¿½o de los mapas
'ademas hay mas funciones que manejan estos datos, no basta con cambiar el XMax & YMax.
       
    ClientSetup.MapTam = Index
    
    Call WriteVar(WEConfigDir, "MOSTRAR", "MapTam", CStr(ClientSetup.MapTam))
       
    'Seteamos el nuevo tamaño del mapa
    Call setMapSize
End Sub

Public Sub OptTam_Click(Index As Integer)
    Call Resolucion
    
    Call WriteVar(WEConfigDir, "MOSTRAR", "Resolution", CStr(Index))
End Sub

Private Sub PicColorMap_Click()
    If chkLuzClimatica.value = False Then Exit Sub
    
    frmColorPicker.Show
End Sub

Private Sub txtAmbient_Change()
'*************************************************
'Author: Lorwik
'Last modified: 10/08/14
'*************************************************
    MapInfo.ambient = TxtAmbient.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapMusica_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.Music = txtMapMusica.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapVersion_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    MapInfo.MapVersion = txtMapVersion.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapNombre_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.name = txtMapNombre.Text
    MapInfo.Changed = 1
    Call AddtoRichTextBox(frmMain.StatTxt, "Nombre de mapa cambiado a:  " & MapInfo.name, 255, 255, 255, False, True, True)
    
End Sub

Private Sub txtlvlminimo_LostFocus()
'*************************************************
'Author: Lorwik
'Last modified: 13/09/11
'*************************************************
    MapInfo.lvlMinimo = TxtlvlMinimo.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapRestringir_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    KeyAscii = 0
    
End Sub

Private Sub txtMapRestringir_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.Restringir = txtMapRestringir.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapTerreno_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    KeyAscii = 0
    
End Sub

Private Sub txtMapTerreno_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.Terreno = txtMapTerreno.Text
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapZona_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    KeyAscii = 0
    
End Sub

Private Sub txtMapZona_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.Zona = txtMapZona.Text
    MapInfo.Changed = 1
    
End Sub

Public Sub CambiarColorMap()
On Error GoTo PicColorMap_Err
    
    If ClientSetup.WeMode = eWeMode.WinterAO Then
        PicColorMap.BackColor = MapZonas(frmZonas.LstZona.ListIndex + 1).LuzBase
    Else
        PicColorMap.BackColor = MapInfo.LuzBase
        
    End If
    
    frmMapInfo.PicColorMap.BackColor = PicColorMap.BackColor
    
    MapInfo.Changed = 1
    
    Exit Sub

PicColorMap_Err:
    Call RegistrarError(Err.Number, Err.Description, " FrmMain.Picture3_Click", Erl)
    Resume Next
End Sub
