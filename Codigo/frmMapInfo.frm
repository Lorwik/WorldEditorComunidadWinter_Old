VERSION 5.00
Begin VB.Form frmMapInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informaci�n del Mapa"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   ControlBox      =   0   'False
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraLuzBase 
      Caption         =   "Luz base"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   4215
      Begin VB.CheckBox chkLuzClimatica 
         Caption         =   "Desactivado"
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
         Left            =   1080
         MaskColor       =   &H00404040&
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox PicColorMap 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox LuzMapa 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   27
         Text            =   "0-0-0"
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.CheckBox chkOcultarSin 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CheckBox chkInvocarSin 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox TxtAmbient 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   22
      Text            =   "0"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox TxtlvlMinimo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Text            =   "0"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CheckBox ChkMapNpc 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CheckBox chkMapResuSinEfecto 
      Caption         =   "ResuSinEfecto"
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox chkMapInviSinEfecto 
      Caption         =   "InviSinEfecto"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtMapVersion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Text            =   "0"
      Top             =   480
      Width           =   2655
   End
   Begin WorldEditor.lvButtons_H cmdMusica 
      Height          =   330
      Left            =   3600
      TabIndex        =   14
      Top             =   810
      Width           =   735
      _extentx        =   1296
      _extenty        =   582
      caption         =   "&M�s"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":628A
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   1695
      _extentx        =   2990
      _extenty        =   661
      caption         =   "&Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":62B6
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.ComboBox txtMapRestringir 
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
      ItemData        =   "frmMapInfo.frx":62E2
      Left            =   1680
      List            =   "frmMapInfo.frx":62FB
      TabIndex        =   11
      Text            =   "NO"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CheckBox chkMapPK 
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox txtMapTerreno 
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
      ItemData        =   "frmMapInfo.frx":6335
      Left            =   1680
      List            =   "frmMapInfo.frx":6342
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
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
      ItemData        =   "frmMapInfo.frx":635F
      Left            =   1680
      List            =   "frmMapInfo.frx":636C
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtMapMusica 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtMapNombre 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "Nuevo Mapa"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox chkMapBackup 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   3360
      Value           =   2  'Grayed
      Width           =   1575
   End
   Begin VB.CheckBox chkMapMagiaSinEfecto 
      Caption         =   "Magia Sin Efecto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin WorldEditor.lvButtons_H LvBGuardar 
      Height          =   375
      Left            =   2520
      TabIndex        =   30
      Top             =   5760
      Width           =   1695
      _extentx        =   2990
      _extenty        =   661
      caption         =   "&Guardar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":6388
      mode            =   0
      value           =   0   'False
      cback           =   12632319
   End
   Begin VB.Label Label8 
      Caption         =   "Sonido Ambiental:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Nivel Minimo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Versi�n del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   4300
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label5 
      Caption         =   "Restringir:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Terreno:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Zona:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Musica:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del Mapa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4315
      Y1              =   4200
      Y2              =   4200
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
    frmMain.chkLuzClimatica.value = chkLuzClimatica.value
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
    frmMain.chkPKInseguro.value = IIf(MapInfo.PK = True, 1, 0)
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

Private Sub LvBGuardar_Click()
    Call guardarInfoZona(frmMain.LstZona.ListIndex + 1)
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
    frmMain.txtMapMusica.Text = MapInfo.Music
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapVersion_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
    MapInfo.MapVersion = txtMapVersion.Text
    frmMain.txtMapVersion.Text = MapInfo.MapVersion
    MapInfo.Changed = 1
    
End Sub

Private Sub txtMapNombre_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    MapInfo.name = txtMapNombre.Text
    frmMain.txtMapNombre.Text = MapInfo.name
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
