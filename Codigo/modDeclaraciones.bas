Attribute VB_Name = "modDeclaraciones"
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

''
' modDeclaraciones
'
' @remarks Declaraciones
' @author ^[GS]^
' @version 0.1.12
' @date 20081218

Option Explicit

Public MMiniMap_capa1      As Boolean
Public MMiniMap_capa2      As Boolean
Public MMiniMap_capa3      As Boolean
Public MMiniMap_capa4      As Boolean
Public MMiniMap_Npcs       As Boolean
Public MMiniMap_objetos    As Boolean
Public MMiniMap_Bloqueos   As Boolean
Public MMiniMap_particulas As Boolean
Public MMiniMap_Nombre     As Boolean

Public ParticlePreview As Long

Public NoSobreescribir As Boolean

Public Radio As Byte
Public ToWorldMap2 As Boolean

Public MapaActual As Integer

Public MousePos As String

Public Const MSGMod As String = "Este mapa há sido modificado." & vbCrLf & "Si no lo guardas perderas todos los cambios ¿Deseas guardarlo?"
Public Const MSGDang As String = "CUIDADO! Este comando puede arruinar el mapa." & vbCrLf & "¿Estas seguro que desea continuar?"

'[Loopzer]
Public SeleccionIX As Integer
Public SeleccionFX As Integer
Public SeleccionIY As Integer
Public SeleccionFY As Integer
Public SeleccionAncho As Integer
Public SeleccionAlto As Integer
Public Seleccionando As Boolean
Public SeleccionMap() As MapBlock

Public DeSeleccionOX As Integer
Public DeSeleccionOY As Integer
Public DeSeleccionIX As Integer
Public DeSeleccionFX As Integer
Public DeSeleccionIY As Integer
Public DeSeleccionFY As Integer
Public DeSeleccionAncho As Integer
Public DeSeleccionAlto As Integer
Public DeSeleccionando As Boolean
Public DeSeleccionMap() As MapBlock

Public VerBlockeados As Boolean
Public VerTriggers As Boolean
Public VerGrilla As Boolean ' grilla
Public VerParticulas As Boolean
Public VerCapa1 As Boolean
Public VerCapa2 As Boolean
Public VerCapa3 As Boolean
Public VerCapa4 As Boolean
Public VerTranslados As Boolean
Public VerObjetos As Boolean
Public VerNpcs As Boolean
'[/Loopzer]

' Objeto de Translado
Public Cfg_TrOBJ As Integer

'Path
Public IniPath As String
Public DirRecursos As String
Public DirDats As String

Public bAutoGuardarMapa As Byte
Public bAutoGuardarMapaCount As Byte
Public HotKeysAllow As Boolean  ' Control Automatico de HotKeys
Public vMostrando As Byte
Public WORK As Boolean
Public PATH_Save As String
Public NumMap_Save As Integer
Public NameMap_Save As String

' DX Config
Public PantallaX As Integer
Public PantallaY As Integer

' [GS] 02/10/06
' Client Config
Public ClienteHeight As Integer
Public ClienteWidth As Integer

Public SobreX As Integer ' Posicion X bajo el Cursor
Public SobreY As Integer   ' Posicion Y bajo el Cursor

' Radar
Public MiRadarX As Integer
Public MiRadarY As Integer

Type SupData
    name As String
    Grh As Long
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte
End Type
Public MaxSup As Long
Public SupData() As SupData

Public Type NpcData
    name As String
    ELV As Integer
    Hostile As Byte
    Body As Integer
    Head As Integer
    Heading As Byte
    NpcType As Byte
End Type

Public NumNPCs As Long
'Public NumNPCsHOST As Integer
Public NpcData() As NpcData

Public Type ObjData
    name As String 'Nombre del obj
    ObjType As Integer 'Tipo enum que determina cuales son las caract del obj
    GrhIndex As Long ' Indice del grafico que representa el obj
    GrhSecundario As Integer
    Info As String
    Ropaje As Integer 'Indice del grafico del ropaje
    WeaponAnim As Integer ' Apunta a una anim de armas
    ShieldAnim As Integer ' Apunta a una anim de escudo
    Texto As String
End Type
Public NumOBJs As Integer
Public ObjData() As ObjData

Public Conexion As New Connection
Public prgRun As Boolean
Public CurrentGrh As Grh
Public Play As Boolean
Public MapaCargado As Boolean
Public dLastWalk As Double

'Hold info about each map
Public Type tMapInfo
    Music As String
    name As String
    MapVersion As Integer
    PK As Boolean
    MagiaSinEfecto As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    LuzBase As Long
    Terreno As String
    Zona As String
    Restringir As String
    BackUp As Byte
    Changed As Byte ' flag for WorldEditor
    RoboNpcsPermitido As Byte
    InvocarSinEfecto As Byte
    OcultarSinEfecto As Byte
    lvlMinimo As Byte
    ambient As String
    NoEncriptarMP As Byte
End Type

Public MapZonas() As tMapInfo
Public CantZonas As Integer

'********** CONSTANTS ***********
'Heading Constants
Public Enum eDireccion
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'********** TYPES ***********
'Holds a local position
Public Type Position
    X As Integer
    Y As Integer
End Type

'Holds a world position
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Points to a grhData and keeps animation info
Public Type Grh
    GrhIndex As Long
    GrhIndexInt As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single

    mini_map_color As Long
End Type

' Cuerpos body.dat
Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type
' Lista de Cuerpos body.dat
Public Type tBodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

' body.dat
Public BodyData() As tBodyData
Public NumBodies As Integer

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Long
End Type

'Heads list
Public Type tHeadData
    Head(0 To 4) As Grh
End Type
Public HeadData() As tHeadData

'Hold info about a character
Public Type Char
    active As Byte
    Heading As Byte
    Pos As Position

    Body As tBodyData
    Head As tHeadData
    
    Moving As Byte
    MoveOffset As Position
    
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
End Type

'Holds info about a object
Public Type Obj
    ObjIndex As Integer
    Amount As Integer
End Type

Private Type tLight
    RGBCOLOR As D3DCOLORVALUE
    active As Boolean
    map_x As Integer
    map_y As Integer
    range As Byte
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    Engine_Light(0 To 3) As Long
    Light As tLight
    
    Particle_Index As Integer
    Particle_Group_Index As Long 'Particle Engine
    
    fX As Grh
    FxIndex As Integer
    
    ZonaIndex As Integer
End Type

Public Enum eTipoMapa
    tInt
    tLong
    tWinter
    tIAOClasico
    tIAOnew
    tIAOold
    tWinter_Old
End Enum

Public TipoMapaCargado As Byte

'********** Public VARS ***********

'Map sizes in tiles
Public XMaxMapSize As Integer
Public YMaxMapSize As Integer
Public Const XMinMapSize As Integer = 1
Public Const YMinMapSize As Integer = 1

'Where the map borders are.. Set during load
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

'********** Public ARRAYS ***********
Public GrhData() As GrhData 'Holds all the grh data
Public MapData() As MapBlock 'Holds map data for current map
Public SuperMapData() As MapBlock
Public MapInfo As tMapInfo 'Holds map info for current map
Public CharList(1 To 10000) As Char 'Holds info about all characters on map

'User status vars
Public CurMap As Integer 'Current map loaded
Public UserIndex As Integer
Global UserBody As Integer
Global UserHead As Integer
Public UserPos As Position 'Holds current user pos
Public AddtoUserPos As Position 'For moving user
Public UserCharIndex As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Map editor variables
Public WalkMode As Boolean

'Totals
Public NumMaps As Integer 'Number of maps
Public Numheads As Integer
Public NumGrhFiles As Integer 'Number of bmps
Public MaxGrhs As Long 'Number of Grhs
Global NumChars As Integer
Global LastChar As Integer

'********** Direct X ***********
Public MainViewRect As RECT
Public MainDestRect As RECT
Public MainViewWidth As Integer
Public MainViewHeight As Integer
Public BackBufferRect As RECT

'********** OUTSIDE FUNCTIONS ***********

'For Get and Write Var
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

'For KeyInput
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
