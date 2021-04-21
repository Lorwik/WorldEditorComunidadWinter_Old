Attribute VB_Name = "modCarga"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Enum eWeMode
    WinterAO
    ImperiumClasico
End Enum

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    LimiteFPS As Boolean
    OverrideVertexProcess As Byte
    
    'MOSTRAR
    MapTam As Byte
    Preview As Boolean
    
    'CONFIGURACION
    WeMode As Byte
End Type

Public ClientSetup As tSetupMods

Public grhCount As Long

Public Function WEConfigDir() As String
    WEConfigDir = App.Path & "\Datos\WorldEditor.ini"
End Function

''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal Path As String) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
    Path = Replace(Path, "/", "\")
    If Left(Path, 1) = "\" Then
        ' agrego app.path & path
        Path = App.Path & Path
    End If
    If Right(Path, 1) <> "\" Then
        ' me aseguro que el final sea con "\"
        Path = Path & "\"
    End If
    autoCompletaPath = Path
End Function

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Sub LeerConfiguracion()

On Local Error GoTo fileErr:
    
    Dim Lector As clsIniManager
    Dim i As Byte

    Set Lector = New clsIniManager
    Call Lector.Initialize(WEConfigDir)
    
    With ClientSetup
    
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        .LimiteFPS = CBool(Lector.GetValue("Video", "LimitarFPS"))
        .Preview = CBool(Lector.GetValue("MOSTRAR", "Preview"))
        .WeMode = Lector.GetValue("CONFIGURACION", "WeMode")
        
    End With

  Exit Sub
  
fileErr:

    If Err.Number <> 0 Then
       MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.Number & " : " & Err.Description)
       End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Public Sub CargarMapIni()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
    On Error GoTo Fallo
    Dim tStr As String
    Dim Leer As New clsIniReader
    Dim NewPath As String
    
    'Si el WorldEditor.ini no existe, tomamos estos parametros por defecto
    If Not FileExist(WEConfigDir, vbArchive) Then
        frmMain.mnuGuardarUltimaConfig.Checked = True
        MaxGrhs = 32000
        UserPos.X = 50
        UserPos.Y = 50
        PantallaX = 19
        PantallaY = 22
        MsgBox "Falta el archivo 'WorldEditor.ini' de configuración.", vbInformation
        Exit Sub
    End If
    
    Call Leer.Initialize(WEConfigDir)
    
    ClientSetup.MapTam = Val(Leer.GetValue("MOSTRAR", "MapTam"))
    
    ' Obj de Translado
    Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
    frmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
    frmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))
    frmMain.mnuUtilizarDeshacer.Checked = Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer"))
    
    ' Guardar Ultima Configuracion
    frmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))
    
    ' Index
    MaxGrhs = Val(GetVar(WEConfigDir, "INDEX", "MaxGrhs"))
    If MaxGrhs < 1 Then MaxGrhs = 32000
    
    'Reciente
    frmMain.Dialog.InitDir = Leer.GetValue("PATH" & ClientSetup.WeMode, "UltimoMapa")
    DirRecursos = autoCompletaPath(Leer.GetValue("PATH" & ClientSetup.WeMode, "DirRecursos"))
    
    '-------
    'Carga de rutas
    '-----------------------------------------

    '****
    'RUTA DE GRAFFICOS
    '*****************

    If FileExist(DirRecursos, vbDirectory) = False Or DirRecursos = "\" Then
        MsgBox "El directorio de Recursos es incorrecto", vbCritical + vbOKOnly
        
        NewPath = Buscar_Carpeta("DirRecursos", "")
        Call WriteVar(WEConfigDir, "PATH" & ClientSetup.WeMode, "DirRecursos", NewPath)
        DirRecursos = NewPath & "\"
    End If
    
    If FileExist(DirRecursos & "Graficos" & Formato, vbArchive) = False Then
        MsgBox "No se encontro el recursos de graficos."
        End
    End If
    
    If FileExist(DirRecursos & "Scripts" & Formato, vbArchive) = False Then
        MsgBox "No se encontro el recursos de Scripts."
        End
    End If
    
    '****
    'RUTA DE DATS
    '*****************
    DirDats = autoCompletaPath(Leer.GetValue("PATH" & ClientSetup.WeMode, "DirDats"))
    
    If FileExist(DirDats, vbDirectory) = False Or DirDats = "\" Then
        MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
        
        NewPath = Buscar_Carpeta("DirDats", "")
        Call WriteVar(WEConfigDir, "PATH" & ClientSetup.WeMode, "DirDats", NewPath)
        DirDats = NewPath & "\"
    End If
    
    If FileExist(DirDats & "Obj.dat", vbArchive) = False Then
        MsgBox "No se encontro el archivo Obj.dat."
        End
    End If
    
    If FileExist(DirDats & "NPcs.dat", vbArchive) = False Then
        MsgBox "No se encontro el archivo NPCs.dat."
        End
    End If
    
    '-----------------------------------------
    
    tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
    UserPos.X = Val(ReadField(1, tStr, Asc("-")))
    UserPos.Y = Val(ReadField(2, tStr, Asc("-")))
    
    If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
        UserPos.X = 50
    End If
    
    If UserPos.Y < YMinMapSize Or UserPos.Y > YMaxMapSize Then
        UserPos.Y = 50
    End If
    
    ' Menu Mostrar
    frmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))
    frmMain.mnuVerCapa2.Checked = Val(Leer.GetValue("MOSTRAR", "Capa2"))
    frmMain.mnuVerCapa3.Checked = Val(Leer.GetValue("MOSTRAR", "Capa3"))
    frmMain.mnuVerCapa4.Checked = Val(Leer.GetValue("MOSTRAR", "Capa4"))
    frmMain.mnuVerTranslados.Checked = Val(Leer.GetValue("MOSTRAR", "Translados"))
    frmMain.mnuVerObjetos.Checked = Val(Leer.GetValue("MOSTRAR", "Objetos"))
    frmMain.mnuVerNPCs.Checked = Val(Leer.GetValue("MOSTRAR", "NPCs"))
    frmMain.mnuVerTriggers.Checked = Val(Leer.GetValue("MOSTRAR", "Triggers"))
    frmMain.mnuVerGrilla.Checked = Val(Leer.GetValue("MOSTRAR", "Grilla")) ' Grilla
    VerGrilla = frmMain.mnuVerGrilla.Checked
    frmMain.mnuVerParticulas.Checked = Val(Leer.GetValue("MOSTRAR", "Particulas"))
    VerParticulas = frmMain.mnuVerParticulas.Checked
    frmMain.mnuVerBloqueos.Checked = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
    frmMain.cVerTriggers.value = frmMain.mnuVerTriggers.Checked
    frmMain.cVerBloqueos.value = frmMain.mnuVerBloqueos.Checked
    
    frmMain.Minimap_capa1.Checked = Val(Leer.GetValue("MINIMAP", "Capa1"))
    frmMain.Minimap_capa2.Checked = Val(Leer.GetValue("MINIMAP", "Capa2"))
    frmMain.Minimap_capa3.Checked = Val(Leer.GetValue("MINIMAP", "Capa3"))
    frmMain.Minimap_capa4.Checked = Val(Leer.GetValue("MINIMAP", "Capa4"))
    frmMain.Minimap_objetos.Checked = Val(Leer.GetValue("MINIMAP", "Obj"))
    frmMain.Minimap_npcs.Checked = Val(Leer.GetValue("MINIMAP", "NPC"))
    frmMain.Minimap_particulas.Checked = Val(Leer.GetValue("MINIMAP", "Particulas"))
    frmMain.Minimap_ndemapa.Checked = Val(Leer.GetValue("MINIMAP", "Nombre"))
    frmMain.Minimap_bloqueos.Checked = Val(Leer.GetValue("MINIMAP", "Bloqueos"))
    
    ' Tamaño de visualizacion
    PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
    PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))
    If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
    If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32
    
    ' [GS] 02/10/06
    ' Tamaño de visualizacion en el cliente
    ClienteHeight = Val(Leer.GetValue("MOSTRAR", "ClienteHeight"))
    ClienteWidth = Val(Leer.GetValue("MOSTRAR", "ClienteWidth"))
    If ClienteHeight <= 0 Then ClienteHeight = 13
    If ClienteWidth <= 0 Then ClienteWidth = 17
    
    Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en WorldEditor.ini" & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    Dim Leer As New clsIniReader
    Dim i As Integer

    If FileExist(IniPath & "Datos\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Datos\indices.ini'", vbCritical
        End
    End If
    
    Leer.Initialize IniPath & "Datos\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).name & " - #" & i
    Next
    
    DoEvents
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de Datos\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(IniPath & "Datos\Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en Datos\Triggers.ini", vbCritical
        End
    End If
    
    Dim NumT As Integer
    Dim T As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(IniPath & "Datos\Triggers.ini")
    
    frmMain.lListado(4).Clear
    
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For T = 1 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & (T - 1)
    Next T

    Set Leer = Nothing

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & App.Path & "\Datos\Triggers.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

Sub CargarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cuerpos
'*************************************
On Error GoTo ErrHandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(DirRecursos & "Scripts" & modCompression.Formato, LCase$("Personajes.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Personajes.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de cabezas
        NumCuerpos = fileBuff.getInteger()
    
        'Resize array
        ReDim BodyData(0 To NumCuerpos) As tBodyData
        ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
        
    
        For i = 1 To NumCuerpos
            MisCuerpos(i).Body(1) = fileBuff.getLong()
            MisCuerpos(i).Body(2) = fileBuff.getLong()
            MisCuerpos(i).Body(3) = fileBuff.getLong()
            MisCuerpos(i).Body(4) = fileBuff.getLong()
            MisCuerpos(i).HeadOffsetX = fileBuff.getInteger()
            MisCuerpos(i).HeadOffsetY = fileBuff.getInteger()
            
            If MisCuerpos(i).Body(1) Then
                Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
                Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
                Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
                Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
                
                BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
                BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. ")
            Call CloseClient
        End If
        
    End If
    
End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarCabezas()

End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
'On Error GoTo Fallo
Debug.Print DirDats & "NPCs.dat"

    If FileExist(DirDats & "NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If

    Dim Trabajando As String
    Dim NPC As Long
    Dim Leer As New clsIniReader
    Dim vDatos As String
    
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    
    Call Leer.Initialize(DirDats & "NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    
    ReDim NpcData(NumNPCs) As NpcData
    Trabajando = "Dats\NPCs.dat"
    
    'Call Leer.Initialize(DirDats & "\NPCs.dat")
    'MsgBox "  "
    For NPC = 1 To NumNPCs
        NpcData(NPC).name = CStr(Leer.GetValue("NPC" & NPC, "Name"))
        NpcData(NPC).ELV = Val(Leer.GetValue("NPC" & NPC, "ELV"))
        NpcData(NPC).Hostile = Val(Leer.GetValue("NPC" & NPC, "Hostile"))
        NpcData(NPC).NpcType = Val(Leer.GetValue("NPC" & NPC, "NPCType"))
        
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        
        vDatos = "#" & NPC & " " & NpcData(NPC).name
                
        'Marcamos a los hostiles y sus lvl
        If NpcData(NPC).Hostile = 1 Then
            vDatos = vDatos & " - [LVL:" & NpcData(NPC).ELV & " <HOSTIL>"
            
            'Marcamos a los WorldBoss
            If NpcData(NPC).NpcType = 12 Then _
                vDatos = vDatos & " <WORLDBOSS>"
                
            vDatos = vDatos & "]"
        End If
        
        If LenB(NpcData(NPC).name) <> 0 Then frmMain.lListado(1).AddItem vDatos
    Next
    
    Set Leer = Nothing
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Loads grh data using the new file format.
'

Public Sub LoadGrhData()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Graficos
'*************************************
On Error GoTo ErrorHandler:

    Dim Grh         As Long
    Dim Frame       As Long
    Dim fileVersion As Long
    Dim LaCabecera  As tCabecera
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    
    InfoHead = File_Find(DirRecursos & "Scripts" & Formato, LCase$("Graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Graficos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        fileVersion = fileBuff.getLong
        
        grhCount = fileBuff.getLong
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Grh < grhCount
            Grh = fileBuff.getLong

            With GrhData(Grh)
            
                '.active = True
                .NumFrames = fileBuff.getInteger
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = fileBuff.getLong
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    .speed = fileBuff.getSingle
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
                    .FileNum = fileBuff.getLong
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = fileBuff.getInteger
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = fileBuff.getInteger
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .sX = fileBuff.getInteger
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    .sY = fileBuff.getInteger
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
Exit Sub

ErrorHandler:
    
    If Err.Number <> 0 Then
        
        If Err.Number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe.")
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarMinimapa()

    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    Dim i           As Long
    
    InfoHead = File_Find(DirRecursos & "Scripts" & Formato, LCase$("minimap.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("minimap.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        For i = 1 To grhCount
            If Grh_Check(i) Then
                GrhData(i).mini_map_color = fileBuff.getLong
            End If
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grhCount Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function

