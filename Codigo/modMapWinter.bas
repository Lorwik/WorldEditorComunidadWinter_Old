Attribute VB_Name = "modMapWinter"
'********************************
'Load Map with .CSM format
'********************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Public Type tDatosLuces
    R As Integer
    G As Integer
    B As Integer
    range As Byte
    X As Integer
    Y As Integer
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    RoboNpcsPermitido As Boolean
    InvocarSinEfecto As Boolean
    OcultarSinEfecto As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    LuzBase As Long
    version As Long
    NoTirarItems As Boolean
End Type

Public MapSize As tMapSize
Private MapDat As tMapDat
'********************************
'END - Load Map with .CSM format
'********************************

Sub Cargar_CSM(ByVal Map As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 14/03/2021
    'Descripcion: Carga los mapas de WinterAO
    '***************************************************
    
    On Error GoTo ErrorHandler
    
    Dim fh As Integer
    Dim File As Integer
    Dim MH As tMapHeader
    Dim Blqs() As tDatosBloqueados
    Dim L1() As Long
    Dim L2() As tDatosGrh
    Dim L3() As tDatosGrh
    Dim L4() As tDatosGrh
    Dim Triggers() As tDatosTrigger
    Dim Luces() As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos() As tDatosObjs
    Dim NPCs() As tDatosNPC
    Dim TEs() As tDatosTE
    Dim LaCabecera As tCabecera
    
    Dim i As Long
    Dim j As Long
    DoEvents
        
    TipoMapaCargado = eTipoMapa.tWinter
        
    'Change mouse icon
    frmMain.MousePointer = 11
        
    fh = FreeFile
    Open Map For Binary Access Read As fh
    
        Get #fh, , LaCabecera
    
        Get #fh, , MH
        Get #fh, , MapSize
        Get #fh, , MapDat
        Debug.Print MapSize.XMax
        
        With MapSize
            If Not .XMax = XMaxMapSize Or Not .YMax = YMaxMapSize Then
                ReDim MapData(.XMin To .XMax, .YMin To .YMax)
            End If
            ReDim L1(.XMin To .XMax, .YMin To .YMax)
        End With
        
        Get #fh, , L1
        
        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(1 To .NumeroBloqueados)
                Get #fh, , Blqs
                For i = 1 To .NumeroBloqueados
                    MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
                Next i
            End If
            
            If .NumeroLayers(2) > 0 Then
                ReDim L2(1 To .NumeroLayers(2))
                Get #fh, , L2
                For i = 1 To .NumeroLayers(2)
                    InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(3) > 0 Then
                ReDim L3(1 To .NumeroLayers(3))
                Get #fh, , L3
                For i = 1 To .NumeroLayers(3)
                    InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex
                Next i
            End If
            
            If .NumeroLayers(4) > 0 Then
                ReDim L4(1 To .NumeroLayers(4))
                Get #fh, , L4
                For i = 1 To .NumeroLayers(4)
                    InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex
                Next i
            End If
            
            If .NumeroTriggers > 0 Then
                ReDim Triggers(1 To .NumeroTriggers)
                Get #fh, , Triggers
                For i = 1 To .NumeroTriggers
                    MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
                Next i
            End If
            
            If .NumeroParticulas > 0 Then
                ReDim Particulas(1 To .NumeroParticulas)
                Get #fh, , Particulas
                For i = 1 To .NumeroParticulas
                    'MapData(Particulas(i).X, Particulas(i).Y).Particle_Group_Index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
                Next i
            End If
                
            If .NumeroLuces > 0 Then
                ReDim Luces(1 To .NumeroLuces)
                Dim p As Byte
                Get #fh, , Luces
                For i = 1 To .NumeroLuces
                    Call Create_Light_To_Map(Luces(i).X, Luces(i).Y, Luces(i).range, Luces(i).R, Luces(i).G, Luces(i).B)
                Next i
                
                Call LightRenderAll
            End If
                
            If .NumeroOBJs > 0 Then
                ReDim Objetos(1 To .NumeroOBJs)
                Get #fh, , Objetos
                For i = 1 To .NumeroOBJs
                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex = Objetos(i).ObjIndex
                    MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount
                    If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex > NumOBJs Then
                        InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299
                    Else
                        InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.ObjIndex).GrhIndex
                    End If
                Next i
            End If
                
            If .NumeroNPCs > 0 Then
                ReDim NPCs(1 To .NumeroNPCs)
                Get #fh, , NPCs
                For i = 1 To .NumeroNPCs
                    If NPCs(i).NPCIndex > 0 Then
                        MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
                        Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).Y)
                    End If
                Next i
            End If
    
            If .NumeroTE > 0 Then
                ReDim TEs(1 To .NumeroTE)
                Get #fh, , TEs
                For i = 1 To .NumeroTE
                    MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                    MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                    MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
                Next i
            End If
            
        End With
    
    Close fh
    
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            If L1(i, j) > 0 Then
                InitGrh MapData(i, j).Graphic(1), L1(i, j)
            End If
        Next i
    Next j

    '*******************************
    'Render lights
    'Light_Render_All
    '*******************************
    
    'MapInfo_Cargar Map
    frmMain.txtMapVersion.Text = MapInfo.MapVersion
    
    Call Pestanas(Map, ".csm")

    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargar
    
    'Set changed flag
    MapInfo.Changed = 0

    MapaCargado = True
    
    'Call DibujarMiniMapa ' Radar
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)
ErrorHandler:
    If fh <> 0 Then Close fh
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    
    File = FreeFile
    Open App.Path & "\Logs.txt" For Output As #File
        Print #File, Err.Description
    Close #File
End Sub

Public Function Save_CSM(ByVal MapRoute As String) As Boolean

On Error GoTo ErrorHandler

    Dim fh As Integer
    Dim MH As tMapHeader
    Dim Blqs() As tDatosBloqueados
    Dim L1() As Long
    Dim L2() As tDatosGrh
    Dim L3() As tDatosGrh
    Dim L4() As tDatosGrh
    Dim Triggers() As tDatosTrigger
    Dim Luces() As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos() As tDatosObjs
    Dim NPCs() As tDatosNPC
    Dim TEs() As tDatosTE
    
    Dim i As Integer
    Dim j As Integer
    
    If FileExist(MapRoute, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & MapRoute & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Function
        Else
            Kill MapRoute
        End If
    End If
    
    frmMain.MousePointer = 11
    MapSize.XMax = XMaxMapSize
    MapSize.YMax = YMaxMapSize
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)
    
    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax
            With MapData(i, j)
                If .Blocked Then
                    MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                    ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                    Blqs(MH.NumeroBloqueados).X = i
                    Blqs(MH.NumeroBloqueados).Y = j
                End If
                
                L1(i, j) = .Graphic(1).GrhIndex
                
                If .Graphic(2).GrhIndex > 0 Then
                    MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                    ReDim Preserve L2(1 To MH.NumeroLayers(2))
                    L2(MH.NumeroLayers(2)).X = i
                    L2(MH.NumeroLayers(2)).Y = j
                    L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex
                End If
                
                If .Graphic(3).GrhIndex > 0 Then
                    MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                    ReDim Preserve L3(1 To MH.NumeroLayers(3))
                    L3(MH.NumeroLayers(3)).X = i
                    L3(MH.NumeroLayers(3)).Y = j
                    L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex
                End If
                
                If .Graphic(4).GrhIndex > 0 Then
                    MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                    ReDim Preserve L4(1 To MH.NumeroLayers(4))
                    L4(MH.NumeroLayers(4)).X = i
                    L4(MH.NumeroLayers(4)).Y = j
                    L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex
                End If
                
                If .Trigger > 0 Then
                    MH.NumeroTriggers = MH.NumeroTriggers + 1
                    ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                    Triggers(MH.NumeroTriggers).X = i
                    Triggers(MH.NumeroTriggers).Y = j
                    Triggers(MH.NumeroTriggers).Trigger = .Trigger
                End If
                
                If .Particle_Group_Index > 0 Then
                    MH.NumeroParticulas = MH.NumeroParticulas + 1
                    ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                    Particulas(MH.NumeroParticulas).X = i
                    Particulas(MH.NumeroParticulas).Y = j
                    'Particulas(MH.NumeroParticulas).Particula = CLng(particle_group_list(.Particle_Group_Index).stream_type)
                End If
               
               '¿Hay luz activa en este punto?
                'If .Light.Active Then
                '    MH.NumeroLuces = MH.NumeroLuces + 1
                '    ReDim Preserve Luces(1 To MH.NumeroLuces)
                    
                '    Luces(MH.NumeroLuces).R = .Light.RGBCOLOR.R
                '    Luces(MH.NumeroLuces).G = .Light.RGBCOLOR.G
                '    Luces(MH.NumeroLuces).B = .Light.RGBCOLOR.B
                '    Luces(MH.NumeroLuces).range = .Light.range
                '    Luces(MH.NumeroLuces).X = .Light.map_x
                '    Luces(MH.NumeroLuces).Y = .Light.map_y
                'End If
                
                If .OBJInfo.ObjIndex > 0 Then
                    MH.NumeroOBJs = MH.NumeroOBJs + 1
                    ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                    Objetos(MH.NumeroOBJs).ObjIndex = .OBJInfo.ObjIndex
                    Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                    Objetos(MH.NumeroOBJs).X = i
                    Objetos(MH.NumeroOBJs).Y = j
                End If
                
                If .NPCIndex > 0 Then
                    MH.NumeroNPCs = MH.NumeroNPCs + 1
                    ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                    NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                    NPCs(MH.NumeroNPCs).X = i
                    NPCs(MH.NumeroNPCs).Y = j
                End If
                
                If .TileExit.Map > 0 Then
                    MH.NumeroTE = MH.NumeroTE + 1
                    ReDim Preserve TEs(1 To MH.NumeroTE)
                    TEs(MH.NumeroTE).DestM = .TileExit.Map
                    TEs(MH.NumeroTE).DestX = .TileExit.X
                    TEs(MH.NumeroTE).DestY = .TileExit.Y
                    TEs(MH.NumeroTE).X = i
                    TEs(MH.NumeroTE).Y = j
                End If
            End With
        Next i
    Next j
    
    Call CSMInfoSave
              
    fh = FreeFile
    Open MapRoute For Binary As fh
        
        Put #fh, , MiCabecera
        
        Put #fh, , MH
        Put #fh, , MapSize
        Put #fh, , MapDat
        Put #fh, , L1
    
        With MH
            If .NumeroBloqueados > 0 Then _
                Put #fh, , Blqs
            If .NumeroLayers(2) > 0 Then _
                Put #fh, , L2
            If .NumeroLayers(3) > 0 Then _
                Put #fh, , L3
            If .NumeroLayers(4) > 0 Then _
                Put #fh, , L4
            If .NumeroTriggers > 0 Then _
                Put #fh, , Triggers
            If .NumeroParticulas > 0 Then _
                Put #fh, , Particulas
            If .NumeroLuces > 0 Then _
                Put #fh, , Luces
            If .NumeroOBJs > 0 Then _
                Put #fh, , Objetos
            If .NumeroNPCs > 0 Then _
                Put #fh, , NPCs
            If .NumeroTE > 0 Then _
                Put #fh, , TEs
        End With
    
    Close fh
    
    Call Pestanas(MapRoute, ".csm")
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    Save_CSM = True
    
     Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & MapRoute & " guardado...", 0, 255, 0)
    Exit Function

ErrorHandler:
    If fh <> 0 Then Close fh

End Function

Public Sub CSMInfoSave()
'**********************************
'Autor: Lorwik
'Fecha: 14/03/2021
'Descripcion: Guarda la informacion de los mapas de WinterAO.
'**********************************

    MapDat.map_name = MapInfo.name
    MapDat.music_number = MapInfo.Music
    
    MapDat.MagiaSinEfecto = MapInfo.MagiaSinEfecto
    MapDat.InviSinEfecto = MapInfo.InviSinEfecto
    MapDat.ResuSinEfecto = MapInfo.ResuSinEfecto
    MapDat.LuzBase = MapInfo.LuzBase
    MapDat.RoboNpcsPermitido = MapInfo.RoboNpcsPermitido
    MapDat.OcultarSinEfecto = MapInfo.OcultarSinEfecto
    MapInfo.InvocarSinEfecto = MapInfo.InvocarSinEfecto
    
    MapDat.lvlMinimo = MapInfo.lvlMinimo
    
    If frmMain.chkLuzClimatica = Checked Then
        MapDat.LuzBase = base_light
    Else
        MapDat.LuzBase = -1
    End If
    
    MapDat.version = MapInfo.MapVersion
    
    If MapInfo.PK = True Then
        MapDat.battle_mode = True
    Else
        MapDat.battle_mode = False
    End If
    
    MapDat.ambient = MapInfo.ambient
    MapDat.terrain = MapInfo.Terreno
    MapDat.zone = MapInfo.Zona
    MapDat.restrict_mode = MapInfo.Restringir
    MapDat.backup_mode = MapInfo.BackUp
    
End Sub

Public Sub CSMInfoCargar()
'**********************************
'Autor: Lorwik
'Fecha: 14/03/2021
'Descripcion: Cargar la informacion de los mapas de WinterAO.
'**********************************

    Dim tR As Byte
    Dim tG As Byte
    Dim tB As Byte
    
    MapInfo.name = MapDat.map_name
    MapInfo.Music = MapDat.music_number
    
    MapInfo.MagiaSinEfecto = MapDat.MagiaSinEfecto
    MapInfo.InviSinEfecto = MapDat.InviSinEfecto
    MapInfo.ResuSinEfecto = MapDat.ResuSinEfecto
    MapInfo.RoboNpcsPermitido = MapDat.RoboNpcsPermitido
    MapInfo.InvocarSinEfecto = MapInfo.InvocarSinEfecto
    MapInfo.OcultarSinEfecto = MapInfo.OcultarSinEfecto
    
    MapInfo.lvlMinimo = Val(MapDat.lvlMinimo)
    MapInfo.LuzBase = MapDat.LuzBase
    
    If MapDat.LuzBase <> -1 Then
        frmMain.chkLuzClimatica = Checked
        Call ConvertLongToRGB(MapDat.LuzBase, tR, tG, tB)
        
        frmMain.LuzMapa.Text = tR & "-" & tG & "-" & tB
    Else
        frmMain.chkLuzClimatica = Unchecked
    End If
    
    MapInfo.MapVersion = MapDat.version
    
    If MapDat.battle_mode = True Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.ambient = MapDat.ambient
    
    MapInfo.Terreno = MapDat.terrain
    MapInfo.Zona = MapDat.zone
    MapInfo.Restringir = MapDat.restrict_mode
    MapInfo.BackUp = MapDat.backup_mode
    
    Call MapInfo_Actualizar
End Sub

