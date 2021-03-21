Attribute VB_Name = "modMapIO"
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
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit
Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByVal filename As String) As Long
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open filename For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
    If LenB(Dir(File, FileType)) = 0 Then
        FileExist = False
    Else
        FileExist = True
    End If

End Function

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

    frmMain.Dialog.CancelError = True
On Error GoTo ErrHandler
    
    If LenB(Path) = 0 Then
        frmMain.ObtenerNombreArchivo True
        Path = frmMain.Dialog.filename
        If LenB(Path) = 0 Then Exit Sub
    End If
    
    If frmMain.Dialog.FilterIndex = 1 Then
        Call Save_CSM(Path)
            
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        Call MapaV2_Guardar(Path)
        
    End If

ErrHandler:
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            GuardarMapa Path
        End If
    End If
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 21/05/06
'*************************************************

    On Error Resume Next

    Dim loopc As Integer
    Dim Y As Integer
    Dim X As Integer
    Dim i As Byte
    
    bAutoGuardarMapaCount = 0
    
    'frmMain.mnuUtirialNuevoFormato.Checked = True
    frmMain.mnuReAbrirMapa.Enabled = False
    frmMain.TimAutoGuardarMapa.Enabled = False
    frmMain.txtMapVersion.Text = 0
    
    MapaCargado = False
    
    For loopc = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(loopc).Enabled = False
    Next
    
    frmMain.MousePointer = 11
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, Y)
        
                ' Capa 1
                .Graphic(1).GrhIndex = 1
                
                ' Bloqueos
                .Blocked = 0
        
                ' Capas 2, 3 y 4
                .Graphic(2).GrhIndex = 0
                .Graphic(3).GrhIndex = 0
                .Graphic(4).GrhIndex = 0
        
                ' NPCs
                If .NPCIndex > 0 Then
                    EraseChar .CharIndex
                    .NPCIndex = 0
                End If
        
                ' OBJs
                .OBJInfo.ObjIndex = 0
                .OBJInfo.Amount = 0
                .ObjGrh.GrhIndex = 0
        
                ' Translados
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
                
                ' Triggers
                .Trigger = 0
        
                .Particle_Group_Index = 0
                'Particle_Group_Remove .Particle_Group_Index

                Call Engine_Long_To_RGB_List(MapData(X, Y).Engine_Light(), -1)
                
                Call mDx8_Luces.Delete_Light_To_Map(X, Y)
                
                .Light.active = False
                .Light.range = 0
                .Light.map_x = 0
                .Light.map_y = 0
                
                For i = 0 To 3
                    .Engine_Light(i) = 0
                Next i

                InitGrh .Graphic(1), 1
        
            End With
        Next X
    Next Y
    
    MapInfo.MapVersion = 0
    MapInfo.name = "Mapa Desconocido"
    MapInfo.Music = 0
    MapInfo.PK = True
    MapInfo.MagiaSinEfecto = 0
    MapInfo.InviSinEfecto = 0
    MapInfo.ResuSinEfecto = 0
    MapInfo.Terreno = "BOSQUE"
    MapInfo.Zona = "CAMPO"
    MapInfo.Restringir = "No"
    MapInfo.NoEncriptarMP = 0
    MapInfo.LuzBase = -1
    
    Call MapInfo_Actualizar
    
    bRefreshRadar = True ' Radar
    
    Estado_Actual = Estados(e_estados.MedioDia)
    Call Actualizar_Estado
    
    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0
    
    MapaCargado = True
    EngineRun = True
    
    frmMain.SetFocus

End Sub

''
' Guardar Mapa con el formato V2
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV2_Guardar(ByVal SaveAs As String, Optional ByVal Preguntar As Boolean = True)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc       As Long
    Dim TempInt     As Integer
    Dim Y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte

    If FileExist(SaveAs, vbNormal) = True Then
        
        If Preguntar Then
            If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            Else
                Call Kill(SaveAs)
            End If
        
        Else
            Call Kill(SaveAs)
            
        End If
        
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Call Kill(Left$(SaveAs, Len(SaveAs) - 4) & ".inf")
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.txtMapVersion.Text < 32767 Then
        frmMain.txtMapVersion.Text = frmMain.txtMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.txtMapVersion.Text
    End If

    Put FreeFileMap, , CInt(frmMain.txtMapVersion.Text)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, Y)
            
                ByFlags = 0
                
                If .Blocked = 1 Then ByFlags = ByFlags Or 1
                
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8

                If .Trigger Then ByFlags = ByFlags Or 16
                    
                Put FreeFileMap, , ByFlags
                    
                If TipoMapaCargado = eTipoMapa.tInt Then
                    Put FreeFileMap, , .Graphic(1).GrhIndexInt
                Else
                    Put FreeFileMap, , .Graphic(1).GrhIndex
                End If
                
                For loopc = 2 To 4
                    
                    If TipoMapaCargado = eTipoMapa.tInt Then
                        If .Graphic(loopc).GrhIndex Then Put FreeFileMap, , .Graphic(loopc).GrhIndexInt
                    Else
                        If .Graphic(loopc).GrhIndex Then Put FreeFileMap, , .Graphic(loopc).GrhIndex
                    End If

                Next loopc
                    
                If .Trigger Then Put FreeFileMap, , .Trigger
                
                'Escribimos el archivo ".INF"
                ByFlags = 0
                    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                If .NPCIndex Then ByFlags = ByFlags Or 2
                
                If .OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                    
                Put FreeFileInf, , ByFlags
                    
                If .TileExit.Map Then
                    Put FreeFileInf, , .TileExit.Map
                    Put FreeFileInf, , .TileExit.X
                    Put FreeFileInf, , .TileExit.Y
                End If
                    
                If .NPCIndex Then
                    Put FreeFileInf, , CInt(.NPCIndex)
                End If
                    
                If .OBJInfo.ObjIndex Then
                    Put FreeFileInf, , .OBJInfo.ObjIndex
                    Put FreeFileInf, , .OBJInfo.Amount
                End If
            
            End With
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestanas(SaveAs, ".map")

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    Call MapInfo_Guardar(SaveAs)

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description

End Sub


''
' Guardar Mapa con el formato V1
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV1_Guardar(SaveAs As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc As Long
    Dim TempInt As Integer
    Dim T As String
    Dim Y As Long
    Dim X As Long
    
    If FileExist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 11
    T = SaveAs
    If FileExist(Left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill Left(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    
    SaveAs = Left(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"
    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
    If frmMain.txtMapVersion.Text < 32767 Then
        frmMain.txtMapVersion.Text = frmMain.txtMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.txtMapVersion.Text
    End If
    Put FreeFileMap, , CInt(frmMain.txtMapVersion.Text)
    Put FreeFileMap, , MiCabecera
    
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.map file
            
            ' Bloqueos
            Put FreeFileMap, , MapData(X, Y).Blocked
            
            ' Capas
            For loopc = 1 To 4
                If loopc = 2 Then Call FixCoasts(MapData(X, Y).Graphic(loopc).GrhIndex, X, Y)
                Put FreeFileMap, , MapData(X, Y).Graphic(loopc).GrhIndex
            Next loopc
            
            ' Triggers
            Put FreeFileMap, , MapData(X, Y).Trigger
            Put FreeFileMap, , TempInt
            
            '.inf file
            'Tile exit
            Put FreeFileInf, , MapData(X, Y).TileExit.Map
            Put FreeFileInf, , MapData(X, Y).TileExit.X
            Put FreeFileInf, , MapData(X, Y).TileExit.Y
            
            'NPC
            Put FreeFileInf, , MapData(X, Y).NPCIndex
            
            'Object
            Put FreeFileInf, , MapData(X, Y).OBJInfo.ObjIndex
            Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount
            
            'Empty place holders for future expansion
            Put FreeFileInf, , TempInt
            Put FreeFileInf, , TempInt
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    'Close .inf file
    Close FreeFileInf
    FreeFileMap = FreeFile
    Open T & "2" For Binary Access Write As FreeFileMap
        Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pestanas(SaveAs)
    
    'write .dat file
    SaveAs = Left(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " guardado.", 0, 255, 0)
    
Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub

''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV2_Cargar(ByVal Map As String, Optional ByVal EsInteger As Boolean = False)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc       As Integer
    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim i           As Byte
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Con esto, le digo al WE que estamos usando mapas de tipo integer,
    'lo uso mas que nada para que no crashee cargar los mapas siguientes en las Pestañas.
    If EsInteger Then
        TipoMapaCargado = eTipoMapa.tInt
    Else
        TipoMapaCargado = eTipoMapa.tLong
    End If
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left$(Map, Len(Map) - 4)
    Map = Map & ".inf"
    
    FreeFileInf = FreeFile
    Open Map For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt

    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, Y)
            
                Get FreeFileMap, , ByFlags
                .Blocked = (ByFlags And 1)
            
                'Layer 1
                If EsInteger Then
                    Get FreeFileMap, , .Graphic(1).GrhIndexInt
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndexInt)
                Else
                    Get FreeFileMap, , .Graphic(1).GrhIndex
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                End If
            
                'Layer 2 used?
                If ByFlags And 2 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(2).GrhIndexInt
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndexInt)
                    Else
                        Get FreeFileMap, , .Graphic(2).GrhIndex
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndex)
                    End If
 
                Else
                
                    .Graphic(2).GrhIndex = 0
                    
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(3).GrhIndexInt
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexInt)
                    Else
                        Get FreeFileMap, , .Graphic(3).GrhIndex
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndex)
                    End If

                Else
                
                    .Graphic(3).GrhIndex = 0
                    
                End If
                
                'Layer 4 used?
                If ByFlags And 8 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(4).GrhIndexInt
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexInt)
                    Else
                        Get FreeFileMap, , .Graphic(4).GrhIndex
                        Call InitGrh(.Graphic(4), .Graphic(4).GrhIndex)
                    End If

                Else
                    
                    .Graphic(4).GrhIndex = 0

                End If
             
                'Trigger used?
                If ByFlags And 16 Then
                    Get FreeFileMap, , .Trigger
                Else
                    .Trigger = 0
                End If
            
                'Cargamos el archivo ".INF"
                Get FreeFileInf, , ByFlags
            
                If ByFlags And 1 Then
                    
                    With .TileExit
                    
                        Get FreeFileInf, , .Map
                        Get FreeFileInf, , .X
                        Get FreeFileInf, , .Y
                    
                    End With
                    

                End If
    
                If ByFlags And 2 Then
                
                    'Get and make NPC
                    Get FreeFileInf, , .NPCIndex
    
                    If .NPCIndex < 0 Then
                        .NPCIndex = 0
                    Else
                        Body = NpcData(.NPCIndex).Body
                        Head = NpcData(.NPCIndex).Head
                        Heading = NpcData(.NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
                    End If

                End If
    
                If ByFlags And 4 Then
                    
                    'Get and make Object
                    Get FreeFileInf, , .OBJInfo.ObjIndex
                    Get FreeFileInf, , .OBJInfo.Amount

                    If .OBJInfo.ObjIndex > 0 Then
                        Call InitGrh(.ObjGrh, ObjData(.OBJInfo.ObjIndex).GrhIndex)
                    End If

                End If
            
            End With
    
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pestanas(Map, ".map")
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    Call MapInfo_Cargar(Map)
    
    With frmMain
    
        frmMain.txtMapVersion.Text = MapInfo.MapVersion
    
        ' Avisamos que estamos trabajando con un mapa de tipo integer.
        If EsInteger Then
            .Caption = App.Title & " - Mapa Integer"
        Else
            .Caption = App.Title & " - Mapa Long"
        End If
        
        'Set changed flag
        MapInfo.Changed = 0
        
        'Change mouse icon
        .MousePointer = 0
    
    End With
    
    MapaCargado = True

End Sub

''
' Abrir Mapa con el formato V1
'
' @param Map Especifica el Path del mapa

Public Sub MapaV1_Cargar(ByVal Map As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error Resume Next
    Dim TBlock As Byte
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim T As String
    DoEvents
    'Change mouse icon
    frmMain.MousePointer = 11
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left(Map, Len(Map) - 4)
    Map = Map & ".inf"
    FreeFileInf = FreeFile
    Open Map For Binary As #2
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            '.map file
            Get FreeFileMap, , MapData(X, Y).Blocked
            
            For loopc = 1 To 4
                Get FreeFileMap, , MapData(X, Y).Graphic(loopc).GrhIndex
                'Set up GRH
                If MapData(X, Y).Graphic(loopc).GrhIndex > 0 Then
                    InitGrh MapData(X, Y).Graphic(loopc), MapData(X, Y).Graphic(loopc).GrhIndex
                End If
            Next loopc
            'Trigger
            Get FreeFileMap, , MapData(X, Y).Trigger
            
            Get FreeFileMap, , TempInt
            '.inf file
            
            'Tile exit
            Get FreeFileInf, , MapData(X, Y).TileExit.Map
            Get FreeFileInf, , MapData(X, Y).TileExit.X
            Get FreeFileInf, , MapData(X, Y).TileExit.Y
                          
            'make NPC
            Get FreeFileInf, , MapData(X, Y).NPCIndex
            If MapData(X, Y).NPCIndex > 0 Then
                Body = NpcData(MapData(X, Y).NPCIndex).Body
                Head = NpcData(MapData(X, Y).NPCIndex).Head
                Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
            End If
            
            'Make obj
            Get FreeFileInf, , MapData(X, Y).OBJInfo.ObjIndex
            Get FreeFileInf, , MapData(X, Y).OBJInfo.Amount
            If MapData(X, Y).OBJInfo.ObjIndex > 0 Then
                InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.ObjIndex).GrhIndex
            End If
            
            'Empty place holders for future expansion
            Get FreeFileInf, , TempInt
            Get FreeFileInf, , TempInt
                 
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
     
    Call Pestanas(Map)
    
    bRefreshRadar = True ' Radar
    
    Map = Left(Map, Len(Map) - 4) & ".dat"
        
    MapInfo_Cargar Map
    frmMain.txtMapVersion.Text = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub


Public Sub MapaV3_Cargar(ByVal Map As String)
'*************************************************
'Author: Loopzer
'Last modified: 22/11/07
'*************************************************

    On Error Resume Next
    Dim TBlock As Byte
    Dim loopc As Integer
    Dim TempInt As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    Dim Y As Integer
    Dim X As Integer
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim T As String
    DoEvents
    'Change mouse icon
    frmMain.MousePointer = 11
    
     FreeFileMap = FreeFile
    Open Map For Binary Access Read As FreeFileMap
        Get FreeFileMap, , MapData
    Close FreeFileMap
    
    Call Pestanas(Map)
    
    bRefreshRadar = True ' Radar
    
    Map = Left(Map, Len(Map) - 4) & ".dat"
        
    MapInfo_Cargar Map
    frmMain.txtMapVersion.Text = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True

End Sub
Public Sub MapaV3_Guardar(Mapa As String)
'*************************************************
'Author: Loopzer
'Last modified: 22/11/07
'*************************************************
'copy&paste RLZ
On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc As Long
    Dim TempInt As Integer
    Dim T As String
    Dim Y As Long
    Dim X As Long
    
    If FileExist(Mapa, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & Mapa & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill Mapa
        End If
    End If
    
    frmMain.MousePointer = 11
    
    FreeFileMap = FreeFile
    Open Mapa For Binary Access Write As FreeFileMap
        Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pestanas(Mapa)
    
    
    Mapa = Left(Mapa, Len(Mapa) - 4) & ".dat"
    MapInfo_Guardar Mapa
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " guardado.", 0, 255, 0)
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub




' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
    End If
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    Dim Leer As New clsIniReader
    Dim loopc As Integer
    Dim Path As String
    MapTitulo = Empty
    Leer.Initialize Archivo

    For loopc = Len(Archivo) To 1 Step -1
        If mid(Archivo, loopc, 1) = "\" Then
            Path = Left(Archivo, loopc)
            Exit For
        End If
    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    frmMapInfo.txtMapNombre.Text = MapInfo.name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
    frmMapInfo.txtMapRestringir.Text = MapInfo.Restringir
'    frmMapInfo.chkMapBackup.value = MapInfo.BackUp
    frmMapInfo.chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
    frmMain.chkPKInseguro.value = IIf(MapInfo.PK = True, 1, 0)
    frmMain.txtMapNombre.Text = MapInfo.name
    frmMain.txtMapMusica.Text = MapInfo.Music
    frmMain.TxtAmbient.Text = MapInfo.ambient
    frmMapInfo.TxtAmbient.Text = MapInfo.ambient
    frmMapInfo.TxtlvlMinimo = MapInfo.lvlMinimo
    frmMapInfo.chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
    frmMapInfo.chkMapInviSinEfecto.value = IIf(MapInfo.InviSinEfecto, vbChecked, vbUnchecked)
    frmMapInfo.chkInvocarSin.value = MapInfo.InvocarSinEfecto
    frmMapInfo.chkOcultarSin.value = MapInfo.OcultarSinEfecto
    frmMapInfo.chkMapResuSinEfecto.value = IIf(MapInfo.ResuSinEfecto, vbChecked, vbUnchecked)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion
    frmMapInfo.ChkMapNpc.value = MapInfo.RoboNpcsPermitido

End Sub

''
' Calcula la orden de Pestanas
'
' @param Map Especifica path del mapa

Public Sub Pestanas(ByVal Map As String, Optional ByVal MapFormat As String = ".map")
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
    Dim loopc As Integer
    
    For loopc = Len(Map) To 1 Step -1
        If mid(Map, loopc, 1) = "\" Then
            PATH_Save = Left(Map, loopc)
            Exit For
        End If
    Next
    
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))
    
    For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1
        If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
            NameMap_Save = Left(Map, loopc)
            Exit For
        End If
    Next
    
    For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)
            If FileExist(PATH_Save & NameMap_Save & loopc & MapFormat, vbArchive) = True Then
                frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
                frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
                frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
            Else
                frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False
            End If
    Next
    
End Sub

Public Sub AbrirMapa(Optional ByVal IntMode As Boolean = False)
    frmMain.Dialog.CancelError = True
    On Error GoTo ErrHandler
    
    DeseaGuardarMapa frmMain.Dialog.filename
    
    frmMain.ObtenerNombreArchivo False
    
    If Len(frmMain.Dialog.filename) < 3 Then Exit Sub
    
        If WalkMode = True Then
            Call modGeneral.ToggleWalkMode
        End If
        
        Call modMapIO.NuevoMapa
        
        If frmMain.Dialog.FilterIndex = 1 Then
            Call modMapWinter.Cargar_CSM(frmMain.Dialog.filename)
        Else
            Call MapaV2_Cargar(frmMain.Dialog.filename, IntMode)
            
        End If
        
        DoEvents
        frmMain.mnuReAbrirMapa.Enabled = True
        EngineRun = True
    
    Exit Sub
ErrHandler:
End Sub
