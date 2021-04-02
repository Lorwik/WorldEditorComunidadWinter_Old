Attribute VB_Name = "modEdicion"
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
' modEdicion
'
' @remarks Funciones de Edicion
' @author gshaxor@gmail.com
' @version 0.1.38
' @date 20061016

Option Explicit

' Deshacer
Public Const maxDeshacer As Integer = 10
Public MapData_Deshacer() As MapBlock

Type tDeshacerInfo

    Libre As Boolean
    Desc As String

End Type

Public MapData_Deshacer_Info(1 To maxDeshacer) As tDeshacerInfo

Public Sub InitDeshacer()
    '*************************************************
    'Author: Lorwik
    'Last modified: 22/03/2021
    '*************************************************
    
    'ReDim MapData_Deshacer(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
End Sub

''
' Vacia el Deshacer
'
Public Sub Deshacer_Clear()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Clear_Err
    
    Dim i As Integer

    ' Vacio todos los campos afectados
    For i = 1 To maxDeshacer
        MapData_Deshacer_Info(i).Libre = True
    Next
    ' no ahi que deshacer
    frmMain.mnuDeshacer.Enabled = False

    
    Exit Sub

Deshacer_Clear_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Deshacer_Clear", Erl)
    Resume Next
    
End Sub

''
' Agrega un Deshacer
'
Public Sub Deshacer_Add(ByVal Desc As String)
    
    On Error GoTo Deshacer_Add_Err
    
    
    Exit Sub

Deshacer_Add_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Deshacer_Add", Erl)
    Resume Next
    
End Sub

''
' Deshacer un paso del Deshacer
'
Public Sub Deshacer_Recover()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    
    On Error GoTo Deshacer_Recover_Err
    
    Dim i       As Integer
    Dim F       As Integer
    Dim j       As Integer
    Dim Body    As Integer
    Dim Head    As Integer
    Dim Heading As Byte

    If MapData_Deshacer_Info(1).Libre = False Then

        ' Aplico deshacer
        For F = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize

                If (MapData(F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex <> MapData_Deshacer(F, j).NPCIndex) Or (MapData(F, j).NPCIndex <> 0 And MapData_Deshacer(F, j).NPCIndex = 0) Then
                    ' Si ahi un NPC, y en el deshacer es otro lo borramos
                    ' (o) Si aun no NPC y en el deshacer no esta
                    MapData(F, j).NPCIndex = 0
                    Call EraseChar(MapData(F, j).CharIndex)

                End If

                If MapData_Deshacer(F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex = 0 Then
                    ' Si ahi un NPC en el deshacer y en el no esta lo hacemos
                    Body = NpcData(MapData_Deshacer(F, j).NPCIndex).Body
                    Head = NpcData(MapData_Deshacer(F, j).NPCIndex).Head
                    Heading = NpcData(MapData_Deshacer(F, j).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, F, j)
                Else
                    MapData(F, j) = MapData_Deshacer(F, j)

                End If

            Next
        Next
        MapData_Deshacer_Info(1).Libre = True

        ' Desplazo todos los deshacer uno hacia adelante
        For i = 1 To maxDeshacer - 1
            For F = XMinMapSize To XMaxMapSize
                For j = YMinMapSize To YMaxMapSize
                    MapData_Deshacer(F, j) = MapData_Deshacer(F, j)
                Next
            Next
            MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i + 1)
        Next
        ' borro el ultimo
        MapData_Deshacer_Info(maxDeshacer).Libre = True

        ' ahi para deshacer?
        If MapData_Deshacer_Info(1).Libre = True Then
            frmMain.mnuDeshacer.Caption = "&Deshacer (no ahi nada que deshacer)"
            frmMain.mnuDeshacer.Enabled = False
        Else
            frmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
            frmMain.mnuDeshacer.Enabled = True

        End If

    Else
        MsgBox "No ahi acciones para deshacer", vbInformation

    End If

    Exit Sub

Deshacer_Recover_Err:
    Call RegistrarError(Err.Number, Err.Description, "modEdicion.Deshacer_Recover", Erl)
    Resume Next
    
End Sub

''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio

Private Function EditWarning() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
    EditWarning = True
Else
    EditWarning = False
End If
End Function


''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes(Optional ByVal ac As Byte = 1)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Bloquear los bordes" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                MapData(X, Y).Blocked = ac
            End If
        Next X
    Next Y
    
    MapData(MinXBorder, MinYBorder).Blocked = ac
    MapData(MaxXBorder, MinYBorder).Blocked = ac
    MapData(MinXBorder, MaxYBorder).Blocked = ac
    MapData(MaxXBorder, MaxYBorder).Blocked = ac
    
    'Set changed flag
    MapInfo.Changed = 1
End Sub

''
' Coloca la superficie seleccionada al azar en el mapa
'

Public Sub Superficie_Azar()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    On Error Resume Next
    Dim Y As Integer
    Dim X As Integer
    Dim Cuantos As Integer
    Dim k As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    Cuantos = InputBox("Cuantos Grh se deben poner en este mapa?", "Poner Grh Al Azar", 0)
    
    If Cuantos > 0 Then
        modEdicion.Deshacer_Add "Insertar Superficie al Azar" ' Hago deshacer
        
        For k = 1 To Cuantos
            X = RandomNumber(10, 90)
            Y = RandomNumber(10, 90)
            
            If frmMain.MOSAICO.value = vbChecked Then
              Dim aux As Integer
              Dim dy As Integer
              Dim dX As Integer
              If frmMain.DespMosaic.value = vbChecked Then
                            dy = Val(frmMain.DMLargo)
                            dX = Val(frmMain.DMAncho.Text)
              Else
                        dy = 0
                        dX = 0
              End If
                    
              If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    aux = Val(frmMain.cGrh.Text) + _
                    (((Y + dy) Mod frmMain.mLargo.Text) * frmMain.mAncho.Text) + ((X + dX) Mod frmMain.mAncho.Text)
                    If frmMain.cInsertarBloqueo.value = True Then
                        MapData(X, Y).Blocked = 1
                    Else
                        MapData(X, Y).Blocked = 0
                    End If
                    MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                    InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), aux
              Else
                    Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                    tXX = X
                    tYY = Y
                    desptile = 0
                    For i = 1 To frmMain.mLargo.Text
                        For j = 1 To frmMain.mAncho.Text
                            aux = Val(frmMain.cGrh.Text) + desptile
                             
                            If frmMain.cInsertarBloqueo.value = True Then
                                MapData(tXX, tYY).Blocked = 1
                            Else
                                MapData(tXX, tYY).Blocked = 0
                            End If
    
                             MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                             
                             InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.Text)), aux
                             tXX = tXX + 1
                             desptile = desptile + 1
                        Next
                        tXX = X
                        tYY = tYY + 1
                    Next
                    tYY = Y
              End If
            End If
        Next
    End If
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Insertar Superficie en todos los bordes" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    
              If frmMain.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmMain.cGrh.Text) + _
                ((Y Mod frmMain.mLargo) * frmMain.mAncho) + (X Mod frmMain.mAncho)
                If frmMain.cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = 1
                Else
                    MapData(X, Y).Blocked = 0
                End If
                MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), aux
              Else
                'Else Place graphic
                If frmMain.cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = 1
                Else
                    MapData(X, Y).Blocked = 0
                End If
                
                MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = Val(frmMain.cGrh.Text)
                
                'Setup GRH
        
                InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), Val(frmMain.cGrh.Text)
            End If
                 'Erase NPCs
                If MapData(X, Y).NPCIndex > 0 Then
                    EraseChar MapData(X, Y).CharIndex
                    MapData(X, Y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(X, Y).OBJInfo.ObjIndex = 0
                MapData(X, Y).OBJInfo.Amount = 0
                MapData(X, Y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
    
            End If
    
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Insertar Superficie en todo el mapa" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If frmMain.MOSAICO.value = vbChecked Then
                Dim aux As Integer
                aux = Val(frmMain.cGrh.Text) + _
                ((Y Mod frmMain.mLargo) * frmMain.mAncho) + (X Mod frmMain.mAncho)
                 MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                'Setup GRH
                InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), aux
            Else
                'Else Place graphic
                MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = Val(frmMain.cGrh.Text)
                'Setup GRH
                InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), Val(frmMain.cGrh.Text)
            End If
    
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Superficie_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Superficie en area" ' Hago deshacer

    For Y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmMain.MOSAICO.value = vbChecked Then
                    Dim aux As Integer
                    aux = Val(frmMain.cGrh.Text) + _
                    ((Y Mod frmMain.mLargo) * frmMain.mAncho) + (X Mod frmMain.mAncho)
                     MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                    'Setup GRH
                    InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), aux
                Else
                    'Else Place graphic
                    MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = Val(frmMain.cGrh.Text)
                    'Setup GRH
                    InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)), Val(frmMain.cGrh.Text)
                End If
            Else
                MapData(X, Y).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = 0
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Bloqueos_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Inserta As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 07/12/2018
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Bloqueo en Area" ' Hago deshacer

    For Y = y1 To y2
        For X = x1 To x2
    
            If Inserta = True Then
                MapData(X, Y).Blocked = 1
            Else
                MapData(X, Y).Blocked = 0
            End If
    
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Triggers_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 25/03/2021
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Triggers en Area" ' Hago deshacer

    For Y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmMain.MOSAICO.value = vbChecked Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(X, Y).Trigger = frmMain.lListado(4).ListIndex
                Else
                    MapInfo.Changed = 1
                    'Else Place trigger
                    MapData(X, Y).Trigger = 0

                End If
            Else
                MapInfo.Changed = 1
                MapData(X, Y).Trigger = 0
                
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

Public Sub Zonas_Area(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer, ByVal Poner As Boolean)
'*************************************************
'Author: Lorwik
'Last modified: 01/04/2021
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Zonas en Area" ' Hago deshacer

    For Y = y1 To y2
        For X = x1 To x2
            If Poner = True Then
                If frmMain.MOSAICO.value = vbChecked Then
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(X, Y).ZonaIndex = frmMain.LstZona.ListIndex + 1
                    Debug.Print "Inserto"
                Else
                    MapInfo.Changed = 1
                    'Else Place Zona
                    MapData(X, Y).ZonaIndex = 0
                    Debug.Print "Quito"
                End If
            Else
                MapInfo.Changed = 1
                MapData(X, Y).ZonaIndex = 0
                Debug.Print "Quito 2"
                
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara


Public Sub Bloqueo_Todo(ByVal Valor As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Bloqueo todo" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, Y).Blocked = Valor
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Borra todo el Mapa menos los Triggers
'

Public Sub Borrar_Mapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Borrar todo el mapa menos Triggers" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            MapData(X, Y).Graphic(1).GrhIndex = 1
            'Change blockes status
            MapData(X, Y).Blocked = 0
    
            'Erase layer 2 and 3
            MapData(X, Y).Graphic(2).GrhIndex = 0
            MapData(X, Y).Graphic(3).GrhIndex = 0
            MapData(X, Y).Graphic(4).GrhIndex = 0
    
            'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If
    
            'Erase Objs
            MapData(X, Y).OBJInfo.ObjIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0
    
            'Clear exits
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
            
            InitGrh MapData(X, Y).Graphic(1), 1
    
        Next X
    Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    modEdicion.Deshacer_Add "Quitar todos los NPCs" & IIf(Hostiles = True, " Hostiles", "") ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).NPCIndex > 0 Then
                If (Hostiles = True And MapData(X, Y).NPCIndex >= 500) Or (Hostiles = False And MapData(X, Y).NPCIndex < 500) Then
                    Call EraseChar(MapData(X, Y).CharIndex)
                    MapData(X, Y).NPCIndex = 0
                End If
            End If
        Next X
    Next Y
    
    Call DibujarMinimapa(True)
    
    'Set changed flag
    MapInfo.Changed = 1
End Sub

''
' Elimita todos los Objetos del mapa
'

Public Sub Quitar_Objetos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    modEdicion.Deshacer_Add "Quitar todos los Objetos" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).OBJInfo.ObjIndex > 0 Then
                If MapData(X, Y).Graphic(3).GrhIndex = MapData(X, Y).ObjGrh.GrhIndex Then MapData(X, Y).Graphic(3).GrhIndex = 0
                MapData(X, Y).OBJInfo.ObjIndex = 0
                MapData(X, Y).OBJInfo.Amount = 0
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1
End Sub

''
' Elimina todos los Triggers del mapa
'

Public Sub Quitar_Triggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    Dim Y As Integer
    Dim X As Integer
    
    modEdicion.Deshacer_Add "Quitar todos los Triggers" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).Trigger > 0 Then
                MapData(X, Y).Trigger = 0
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1
    
End Sub

''
' Elimita todos los translados del mapa
'

Public Sub Quitar_Translados()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************

    If EditWarning Then Exit Sub

    Dim Y As Integer
    Dim X As Integer
    
    modEdicion.Deshacer_Add "Quitar todos los Translados" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).TileExit.Map > 0 Then
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Quitar todos los Bordes" ' Hago deshacer
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
            
                MapData(X, Y).Graphic(1).GrhIndex = 1
                InitGrh MapData(X, Y).Graphic(1), 1
                MapData(X, Y).Blocked = 0
                
                 'Erase NPCs
                If MapData(X, Y).NPCIndex > 0 Then
                    EraseChar MapData(X, Y).CharIndex
                    MapData(X, Y).NPCIndex = 0
                End If
    
                'Erase Objs
                MapData(X, Y).OBJInfo.ObjIndex = 0
                MapData(X, Y).OBJInfo.Amount = 0
                MapData(X, Y).ObjGrh.GrhIndex = 0
    
                'Clear exits
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.X = 0
                MapData(X, Y).TileExit.Y = 0
                
                ' Triggers
                MapData(X, Y).Trigger = 0
    
            End If
    
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1

End Sub

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa


Public Sub Quitar_Capa(ByVal Capa As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    If EditWarning Then Exit Sub
    
    '*****************************************************************
    'Clears one layer
    '*****************************************************************
    
    Dim Y As Integer
    Dim X As Integer
    
    If Not MapaCargado Then
        Exit Sub
    End If
    
    modEdicion.Deshacer_Add "Quitar Capa " & Capa ' Hago deshacer

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If Capa = 1 Then
                MapData(X, Y).Graphic(Capa).GrhIndex = 1
            Else
                MapData(X, Y).Graphic(Capa).GrhIndex = 0
            End If
        Next X
    Next Y
    
    'Set changed flag
    MapInfo.Changed = 1
End Sub

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tX As Integer, tY As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
    Dim Formato As String
    
    If frmMain.Dialog.FilterIndex = 1 Then
        Formato = ".csm"
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        Formato = ".map"
    End If

    ' Selecciones
    Seleccionando = False ' GS
    SeleccionIX = 0
    SeleccionIY = 0
    SeleccionFX = 0
    SeleccionFY = 0
    
    ' Translados
    Dim tTrans As WorldPos
    tTrans = MapData(tX, tY).TileExit
    If tTrans.Map > 0 Then
        If LenB(frmMain.Dialog.filename) <> 0 Then
            If FileExist(PATH_Save & NameMap_Save & tTrans.Map & Formato, vbArchive) = True Then
                Call modMapIO.NuevoMapa
                frmMain.Dialog.filename = PATH_Save & NameMap_Save & tTrans.Map & Formato
                
                Select Case frmMain.Dialog.FilterIndex
                
                    Case 1
                        If ClientSetup.WeMode = eWeMode.WinterAO Then
                            Call modMapWinter.Cargar_CSM(frmMain.Dialog.filename)
                            
                        ElseIf ClientSetup.WeMode = eWeMode.ImperiumClasico Then
                            If TipoMapaCargado = eTipoMapa.tIAOClasico Then
                                Call modMapImpC.Cargar_MapImpClasico(frmMain.Dialog.filename)
                            
                            ElseIf TipoMapaCargado = eTipoMapa.tIAOold Or TipoMapaCargado = eTipoMapa.tIAOnew Then
                                Call modMapImpC.Cargar_MapIAO(frmMain.Dialog.filename, TipoMapaCargado)
                                
                            End If
                        End If
                    Case 2
                
                End Select
                
                'modMapIO.AbrirunMapa frmMain.Dialog.filename
                UserPos.X = tTrans.X
                UserPos.Y = tTrans.Y
                
                If WalkMode = True Then
                    MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
                    CharList(UserCharIndex).Heading = SOUTH
                End If
                
                frmMain.mnuReAbrirMapa.Enabled = True
            End If
        End If
    End If
End Sub

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Sub ClickEdit(Button As Integer, tX As Integer, tY As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim loopc As Integer
    Dim NPCIndex As Integer
    Dim ObjIndex As Integer
    Dim Head As Integer
    Dim Body As Integer
    Dim Heading As Byte
    
    If tY < YMinMapSize Or tY > YMaxMapSize Then Exit Sub
    If tX < XMinMapSize Or tX > XMaxMapSize Then Exit Sub
    
    
    If Button = 0 Then
        ' Pasando sobre :P
        SobreY = tY
        SobreX = tX
        
    End If
    
    Select Case Button
    
        'Right
        Case vbRightButton
            
            ' Posicion
            Call AddtoRichTextBox(frmMain.StatTxt, "Posición X:" & tX & ", Y:" & tY, 255, 255, 255, False, False, True)
            
            ' Bloqueos
            If MapData(tX, tY).Blocked = 1 Then Call AddtoRichTextBox(frmMain.StatTxt, " (BLOQ)", 255, 255, 255, False, False, True)
            
            ' Translados
            If MapData(tX, tY).TileExit.Map > 0 Then
                If frmMain.mnuAutoCapturarTranslados.Checked = True Then
                    frmMain.tTMapa.Text = MapData(tX, tY).TileExit.Map
                    frmMain.tTX.Text = MapData(tX, tY).TileExit.X
                    frmMain.tTY = MapData(tX, tY).TileExit.Y
                End If
                Call AddtoRichTextBox(frmMain.StatTxt, " (Trans.: " & MapData(tX, tY).TileExit.Map & "," & MapData(tX, tY).TileExit.X & "," & MapData(tX, tY).TileExit.Y & ")", 255, 255, 255, False, False, True)
            End If
            
            ' NPCs
            If MapData(tX, tY).NPCIndex > 0 Then
                If MapData(tX, tY).NPCIndex > 499 Then
                    Call AddtoRichTextBox(frmMain.StatTxt, " (NPC-Hostil: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).name & ")", 255, 255, 255, False, False, True)
                Else
                    Call AddtoRichTextBox(frmMain.StatTxt, " (NPC: " & MapData(tX, tY).NPCIndex & " - " & NpcData(MapData(tX, tY).NPCIndex).name & ")", 255, 255, 255, False, False, True)
                End If
            End If
            
            ' OBJs
            If MapData(tX, tY).OBJInfo.ObjIndex > 0 Then
                Call AddtoRichTextBox(frmMain.StatTxt, " (Obj: " & MapData(tX, tY).OBJInfo.ObjIndex & " - " & ObjData(MapData(tX, tY).OBJInfo.ObjIndex).name & " - Cant.:" & MapData(tX, tY).OBJInfo.Amount & ")", 255, 255, 255, False, False, True)
            End If
            
            ' Capas
            Call AddtoRichTextBox(frmMain.StatTxt, "Capa1: " & MapData(tX, tY).Graphic(1).GrhIndex & " - Capa2: " & MapData(tX, tY).Graphic(2).GrhIndex & " - Capa3: " & MapData(tX, tY).Graphic(3).GrhIndex & " - Capa4: " & MapData(tX, tY).Graphic(4).GrhIndex, 255, 255, 255, False, False, True)
            If frmMain.mnuAutoCapturarSuperficie.Checked = True And frmMain.cSeleccionarSuperficie.value = False Then
                If MapData(tX, tY).Graphic(4).GrhIndex <> 0 Then
                    frmMain.cCapas.Text = 4
                    frmMain.cGrh.Text = MapData(tX, tY).Graphic(4).GrhIndex
                    
                ElseIf MapData(tX, tY).Graphic(3).GrhIndex <> 0 Then
                    frmMain.cCapas.Text = 3
                    frmMain.cGrh.Text = MapData(tX, tY).Graphic(3).GrhIndex
                    
                ElseIf MapData(tX, tY).Graphic(2).GrhIndex <> 0 Then
                    frmMain.cCapas.Text = 2
                    frmMain.cGrh.Text = MapData(tX, tY).Graphic(2).GrhIndex
                    
                ElseIf MapData(tX, tY).Graphic(1).GrhIndex <> 0 Then
                    frmMain.cCapas.Text = 1
                    frmMain.cGrh.Text = MapData(tX, tY).Graphic(1).GrhIndex
                    
                End If
            End If
            
            Exit Sub
        
        'Left click
        Case vbLeftButton
                
            'Erase 2-3
            If frmMain.cQuitarEnTodasLasCapas.value = True Then
                MapInfo.Changed = 1 'Set changed flag
                For loopc = 2 To 3
                    MapData(tX, tY).Graphic(loopc).GrhIndex = 0
                Next loopc
                    
                Exit Sub
            End If
        
            'Borrar "esta" Capa
            If frmMain.cQuitarEnEstaCapa.value = True Then
                If Val(frmMain.cCapas.Text) = 1 Then
                    If MapData(tX, tY).Graphic(1).GrhIndex <> 1 Then
                        modEdicion.Deshacer_Add "Quitar Todas las Capas (2/3)" ' Hago deshacer
                        MapInfo.Changed = 1 'Set changed flag
                        MapData(tX, tY).Graphic(1).GrhIndex = 1
                        Exit Sub
                    End If
                    
                ElseIf MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex <> 0 Then
                    modEdicion.Deshacer_Add "Quitar Capa " & frmMain.cCapas.Text  ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = 0
                    Exit Sub
                End If
            End If
        
            '************** Place grh
            If frmMain.cSeleccionarSuperficie.value = True Then
                
                If frmMain.MOSAICO.value = vbChecked Then
                    Dim aux As Long
                    Dim dy As Integer
                    Dim dX As Integer
                    If frmMain.DespMosaic.value = vbChecked Then
                        dy = Val(frmMain.DMLargo)
                        dX = Val(frmMain.DMAncho.Text)
                    Else
                        dy = 0
                        dX = 0
                    End If
                        
                    If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                        modEdicion.Deshacer_Add "Insertar Superficie' Hago deshacer"
                        MapInfo.Changed = 1 'Set changed flag
                        aux = Val(frmMain.cGrh.Text) + _
                        (((tY + dy) Mod frmMain.mLargo.Text) * frmMain.mAncho.Text) + ((tX + dX) Mod frmMain.mAncho.Text)
                        
                         If MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex <> aux Or MapData(tX, tY).Blocked <> frmMain.SelectPanel(2).value Then
                            MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                            InitGrh MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)), aux
                            
                        End If
                        
                    Else
                        modEdicion.Deshacer_Add "Insertar Auto-Completar Superficie' Hago deshacer"
                        MapInfo.Changed = 1 'Set changed flag
                        
                        Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                        
                        tXX = tX
                        tYY = tY
                        desptile = 0
                        
                        If tXX > XMaxMapSize Then tXX = XMaxMapSize
                        If tXX < XMinMapSize Then tXX = XMinMapSize
                        If tYY > YMaxMapSize Then tYY = YMaxMapSize
                        If tYY < YMinMapSize Then tYY = YMinMapSize
                        
                        For i = 1 To frmMain.mLargo.Text
                            For j = 1 To frmMain.mAncho.Text
                                If tYY >= 1 And tYY <= 100 Then
                                    If tXX >= 1 And tXX <= 100 Then
                                        aux = Val(frmMain.cGrh.Text) + desptile
                                        MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = aux
                                        InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.Text)), aux
                                        tXX = tXX + 1
                                        desptile = desptile + 1
                                    End If
                                End If
                            Next
                            tXX = tX
                            tYY = tYY + 1
                        Next
                        
                        tYY = tY
                        
                        
                    End If
                  
                Else
                
                    'Else Place graphic
                    If MapData(tX, tY).Blocked <> frmMain.SelectPanel(2).value Or MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex <> Val(frmMain.cGrh.Text) Then
                        modEdicion.Deshacer_Add "Quitar Superficie en Capa " & frmMain.cCapas.Text ' Hago deshacer
                        MapInfo.Changed = 1 'Set changed flag
                        MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)).GrhIndex = Val(frmMain.cGrh.Text)
                        'Setup GRH
                        InitGrh MapData(tX, tY).Graphic(Val(frmMain.cCapas.Text)), Val(frmMain.cGrh.Text)
                        
                    End If
                    
                End If
                
            End If
            
            '***************** Place blocked tile *****************
            If frmMain.cInsertarBloqueo.value = True Then
                If MapData(tX, tY).Blocked <> 1 Then
                    modEdicion.Deshacer_Add "Insertar Bloqueo" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Blocked = 1
                End If
                
            ElseIf frmMain.cQuitarBloqueo.value = True Then
                If MapData(tX, tY).Blocked <> 0 Then
                    modEdicion.Deshacer_Add "Quitar Bloqueo" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Blocked = 0
                End If
            End If
        
            '***************** Place exit *****************
            If frmMain.cInsertarTrans.value = True Then
                If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= NumOBJs And frmMain.cInsertarTransOBJ.value = True Then
                    If ObjData(Cfg_TrOBJ).ObjType = 19 Then
                        modEdicion.Deshacer_Add "Insertar Objeto de Translado" ' Hago deshacer
                        MapInfo.Changed = 1 'Set changed flag
                        InitGrh MapData(tX, tY).ObjGrh, ObjData(Cfg_TrOBJ).GrhIndex
                        MapData(tX, tY).OBJInfo.ObjIndex = Cfg_TrOBJ
                        MapData(tX, tY).OBJInfo.Amount = 1
                    End If
                End If
                
                If Val(frmMain.tTMapa.Text) < 0 Or Val(frmMain.tTMapa.Text) > 9000 Then
                    MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                    Exit Sub
                    
                ElseIf Val(frmMain.tTX.Text) < 0 Or Val(frmMain.tTX.Text) > XMaxMapSize Then
                    MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                    Exit Sub
                    
                ElseIf Val(frmMain.tTY.Text) < 0 Or Val(frmMain.tTY.Text) > XMaxMapSize Then
                    MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                    Exit Sub
                    
                End If
                
                If frmMain.cUnionManual.value = True Then
                    modEdicion.Deshacer_Add "Insertar Translado de Union Manual' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).TileExit.Map = Val(frmMain.tTMapa.Text)
                    
                    If tX >= 90 Then ' 21 ' derecha
                                MapData(tX, tY).TileExit.X = 12
                                MapData(tX, tY).TileExit.Y = tY
                    ElseIf tX <= 11 Then ' 9 ' izquierda
                        MapData(tX, tY).TileExit.X = 91
                        MapData(tX, tY).TileExit.Y = tY
                    End If
                    
                    If tY >= 91 Then ' 94 '''' hacia abajo
                        MapData(tX, tY).TileExit.Y = 11
                        MapData(tX, tY).TileExit.X = tX
                    ElseIf tY <= 10 Then ''' hacia arriba
                        MapData(tX, tY).TileExit.Y = 90
                        MapData(tX, tY).TileExit.X = tX
                    End If
                    
                Else
                    modEdicion.Deshacer_Add "Insertar Translado" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).TileExit.Map = Val(frmMain.tTMapa.Text)
                    MapData(tX, tY).TileExit.X = Val(frmMain.tTX.Text)
                    MapData(tX, tY).TileExit.Y = Val(frmMain.tTY.Text)
                End If
                    
            ElseIf frmMain.cQuitarTrans.value = True Then
                modEdicion.Deshacer_Add "Quitar Translado" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tX, tY).TileExit.Map = 0
                MapData(tX, tY).TileExit.X = 0
                MapData(tX, tY).TileExit.Y = 0
            End If
        
            '***************** Place NPC *****************
            If frmMain.cInsertarFunc(0).value = True Then
                If frmMain.cNumFunc(0).Text > 0 Then
                    NPCIndex = frmMain.cNumFunc(0).Text
                    If NPCIndex <> MapData(tX, tY).NPCIndex Then
                        modEdicion.Deshacer_Add "Insertar NPC" ' Hago deshacer
                        MapInfo.Changed = 1 'Set changed flag
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                        MapData(tX, tY).NPCIndex = NPCIndex
                    End If
                End If
                
            ElseIf frmMain.cInsertarFunc(1).value = True Then
                If frmMain.cNumFunc(1).Text > 0 Then
                    NPCIndex = frmMain.cNumFunc(1).Text
                    If NPCIndex <> (MapData(tX, tY).NPCIndex) Then
                        modEdicion.Deshacer_Add "Insertar NPC Hostil' Hago deshacer"
                        MapInfo.Changed = 1 'Set changed flag
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, tX, tY)
                        MapData(tX, tY).NPCIndex = NPCIndex
                    End If
                End If
                
            ElseIf frmMain.cQuitarFunc(0).value = True Or frmMain.cQuitarFunc(1).value = True Then
                If MapData(tX, tY).NPCIndex > 0 Then
                    modEdicion.Deshacer_Add "Quitar NPC" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).NPCIndex = 0
                    Call EraseChar(MapData(tX, tY).CharIndex)
                End If
            End If
        
            '***************** Control de Funcion de Objetos *****************
            If frmMain.cInsertarFunc(2).value = True Then ' Insertar Objeto
                If frmMain.cNumFunc(2).Text > 0 Then
                    ObjIndex = frmMain.cNumFunc(2).Text
                    
                    If MapData(tX, tY).OBJInfo.ObjIndex <> ObjIndex Or MapData(tX, tY).OBJInfo.Amount <> Val(frmMain.cCantFunc(2).Text) Then
                        modEdicion.Deshacer_Add "Insertar Objeto" ' Hago deshacer
                        MapInfo.Changed = 1 'Set changed flag
                        InitGrh MapData(tX, tY).ObjGrh, ObjData(ObjIndex).GrhIndex
                        MapData(tX, tY).OBJInfo.ObjIndex = ObjIndex
                        MapData(tX, tY).OBJInfo.Amount = Val(frmMain.cCantFunc(2).Text)
                        
                        Select Case ObjData(ObjIndex).ObjType
                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(tX, tY).Graphic(3) = MapData(tX, tY).ObjGrh
                        End Select
                        
                    End If
                End If
            ElseIf frmMain.cQuitarFunc(2).value = True Then ' Quitar Objeto
            
                If MapData(tX, tY).OBJInfo.ObjIndex <> 0 Or MapData(tX, tY).OBJInfo.Amount <> 0 Then
                    modEdicion.Deshacer_Add "Quitar Objeto" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    
                    If MapData(tX, tY).Graphic(3).GrhIndex = MapData(tX, tY).ObjGrh.GrhIndex Then MapData(tX, tY).Graphic(3).GrhIndex = 0
                    
                    MapData(tX, tY).ObjGrh.GrhIndex = 0
                    MapData(tX, tY).OBJInfo.ObjIndex = 0
                    MapData(tX, tY).OBJInfo.Amount = 0
                End If
                
            End If
            
            '***************** Control de Funcion de Triggers *****************
            If frmMain.cInsertarTrigger.value = True Then ' Insertar Trigger
                If MapData(tX, tY).Trigger <> frmMain.lListado(4).ListIndex Then
                    modEdicion.Deshacer_Add "Insertar Trigger" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Trigger = frmMain.lListado(4).ListIndex
                End If
                
            ElseIf frmMain.cQuitarTrigger.value = True Then ' Quitar Trigger
                If MapData(tX, tY).Trigger <> 0 Then
                    modEdicion.Deshacer_Add "Quitar Trigger" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Trigger = 0
                End If
                
            End If
            
            ' ***************** Control de Funcion de Particles! *****************
            If frmParticle.cmdAdd.value = True Then ' Insertar Particle
                MapInfo.Changed = 1
                General_Particle_Create CLng(frmParticle.lstParticle.ListIndex + 1), tX, tY, CLng(frmParticle.Life.Text)
                MapData(tX, tY).Particle_Index = CLng(frmParticle.lstParticle.ListIndex + 1)
                
            ElseIf frmParticle.cmdDel.value = True Then ' Quitar Particle
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tX, tY).Particle_Index = 0
                    MapData(tX, tY).Particle_Group_Index = 0
                
            End If
            
            '***************** LUCES **********************
            If frmLuces.cInsertarLuz.value Then
                If Val(frmLuces.cRango = 0) Then Exit Sub
                Call mDx8_Luces.Create_Light_To_Map(tX, tY, frmLuces.cRango, Val(frmLuces.R), Val(frmLuces.G), Val(frmLuces.B))
                Call mDx8_Luces.LightRenderAll
                
                With MapData(tX, tY).Light
                    .active = True
                    .range = frmLuces.cRango
                    .RGBCOLOR.a = 255
                    .RGBCOLOR.R = Val(frmLuces.R)
                    .RGBCOLOR.G = Val(frmLuces.G)
                    .RGBCOLOR.B = Val(frmLuces.B)
                    
                End With
                
                MapInfo.Changed = 1 'Set changed flag
                
            ElseIf frmLuces.cQuitarLuz.value Then
            
                With MapData(tX, tY).Light
                    .range = 0
                    .RGBCOLOR.a = 255
                    .RGBCOLOR.R = Val(frmLuces.R)
                    .RGBCOLOR.G = Val(frmLuces.G)
                    .RGBCOLOR.B = Val(frmLuces.B)
                    
                End With
    
                mDx8_Luces.Delete_Light_To_Map tX, tY
    
                MapInfo.Changed = 1 'Set changed flag
                
            End If
            
            If frmMain.LvBPintar.value Then
                MapData(tX, tY).ZonaIndex = frmMain.LstZona.ListIndex + 1
                MapInfo.Changed = 1
                
            ElseIf frmMain.LvBQuitar.value Then
                MapData(tX, tY).ZonaIndex = 1
                MapInfo.Changed = 1
                
            End If
    
    End Select

End Sub
