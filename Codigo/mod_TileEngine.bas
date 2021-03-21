Attribute VB_Name = "mod_TileEngine"
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
' modDirectDraw
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Option Explicit

'<<<< PUBLICAS >>>>>

Public Normal_RGBList(3) As Long
Public temp_rgb(3) As Long

Public OffsetCounterX As Single
Public OffsetCounterY As Single

Public EngineRun As Boolean
Public FramesPerSec As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

'Tile size in pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long

'<<<<< PRIVADAS >>>>>

'Tamano del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private MouseTileX As Integer
Private MouseTileY As Integer

Private DrawBuffer As cDIBSection

'<<<<< CONSTANTES >>>>>>

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

Private Const MOVEMENT_SPEED As Single = 1

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

'Grafico que se muestra si hay error en un Grh
Public Const GRH_ERROR As Long = 22512

'<<<<< API >>>>>>

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************

    tX = (UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2) + 1
    tY = (UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2) + 1
End Sub

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

    'Update LastChar
    If CharIndex > LastChar Then LastChar = CharIndex
    NumChars = NumChars + 1
    
    With CharList(CharIndex)
    
    'Update head, body, ect.
    If Body > 0 Then _
    .Body = BodyData(Body)
    
    'If Head > 0 Then _
        .Head = HeadData(Head)
        
    .Heading = Heading
    
    'Reset moving stats
    .Moving = 0
    .MoveOffset.X = 0
    .MoveOffset.Y = 0
    
    'Update position
    .Pos.X = X
    .Pos.Y = Y
    
    'Make active
    .active = 1
    
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
    
    bRefreshRadar = True ' GS

End Sub

Sub EraseChar(CharIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
    If CharIndex = 0 Then Exit Sub
    'Make un-active
    CharList(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0
    
    'Update NumChars
    NumChars = NumChars - 1
    
    bRefreshRadar = True ' GS

End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

    '¿Es un Grh invalido?
    If GrhIndex <= 0 Or GrhIndex > grhCount Then GrhIndex = GRH_ERROR
    
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed
End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As Byte
    
    X = CharList(CharIndex).Pos.X
    Y = CharList(CharIndex).Pos.Y
    
    addX = nX - X
    addY = nY - Y
    
    If Sgn(addX) = 1 Then
        nHeading = eDireccion.EAST
    End If
    
    If Sgn(addX) = -1 Then
        nHeading = eDireccion.WEST
    End If
    
    If Sgn(addY) = -1 Then
        nHeading = eDireccion.NORTH
    End If
    
    If Sgn(addY) = 1 Then
        nHeading = eDireccion.SOUTH
    End If
    
    MapData(nX, nY).CharIndex = CharIndex
    CharList(CharIndex).Pos.X = nX
    CharList(CharIndex).Pos.Y = nY
    MapData(X, Y).CharIndex = 0
    
    CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
    CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)
    
    CharList(CharIndex).Moving = 1
    CharList(CharIndex).Heading = nHeading
    
    bRefreshRadar = True ' GS

End Sub

Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
    Dim loopc As Integer
    
    loopc = 1
    Do While CharList(loopc).active
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc

End Function

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

    LegalPos = True
    
    'Check to see if its out of bounds
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        LegalPos = False
        Exit Function
    End If
    
    'Check to see if its blocked
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    'Check for character
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
On Error GoTo Error

    If Grh.GrhIndex > grhCount Or GrhData(Grh.GrhIndex).NumFrames = 0 And GrhData(Grh.GrhIndex).FileNum = 0 Then
        Call InitGrh(Grh, GRH_ERROR) ' 23829
        Call AddtoRichTextBox(frmMain.StatTxt, "Error en Grh. Posicion: X:" & X & " Y:" & Y, 255, 0, 0)
    End If

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed) * MOVEMENT_SPEED

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth * ScaleX - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        #If Desarrollo = 0 Then
            Call RegistrarError(Err.Number, "Error in Draw_Grh, " & Err.Description, "Draw_Grh", Erl)
            MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
            Call CloseClient
        
        #Else
            Debug.Print "Error en Draw_Grh en el grh" & CurrentGrhIndex & ", " & Err.Description & ", (" & Err.Number & ")"
        #End If
    End If
End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)
    End With
    
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
    
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
    
End Sub

' [Loopzer]
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, Y + DeSeleccionOY) = DeSeleccionMap(X, Y)
        Next
    Next
End Sub

Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    
    Dim X As Integer
    Dim Y As Integer
    
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    
    UltimoX = SobreX
    UltimoY = SobreY
    
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SobreX, Y + SobreY)
        Next
    Next
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, Y + SobreY) = SeleccionMap(X, Y)
        Next
    Next
    Seleccionando = False
End Sub

Public Sub AccionSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + Y
        Next
    Next
    Seleccionando = False
End Sub

Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    
    Seleccionando = False
End Sub

Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, Y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub

Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim Y As Integer
    
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            SeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
End Sub

Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    
End Sub
' [/Loopzer]

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************

On Error GoTo RenderScreen_Err

    Dim Y                As Long     'Keeps track of where on map we are
    Dim X                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen
    Dim screenmaxY       As Integer  'End Y pos on current screen
    Dim screenminX       As Integer  'Start X pos on current screen
    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim minY             As Integer  'Start Y pos on current map
    Dim maxY             As Integer  'End Y pos on current map
    Dim minX             As Integer  'Start X pos on current map
    Dim maxX             As Integer  'End X pos on current map
    
    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim Sobre            As Long
    Dim Grh              As Grh                  'Temp Grh for show tile and blocked
    Dim bCapa            As Byte                 'cCapas ' 31/05/2006 - GS, control de Capas
    Dim iGrhIndex        As Integer  'Usado en el Layer 1
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth + 1
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize

    ' 31/05/2006 - GS, control de Capas
    If Val(frmMain.cCapas.Text) >= 1 And (frmMain.cCapas.Text) <= 4 Then
        bCapa = Val(frmMain.cCapas.Text)
    Else
        bCapa = 1
    End If
    
    GenerarVista 'Loopzer
    
     'Draw floor layer
    For Y = screenminY To screenmaxY
    
        For X = screenminX To screenmaxX
            
            'Previsualización
            '*******************************
            If SobreX = X And SobreY = Y Then
                        
                ' Pone Grh !
                Sobre = -1

                If frmMain.cSeleccionarSuperficie.value = True Then
                    Sobre = MapData(X, Y).Graphic(bCapa).GrhIndex

                    If frmMain.MOSAICO.value = vbChecked Then
                        Dim aux As Long
                        Dim dy  As Integer
                        Dim dX  As Integer

                        If frmMain.DespMosaic.value = vbChecked Then
                            dy = Val(frmMain.DMLargo.Text)
                            dX = Val(frmMain.DMAncho.Text)
                        Else
                            dy = 0
                            dX = 0

                        End If

                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmMain.cGrh.Text) + (((Y + dy) Mod frmMain.mLargo.Text) * frmMain.mAncho.Text) + ((X + dX) Mod frmMain.mAncho.Text)

                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, Y).Graphic(bCapa), aux

                            End If

                        Else
                            aux = Val(frmMain.cGrh.Text) + (((Y + dy) Mod frmMain.mLargo.Text) * frmMain.mAncho.Text) + ((X + dX) Mod frmMain.mAncho.Text)

                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, Y).Graphic(bCapa), aux

                            End If

                        End If

                    Else

                        If MapData(X, Y).Graphic(bCapa).GrhIndex <> Val(frmMain.cGrh.Text) Then
                            MapData(X, Y).Graphic(bCapa).GrhIndex = Val(frmMain.cGrh.Text)
                            InitGrh MapData(X, Y).Graphic(bCapa), Val(frmMain.cGrh.Text)

                        End If

                    End If

                End If

            Else
            
                Sobre = -1
            
            End If
                
            If InMapBounds(X, Y) Then
                
                PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
                
                'Layer 1 **********************************
                If MapData(X, Y).Graphic(1).GrhIndex And VerCapa1 Then _
                    Call Draw_Grh(MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
                    
                'Layer 2 **********************************
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then _
                    Call Draw_Grh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
                    
                If Sobre >= 0 Then
                    If MapData(X, Y).Graphic(bCapa).GrhIndex <> Sobre Then
                        MapData(X, Y).Graphic(bCapa).GrhIndex = Sobre
                        InitGrh MapData(X, Y).Graphic(bCapa), Sobre
                            
                        If MapData(X, Y).Graphic(bCapa).GrhIndex = GRH_ERROR Then _
                            MapData(X, Y).Graphic(bCapa).GrhIndex = 0
                    End If
                    
                End If
                
            End If
            
            ScreenX = ScreenX + 1
        Next X
        
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
    
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
        
            If InMapBounds(X, Y) Then
                PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX - 32
                PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY - 32
                
                If X > XMaxMapSize Or X < -3 Then Exit For ' 30/05/2006
                
                With MapData(X, Y)
                    
                     'Object Layer **********************************
                     If .OBJInfo.ObjIndex <> 0 And VerObjetos Then
                         Call Draw_Grh(.ObjGrh, _
                                PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                     End If
                    
                    'Char layer **********************************
                    If .CharIndex <> 0 And VerNpcs Then
                        'Llamada al CharRender
                    End If
                     
                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 And VerCapa3 Then _
                        Call Draw_Grh(.Graphic(3), _
                                PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                                
                    'Particulas *****************************************
                    If .Particle_Group_Index Then _
                        Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                         
                End With
                
            End If
            ScreenX = ScreenX + 1
            
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize

    For Y = minY To maxY

        ScreenX = minXOffset - TileBufferSize

        For X = minX To maxX
            
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX - 32
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY - 32
        
            If X < XMaxMapSize + 1 And X > 0 And Y < XMaxMapSize + 1 And Y > 0 Then ' 30/05/2006
            
                '<----- Layer 4 ----->
                If MapData(X, Y).Graphic(4).GrhIndex <> 0 And (frmMain.mnuVerCapa4.Checked = True) Then
                    Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
                    
                End If

                If MapData(X, Y).TileExit.Map <> 0 And VerTranslados Then
                    Grh.GrhIndex = 3
                    Grh.FrameCounter = 1
                    Grh.Started = 0
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
                End If
                
                'Show blocked tiles
                If VerBlockeados And MapData(X, Y).Blocked = 1 Then
                    Grh.GrhIndex = 4
                    Grh.FrameCounter = 1
                    Grh.Started = 0
                    
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 1)
                        
                End If
                
                If VerGrilla Then
                    Grh.GrhIndex = 2
                    Grh.FrameCounter = 1
                    Grh.Started = 0
                    
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, Normal_RGBList(), 0)
                        
                End If

                If VerTriggers Then '4978
                    If MapData(X, Y).Trigger > 0 Then _
                        Call DrawText(PixelOffsetXTemp, PixelOffsetYTemp, MapData(X, Y).Trigger, -1, False, 1)
                End If
                    
                If Seleccionando Then
                        If X >= SeleccionIX And Y >= SeleccionIY Then
                            If X <= SeleccionFX And Y <= SeleccionFY Then
                                Grh.GrhIndex = 2
                                Grh.FrameCounter = 1
                                Grh.Started = 0
                                
                                Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
                            End If
                        End If
                End If
    
            End If
            ScreenX = ScreenX + 1
        Next X
        
        ScreenY = ScreenY + 1
    Next Y

RenderScreen_Err:

    If Err.Number Then
        Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.RenderScreen")
    End If
End Sub

Public Sub RenderPreview()
    Dim destRect     As RECT
    
    Dim i As Integer, j As Integer
    Dim Cont As Integer
    
    With destRect
        .Bottom = frmMain.PreviewGrh.Height
        .Right = frmMain.PreviewGrh.Width
    End With
    
    'Clear the inventory window
    Call Engine_BeginScene
    
    If frmMain.MOSAICO = vbUnchecked Then
        Call Draw_GrhIndex(CurrentGrh.GrhIndex, frmMain.PreviewGrh.Height / 2, frmMain.PreviewGrh.Width - 50, 1, Normal_RGBList(), 0)
        
    Else
        For i = 1 To CInt(Val(frmMain.mLargo))
            For j = 1 To CInt(Val(frmMain.mAncho))
            
                Call Draw_GrhIndex(CurrentGrh.GrhIndex, (j - 1) * 32, (i - 1) * 32, 1, Normal_RGBList(), 0)
                
                If Cont < CInt(Val(frmMain.mLargo)) * CInt(Val(frmMain.mAncho)) Then _
                    Cont = Cont + 1: CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
            Next j
        Next i
        
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont
    End If
    
    frmMain.PreviewGrh.AutoRedraw = False

    Call Engine_EndScene(destRect, frmMain.PreviewGrh.hwnd)

    Call DrawBuffer.LoadPictureBlt(frmMain.PreviewGrh.hDC)

    frmMain.PreviewGrh.AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmMain.PreviewGrh.hDC, 0, 0, frmMain.PreviewGrh.Width, frmMain.PreviewGrh.Height, 0, 0, vbSrcCopy)
End Sub

Function PixelPos(X As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Public Sub InitTileEngine(ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer)
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
'Configures the engine to start running.
'***************************************************

On Error GoTo ErrorHandler:

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)

    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = 50
    UserPos.Y = 50
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

On Error GoTo 0
    
    'Cargamos indice de graficos.
    'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
    Call LoadGrhData
    
    'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
    With frmCargando

        .P1.Visible = True
        .L(0).Visible = True
        .X.Caption = "Cargando Cuerpos..."
        modIndices.CargarCuerpos
        DoEvents
        
        .P2.Visible = True
        .L(1).Visible = True
        .X.Caption = "Cargando Cabezas..."
        modIndices.CargarCabezas
        DoEvents
        
        .P3.Visible = True
        .L(2).Visible = True
        .X.Caption = "Cargando NPC's..."
        modIndices.CargarIndicesNPC
        DoEvents
        
        .P4.Visible = True
        .L(3).Visible = True
        .X.Caption = "Cargando Objetos..."
        modIndices.CargarIndicesOBJ
        DoEvents
        
        .P5.Visible = True
        .L(4).Visible = True
        .X.Caption = "Cargando Triggers..."
        modIndices.CargarIndicesTriggers
        DoEvents
        
        .P6.Visible = True
        .L(5).Visible = True
        DoEvents
    
    End With
    
    Call LoadGraphics
    Call CargarParticulas

    Exit Sub
    
ErrorHandler:

    Call RegistrarError(Err.Number, Err.Description, "Mod_TileEngine.InitTileEngine")
    
    Call CloseClient
    
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame()

On Error GoTo ErrorHandler:

    If EngineRun Then
        
        Call Engine_BeginScene
            
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
    
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
    
            End If
                    
        End If
                
            '****** Move screen Up and Down if needed ******
        If AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
    
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                        
            End If
    
        End If
            
        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

        ' Calculamos los FPS y los mostramos
        Call Engine_Update_FPS
        Call DrawText(10, 5, "FPS: " & mod_TileEngine.FPS, -1, False)
        Call DrawText(10, 20, "Mouse: " & MousePos, -1, False)
    
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
            
        Call Engine_EndScene(MainScreenRect, 0)
    
    End If
    
ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call mDx8_Engine.Engine_DirectX8_Init
        
        Call LoadGraphics
    
    End If
  
End Sub

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub Device_Textured_Render(ByVal X As Single, ByVal Y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef Color() As Long, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)

        Dim Texture As Direct3DTexture8
        
        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
        With SpriteBatch

                Call .SetTexture(Texture)
                    
                Call .SetAlpha(Alpha)
                
                If TextureWidth <> 0 And TextureHeight <> 0 Then
                    Call .Draw(X, Y, Width * ScaleX, Height * ScaleY, Color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
                Else
                    Call .Draw(X, Y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , angle)
                End If
                
        End With
        
End Sub

