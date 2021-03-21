Attribute VB_Name = "modGeneral"
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
' modGeneral
'
' @remarks Funciones Generales
' @author unkwown
' @version 0.4.11
' @date 20061015

Option Explicit

Private lFrameTimer As Long

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

If HotKeysAllow = False Then Exit Sub
        '[Loopzer]
        'If GetKeyState(vbKeyControl) < 0 Then
        '    If Seleccionando Then
        '        If GetKeyState(vbKeyC) < 0 Then CopiarSeleccion
        '        If GetKeyState(vbKeyX) < 0 Then CortarSeleccion
        '        If GetKeyState(vbKeyB) < 0 Then BlockearSeleccion
        '        If GetKeyState(vbKeyD) < 0 Then AccionSeleccion
        ''    Else
        '        If GetKeyState(vbKeyS) < 0 Then DePegar ' GS
        '        If GetKeyState(vbKeyV) < 0 Then PegarSeleccion
        '    End If
        'End If
        '[/Loopzer]
    
    
    If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.Y < 1 Then Exit Sub ' 10
        If LegalPos(UserPos.X, UserPos.Y - 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.Y = UserPos.Y - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y - 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > 100 Then Exit Sub ' 89
        If LegalPos(UserPos.X + 1, UserPos.Y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.Y > 100 Then Exit Sub ' 92
        If LegalPos(UserPos.X, UserPos.Y + 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.Y = UserPos.Y + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y + 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < 1 Then Exit Sub ' 12
        If LegalPos(UserPos.X - 1, UserPos.Y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If
    
End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim i As Integer
Dim lastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
lastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid(Text, lastPos + 1, (InStr(lastPos + 1, Text, Seperator, vbTextCompare) - 1) - (lastPos))
            Exit Function
        End If
        lastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = mid(Text, lastPos + 1)
End If

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

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Private Sub CargarMapIni()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
    On Error GoTo Fallo
    Dim tStr As String
    Dim Leer As New clsIniReader
    
    IniPath = App.Path & "\"
    
    'Si el WorldEditor.ini no existe, tomamos estos parametros por defecto
    If Not FileExist(IniPath & "Datos\WorldEditor.ini", vbArchive) Then
        frmMain.mnuGuardarUltimaConfig.Checked = True
        DirRecursos = IniPath & "Recursos\"
        DirMidi = IniPath & "MIDI\"
        frmMusica.fleMusicas.Path = DirMidi
        DirDats = IniPath & "DATS\"
        MaxGrhs = 15000
        UserPos.X = 50
        UserPos.Y = 50
        PantallaX = 19
        PantallaY = 22
        MsgBox "Falta el archivo 'WorldEditor.ini' de configuración.", vbInformation
        Exit Sub
    End If
    
    Call Leer.Initialize(IniPath & "Datos\WorldEditor.ini")
    
    ' Obj de Translado
    Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
    frmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
    frmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))
    frmMain.mnuUtilizarDeshacer.Checked = Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer"))
    
    ' Guardar Ultima Configuracion
    frmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))
    
    ' Index
    MaxGrhs = Val(GetVar(IniPath & "Datos\WorldEditor.ini", "INDEX", "MaxGrhs"))
    If MaxGrhs < 1 Then MaxGrhs = 15000
    
    'Reciente
    frmMain.Dialog.InitDir = Leer.GetValue("PATH", "UltimoMapa")
    DirRecursos = autoCompletaPath(Leer.GetValue("PATH", "DirRecursos"))
    
    '-------
    'Carga de rutas
    '-----------------------------------------
    
    '****
    'RUTA DE GRAFFICOS
    '*****************
    If DirRecursos = "\" Then
        DirRecursos = IniPath & "Recursos\"
    End If
    
    If FileExist(DirRecursos, vbDirectory) = False Then
        MsgBox "El directorio de Graficos es incorrecto", vbCritical + vbOKOnly
        End
    End If
    
    '****
    'RUTA DE MIDI
    '*****************
    DirMidi = autoCompletaPath(Leer.GetValue("PATH", "DirMidi"))
    
    If DirMidi = "\" Then
        DirMidi = IniPath & "MIDI\"
    End If
    
    If FileExist(DirMidi, vbDirectory) = False Then
        MsgBox "El directorio de MIDI es incorrecto", vbCritical + vbOKOnly
        End
    End If
    
    frmMusica.fleMusicas.Path = DirMidi
    
    '****
    'RUTA DE DATS
    '*****************
    DirDats = autoCompletaPath(Leer.GetValue("PATH", "DirDats"))
    If DirDats = "\" Then
        DirDats = IniPath & "DATS\"
    End If
    
    If FileExist(DirDats, vbDirectory) = False Then
        MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
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
    frmMain.mnuVerBloqueos.Checked = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
    frmMain.cVerTriggers.value = frmMain.mnuVerTriggers.Checked
    frmMain.cVerBloqueos.value = frmMain.mnuVerBloqueos.Checked
    
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

Public Sub Main()
'*************************************************
'Author: Unkwown
'Last modified: 25/11/08 - GS
'*************************************************
On Error Resume Next

    If App.PrevInstance = True Then End
    
    Call GenerateContra
    Call CargarMapIni
    Call LeerConfiguracion
    
    If FileExist(IniPath & "Datos\WorldEditor.jpg", vbArchive) Then frmCargando.Picture1.Picture = LoadPicture(IniPath & "Datos\WorldEditor.jpg")
    
    frmCargando.verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
    frmCargando.Show
    frmCargando.SetFocus
    DoEvents
    
    frmCargando.X.Caption = "Iniciando DirectSound..."
    'IniciarDirectSound
    DoEvents
    
    frmCargando.X.Caption = "Cargando Indice de Superficies..."
    modIndices.CargarIndicesSuperficie
    DoEvents
    
    frmCargando.X.Caption = "Indexando Cargado de Imagenes..."
    DoEvents

    frmCargando.X.Caption = "Iniciando motor grafico..."
    DoEvents
    'Iniciamos el Engine de DirectX 8
    Call mDx8_Engine.Engine_DirectX8_Init
    
    'Tile Engine
    Call InitTileEngine(frmMain.hWnd, 32, 32, 8, 8)
    
    Call mDx8_Engine.Engine_DirectX8_Aditional_Init
    
    frmCargando.SetFocus
    frmCargando.X.Caption = "Iniciando Ventana de Edición..."
    DoEvents
    If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) = 0 Then
        frmMain.mnuManual.Enabled = False
        frmMain.mnuManual.Caption = "&Manual (no implementado)"
    End If
    frmCargando.Hide
    
    frmMain.Show
    modMapIO.NuevoMapa
    DoEvents
    
    prgRun = True
    
    lFrameTimer = GetTickCount

    Do While prgRun

        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            Call CheckKeys

        End If
        
        If CurrentGrh.GrhIndex = 0 Then
            InitGrh CurrentGrh, 1
        End If
        
        If bRefreshRadar = True Then
            Call RefreshAllChars
            bRefreshRadar = False
        End If
        
        If frmMain.PreviewGrh.Visible = True Then
            Call RenderPreview
        End If
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            
            lFrameTimer = GetTickCount
        End If
        
        DoEvents
    Loop
        
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa frmMain.Dialog.filename
        End If
    End If
    
    Call CloseClient

End Sub

Public Function GetVar(File As String, Main As String, Var As String) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
    Dim L As Integer
    Dim Char As String
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
    szReturn = vbNullString
    sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), File
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Sub WriteVar(File As String, Main As String, Var As String, value As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
writeprivateprofilestring Main, Var, value, File
End Sub

Public Sub ToggleWalkMode()
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************
On Error GoTo fin:
If WalkMode = False Then
    WalkMode = True
Else
    frmMain.mnuModoCaminata.Checked = False
    WalkMode = False
End If

If WalkMode = False Then
    'Erase character
    Call EraseChar(UserCharIndex)
    MapData(UserPos.X, UserPos.Y).CharIndex = 0
Else
    'MakeCharacter
    If LegalPos(UserPos.X, UserPos.Y) Then
        Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.Y)
        UserCharIndex = MapData(UserPos.X, UserPos.Y).CharIndex
        frmMain.mnuModoCaminata.Checked = True
    Else
        MsgBox "ERROR: Ubicacion ilegal."
        WalkMode = False
    End If
End If
fin:
End Sub

Public Sub FixCoasts(ByVal GrhIndex As Long, ByVal X As Integer, ByVal Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If GrhIndex = 7284 Or GrhIndex = 7290 Or GrhIndex = 7291 Or GrhIndex = 7297 Or _
   GrhIndex = 7300 Or GrhIndex = 7301 Or GrhIndex = 7302 Or GrhIndex = 7303 Or _
   GrhIndex = 7304 Or GrhIndex = 7306 Or GrhIndex = 7308 Or GrhIndex = 7310 Or _
   GrhIndex = 7311 Or GrhIndex = 7313 Or GrhIndex = 7314 Or GrhIndex = 7315 Or _
   GrhIndex = 7316 Or GrhIndex = 7317 Or GrhIndex = 7319 Or GrhIndex = 7321 Or _
   GrhIndex = 7325 Or GrhIndex = 7326 Or GrhIndex = 7327 Or GrhIndex = 7328 Or GrhIndex = 7332 Or _
   GrhIndex = 7338 Or GrhIndex = 7339 Or GrhIndex = 7345 Or GrhIndex = 7348 Or _
   GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or GrhIndex = 7352 Or _
   GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or _
   GrhIndex = 7354 Or GrhIndex = 7357 Or GrhIndex = 7358 Or GrhIndex = 7360 Or _
   GrhIndex = 7362 Or GrhIndex = 7363 Or GrhIndex = 7365 Or GrhIndex = 7366 Or _
   GrhIndex = 7367 Or GrhIndex = 7368 Or GrhIndex = 7369 Or GrhIndex = 7371 Or _
   GrhIndex = 7373 Or GrhIndex = 7375 Or GrhIndex = 7376 Then MapData(X, Y).Graphic(2).GrhIndex = 0

End Sub

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Randomize Timer
RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function


''
' Actualiza todos los Chars en el mapa
'

Public Sub RefreshAllChars()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer
frmMain.ApuntadorRadar.Move UserPos.X - 12, UserPos.Y - 10
frmMain.picRadar.Cls
For loopc = 1 To LastChar
    If CharList(loopc).Active = 1 Then
        MapData(CharList(loopc).Pos.X, CharList(loopc).Pos.Y).CharIndex = loopc
        If CharList(loopc).Heading <> 0 Then
            frmMain.picRadar.ForeColor = vbGreen
            frmMain.picRadar.Line (0 + CharList(loopc).Pos.X, 0 + CharList(loopc).Pos.Y)-(2 + CharList(loopc).Pos.X, 0 + CharList(loopc).Pos.Y)
            frmMain.picRadar.Line (0 + CharList(loopc).Pos.X, 1 + CharList(loopc).Pos.Y)-(2 + CharList(loopc).Pos.X, 1 + CharList(loopc).Pos.Y)
        End If
    End If
Next loopc
bRefreshRadar = False
End Sub

''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    If Trabajando = vbNullString Then
        Trabajando = "Nuevo Mapa"
    End If
    
    frmMain.Caption = "WorldEditor Comunidad Winter v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"
    
    If Editado = True Then
        frmMain.Caption = frmMain.Caption & " (modificado)"
    End If
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Lorwik
'Last Modify Date: 20/03/2021
'**************************************************************
    On Error Resume Next

    Dim mifrm As Form

    EngineRun = False
    
    'Stop tile engine
    Call Engine_DirectX8_End

    'Destruimos los objetos publicos creados
    Set SurfaceDB = Nothing
    
    For Each mifrm In Forms
        Unload mifrm
    Next
    
    End
    
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                    ByVal Text As String, _
                    Optional ByVal Red As Integer = -1, _
                    Optional ByVal Green As Integer, _
                    Optional ByVal Blue As Integer, _
                    Optional ByVal bold As Boolean = False, _
                    Optional ByVal italic As Boolean = False, _
                    Optional ByVal bCrLf As Boolean = True, _
                    Optional ByVal Alignment As Byte = rtfLeft)
    
'****************************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D apperance!
'****************************************************
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'Jopi 17/08/2019 : Consola transparente.
'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
'****************************************************
    With RichTextBox
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        ' 0 = Left
        ' 1 = Center
        ' 2 = Right
        .SelAlignment = Alignment

        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        
        .SelText = Text

        ' Esto arregla el bug de las letras superponiendose la consola del frmMain
        If Not RichTextBox = frmMain.StatTxt Then RichTextBox.Refresh

    End With
End Sub

