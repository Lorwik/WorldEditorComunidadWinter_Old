Attribute VB_Name = "mDx8_Engine"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' No matter what you do with DirectX8, you will need to start with
' the DirectX8 object. You will need to create a new instance of
' the object, using the New keyword, rather than just getting a
' pointer to it, since there's nowhere to get a pointer from yet (duh!).

Public DirectX As New DirectX8

' The D3DX8 object contains lots of helper functions, mostly math
' to make Direct3D alot easier to use. Notice we create a new
' instance of the object using the New keyword.
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8

' The Direct3DDevice8 represents our rendering device, which could
' be a hardware or a software device. The great thing is we still
' use the same object no matter what it is
Public DirectDevice As Direct3DDevice8

' The D3DDISPLAYMODE type structure that holds
' the information about your current display adapter.
Public DispMode  As D3DDISPLAYMODE
    
' The D3DPRESENT_PARAMETERS type holds a description of the way
' in which DirectX will display it's rendering.
Public D3DWindow As D3DPRESENT_PARAMETERS

Public SurfaceDB As New clsTextureManager
Public SpriteBatch As New clsBatch

Private Viewport As D3DVIEWPORT8
Private Projection As D3DMATRIX
Private View As D3DMATRIX

Public Engine_BaseSpeed As Single

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamano muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Public ScreenWidth As Long
Public ScreenHeight As Long

Public MainScreenRect As RECT

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private EndTime As Long

Public Sub Engine_DirectX8_Init()
    On Error GoTo EngineHandler:

    Dim StError As String

    ' Initialize all DirectX objects.
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    StError = "No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados."
 
    If ClientSetup.OverrideVertexProcess > 0 Then
        
        Select Case ClientSetup.OverrideVertexProcess
            
            Case 1:
                If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                    Call MsgBox(StError)
                    End
                End If
            
            
            Case 2:
                If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                    Call MsgBox(StError)
                    End
                End If

            
            Case 3:
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    Call MsgBox(StError)
                    End
                End If
        End Select
        
    Else
        'Detectamos el modo de renderizado mas compatible con tu PC.
        If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
            
                    Call MsgBox(StError)
                    End
                
                End If
            End If
        End If
    End If

    Call SimpleLogError("Modo de procesado elegido.")

    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, ScreenWidth, ScreenHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)

    ' Set rendering options
    Call Engine_Init_RenderStates
    
    'Carga dinamica de texturas por defecto.
    Set SurfaceDB = New clsTextureManager
    
    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialise(2000)
    
    'Inicializamos el resto de sistemas.
    Call Engine_DirectX8_Aditional_Init
    
    EndTime = timeGetTime
    
    Exit Sub
EngineHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "mDx8_Engine.Engine_DirectX8_Init")
    
    Call CloseClient
End Sub

Private Function Engine_Init_DirectDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean

    'Establecemos cual va a ser el tamano del render.
    ScreenWidth = 6400
    ScreenHeight = 6400
    
    ' Retrieve the information about your current display adapter.
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
        ' Fill the D3DPRESENT_PARAMETERS type, describing how DirectX should
    ' display it's renders.
    With D3DWindow
        .Windowed = True
        
        ' The swap effect determines how the graphics get from the backbuffer to the screen.
        ' D3DSWAPEFFECT_DISCARD:
        '   Means that every time the render is presented, the backbuffer
        '   image is destroyed, so everything must be rendered again.
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = ScreenWidth
        .BackBufferHeight = ScreenHeight
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With
    
    If Not DirectDevice Is Nothing Then
        Set DirectDevice = Nothing
    End If
    
    ' Create the rendering device.
    ' Here we request a Hardware or Mixed rasterization.
    ' If your computer does not have this, the request may fail, so use
    ' D3DDEVTYPE_REF instead of D3DDEVTYPE_HAL if this happens. A real
    ' program would be able to detect an error and automatically switch device.
    ' We also request software vertex processing, which means the CPU has to
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, D3DCREATEFLAGS, D3DWindow)
    
    'Lo pongo xq es bueno saberlo...
    Select Case D3DCREATEFLAGS
    
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: SOFTWARE"
            
    End Select
    
    'Everything was successful
    Engine_Init_DirectDevice = True
    
    Exit Function
    
ErrorDevice:
    
    'Destroy the D3DDevice so it can be remade
    Set DirectDevice = Nothing

    'Return a failure
    Engine_Init_DirectDevice = False
    
End Function

Private Sub Engine_Init_RenderStates()

    'Set the render states
    With DirectDevice
    
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
    End With
    
End Sub

Public Sub Engine_DirectX8_End()
'***************************************************
'Author: Standelf
'Last Modification: 26/05/2010
'Destroys all DX objects
'***************************************************
On Error Resume Next
    Dim i As Byte
    
    '   DeInit Lights
    Call DeInit_LightEngine
    
    '   Clean Particles
    Call Particle_Group_Remove_All
    
    '   Clean Texture
    Call DirectDevice.SetTexture(0, Nothing)
    
    '   Borrar DBI Surface
    Call CleanDrawBuffer
    
    '   Erase Data
    Erase MapData()
    Erase CharList()
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    Set SpriteBatch = Nothing
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    FPS = 101
    FramesPerSecCounter = 101

    TileBufferSize = 5
    
    Engine_BaseSpeed = 0.018
    
    With MainScreenRect
        .Bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With

    'Inicializamos y cargamos los graficos de las Fonts.
    Call mDx8_Text.Engine_Init_FontTextures
    
    If Not prgRun Then
        
        ' Seteamos algunos colores por adelantado y unica vez.
        Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)
        
        ' Inicializamos otros sistemas.
        Call mDx8_Text.Engine_Init_FontSettings
        Call Init_MeteoEngine
        
        ' Inicializa DIB surface, un buffer usado para dejar imagenes estaticas en PictureBox
        Call PrepareDrawBuffer
        
    End If
    
End Sub

Public Function Engine_ElapsedTime() As Long
'**************************************************************
'Gets the time that past since the last call
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************
Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time

End Function

Public Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX
'*****************************************************************

    Engine_PixelPosX = (X - 1) * 32
    
End Function

Public Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * 32
    
End Function

Public Sub Engine_D3DColor_To_RGB_List(rgb_list() As Long, color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    rgb_list(0) = D3DColorARGB(color.a, color.R, color.G, color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Function SetARGB_Alpha(rgb_list() As Long, Alpha As Byte) As Long()

    '***************************************************
    'Author: Juan Manuel Couso (Cucsifae)
    'Last Modification: 29/08/18
    'Obtiene un ARGB list le modifica el alpha y devuelve una copia
    '***************************************************
    Dim TempColor        As D3DCOLORVALUE
    Dim tempARGB(0 To 3) As Long

    'convertimos el valor del rgb list a D3DCOLOR
    Call ARGBtoD3DCOLORVALUE(rgb_list(1), TempColor)

    'comprobamos ue no se salga del rango permitido
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    
    'seteamos el alpha
    TempColor.a = Alpha
    
    'generamos el nuevo RGB_List
    Call Engine_D3DColor_To_RGB_List(tempARGB(), TempColor)

    SetARGB_Alpha = tempARGB()

End Function

Private Function Engine_Collision_Between(ByVal value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte
'*****************************************************************
'Find if a value is between two other values (used for line collision)
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Between
'*****************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If value >= Bound2 Then
            If value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If value >= Bound1 Then
            If value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
End Function

Public Sub Engine_BeginScene(Optional ByVal color As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD Clear & BeginScene
'***************************************************

    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, color, 1#, 0)
    Call SpriteBatch.Begin
    
End Sub

Public Sub Engine_EndScene(ByRef destRect As RECT, Optional ByVal hWndDest As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD EndScene & Present
'***************************************************
On Error GoTo DeviceHandler:

    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
        
    If hWndDest = 0 Then
        Call DirectDevice.Present(destRect, ByVal 0&, ByVal 0&, ByVal 0&)
    
    Else
        Call DirectDevice.Present(destRect, ByVal 0, hWndDest, ByVal 0)
    
    End If
    
    Exit Sub
    
DeviceHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then

        Call mDx8_Engine.Engine_DirectX8_Init

        Call LoadGraphics

    End If
    
End Sub

Public Function ZoomOffset(ByVal offset As Byte) As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/01/2011
'**************************************************************

    ZoomOffset = IIf((offset = 1), (ScreenHeight - MainScreenRect.Bottom) / 2, (ScreenWidth - MainScreenRect.Right) / 2)
    
End Function

Public Sub Engine_Update_FPS()
    '***************************************************
    'Author: ???
    'Last Modification: ????
    'Calculate $ Limitate (if active) FPS.
    '***************************************************

    If ClientSetup.LimiteFPS Then
        While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter
            Call Sleep(5)
        Wend
    End If

    If FPSLastCheck + 1000 < timeGetTime Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = timeGetTime
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1

    End If

End Sub

Public Sub Engine_Get_ARGB(color As Long, Data As D3DCOLORVALUE)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************
    
    Dim a As Long, R As Long, G As Long, B As Long
        
    If color < 0 Then
        a = ((color And (&H7F000000)) / (2 ^ 24)) Or &H80&
    Else
        a = color / (2 ^ 24)
    End If
    
    R = (color And &HFF0000) / (2 ^ 16)
    G = (color And &HFF00&) / (2 ^ 8)
    B = (color And &HFF&)
    
    With Data
        .a = a
        .R = R
        .G = G
        .B = B
    End With
        
End Sub

