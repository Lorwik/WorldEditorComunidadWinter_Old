VERSION 5.00
Begin VB.Form frmRenderer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Renderizar"
   ClientHeight    =   15675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15675
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15000
      Left            =   120
      ScaleHeight     =   1000
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   15000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Renderizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image Smallpic 
      Height          =   5535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "frmRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************
' Capturar la imagen de controles
       
'   1 - Colocar un picturebox llamado picture1, un Command1 y un Command2 _
    2 - Agragar algunos controles _
    3 - Indicar en la Sub " Capturar_Imagen " .. el control a capturar
'*************************************************************
      
      
' Declaraciones del Api
      
'*************************************************************
' Función BitBlt para copiar la imagen del control en un picturebox
Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
      
' Recupera la imagen del área del control
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    
    Call MapCapture(False)

    
    Exit Sub

cmdAceptar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmRender.cmdAceptar_Click", Erl)
    Resume Next
    
End Sub

'*************************************************************
' Sub que copia la imagen del control en un picturebox
'*************************************************************
Public Sub Capturar_Imagen(Control As Control, Destino As Object)
          
    Dim hDC             As Long
    Dim Escala_Anterior As Integer
    Dim Ancho           As Long
    Dim Alto            As Long
          
    ' Para que se mantenga la imagen por si se repinta la ventana
    Destino.AutoRedraw = True
          
    On Error Resume Next

    ' Si da error es por que el control está dentro de un Frame _
      ya que  los Frame no tiene  dicha propiedad
    Escala_Anterior = Control.Container.ScaleMode
          
    If Err.Number = 438 Then
        ' Si el control está en un Frame, convierte la escala
        Ancho = ScaleX(Control.Width, vbTwips, vbPixels)
        Alto = ScaleY(Control.Height, vbTwips, vbPixels)
    Else
        ' Si no cambia la escala del  contenedor a pixeles
        Control.Container.ScaleMode = vbPixels
        Ancho = Control.Width
        Alto = Control.Height

    End If
          
    ' limpia el error
    On Error GoTo 0

    ' Captura el área de pantalla correspondiente al control
    hDC = GetWindowDC(Control.hwnd)
    
    ' Copia esa área al picturebox
    If ToWorldMap2 Then
        Call BitBlt(Destino.hDC, 0 - 50, 0 - 50, Ancho - 50, Alto - 50, hDC, 0, 0, vbSrcCopy)
    Else
        Call BitBlt(Destino.hDC, 0, 0, 3000, 3000, hDC, 0, 0, vbSrcCopy)

    End If
    
    ' Convierte la imagen anterior en un Mapa de bits
    Destino.Picture = Destino.image
    
    ' Borra la imagen ya que ahora usa el Picture
    Call Destino.Cls
          
    On Error Resume Next

    If Err.Number = 0 Then
        ' Si el control no está en un  Frame, restaura la escala del contenedor
        Control.Container.ScaleMode = Escala_Anterior

    End If
          
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo cmdSalir_Click_Err
    
    Unload Me

    Exit Sub

cmdSalir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmRender.cmdSalir_Click", Erl)
    Resume Next
End Sub

