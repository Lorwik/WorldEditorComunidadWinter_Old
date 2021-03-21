Attribute VB_Name = "modCarga"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    LimiteFPS As Boolean
    OverrideVertexProcess As Byte
End Type

Public ClientSetup As tSetupMods

Public Function WEConfigDir() As String
    WEConfigDir = App.Path & "\Datos\WorldEditor.ini"
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

    Call IniciarCabecera

    Set Lector = New clsIniManager
    Call Lector.Initialize(WEConfigDir)
    
    With ClientSetup
    
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        
    End With

  Exit Sub
  
fileErr:

    If Err.Number <> 0 Then
       MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.Number & " : " & Err.Description)
       End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub
