VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Joypad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Private registers As JoypadRegisters
Public keys As Byte

Public Sub Class_Initialize()

    registers.p1 = 0 ' $FF00 P1
    
    Me.keys = 255

End Sub

Public Function ReadByte(ByVal address As Long) As Byte
    Select Case address
        Case 65280
            If registers.p1 = 16 Then
                ReadByte = Utils.RightShift(Me.keys, 4) And 15
                Exit Function
            End If
            If registers.p1 = 32 Then
                ReadByte = Me.keys And 15
                Exit Function
            End If
        Case Else
            ReadByte = 0
    End Select
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)
    Select Case address
        Case 65280
            registers.p1 = wbyte And 48
    End Select
End Sub

Public Sub ButtonPressed(ByVal id As Long)
    Me.keys = Me.keys And (Not (Utils.LeftShift(1, id)))
End Sub

Public Sub ButtonReleased(ByVal id As Long)
    Me.keys = Me.keys Or Utils.LeftShift(1, id)
End Sub
