VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Serial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Private register As SerialRegister

Public Sub Class_Initialize()

    register.sb = 0 ' 0xFF01 (r/w) Serial transfer data
    register.sc = 0 ' 0xFF02 (r/w) Serial I/O control

End Sub

Public Function ReadByte(ByVal address As Long) As Long
    Select Case address
        Case 65281: ReadByte = register.sb
        Case 65282: ReadByte = register.sc
    End Select
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Long)
    Select Case address
        Case 65281
            register.sb = wbyte
        Case 65282
            register.sc = wbyte
    End Select
End Sub

