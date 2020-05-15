VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Private waveram() As Integer
Private registers As ApuRegisters

Public Sub Class_Initialize()

    ReDim waveram(0 To 31) ' FF30 - FF3F Wave pattern RAM, 32 4-bit samples
    
    Dim i As Long
    
    ' Randomize wave RAM.
    For i = 0 To UBound(waveram)
        waveram(i) = Int(Rnd() * 256) ' Reset wave RAM (32B)
    Next i
    
    ' Channel 1 - Tone & Sweep
    registers.nr10 = 0 ' FF10 Sweep
    registers.nr11 = 0 ' FF11 Sound length / Wave pattern duty
    registers.nr12 = 0 ' FF12 Volume envelope
    registers.nr13 = 0 ' FF13 Frequency lo
    registers.nr14 = 0 ' FF14 Frequency hi
    
    ' Channel 2 - Tone
    registers.nr21 = 0 ' FF16 Sound length / Wave pattern duty
    registers.nr22 = 0 ' FF17 Volume envelope
    registers.nr23 = 0 ' FF18 Frequency lo
    registers.nr24 = 0 ' FF19 Frequency hi
    
    ' Channel 3 - Wave Output
    registers.nr30 = 0 ' FF1A Sound on/off
    registers.nr31 = 0 ' FF1B Sound length
    registers.nr32 = 0 ' FF1C Select output level
    registers.nr33 = 0 ' FF1D Frequency lo
    registers.nr34 = 0 ' FF1E Frequency hi
    
    ' Channel 4
    registers.nr41 = 0 ' FF20 Sound length
    registers.nr42 = 0 ' FF21 Volume envelope
    registers.nr43 = 0 ' FF22 Polynomial counter
    registers.nr44 = 0 ' FF23 Counter / consecutive; Initial
    
    ' Control
    registers.nr50 = 0 ' FF24 Control / ON-OFF / Volume
    registers.nr51 = 0 ' FF25 Sound output terminal selection
    registers.nr52 = 0 ' FF26 On / Off
    
End Sub

Public Function ReadByte(ByVal address As Long) As Byte

    If address >= 65328 And address <= 65343 Then
        ReadByte = waveram(address - 65328)
        Exit Function
    End If
    
    Select Case address
        Case 65296: ReadByte = registers.nr10
        Case 65297: ReadByte = registers.nr11
        Case 65298: ReadByte = registers.nr12
        Case 65299: ReadByte = registers.nr13
        Case 65300: ReadByte = registers.nr14
        Case 65302: ReadByte = registers.nr21
        Case 65303: ReadByte = registers.nr22
        Case 65304: ReadByte = registers.nr23
        Case 65305: ReadByte = registers.nr24
        Case 65306: ReadByte = registers.nr30
        Case 65307: ReadByte = registers.nr31
        Case 65308: ReadByte = registers.nr32
        Case 65309: ReadByte = registers.nr33
        Case 65310: ReadByte = registers.nr34
        Case 65312: ReadByte = registers.nr41
        Case 65313: ReadByte = registers.nr42
        Case 65314: ReadByte = registers.nr43
        Case 65315: ReadByte = registers.nr44
        Case 65316: ReadByte = registers.nr50
        Case 65317: ReadByte = registers.nr51
        Case 65318: ReadByte = registers.nr52
    End Select
    
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)

    If address >= 65328 And address <= 65343 Then
        waveram(address - 65328) = wbyte
        Exit Sub
    End If
    
    Select Case address
        Case 65296: registers.nr10 = wbyte
        Case 65297: registers.nr11 = wbyte
        Case 65298: registers.nr12 = wbyte
        Case 65299: registers.nr13 = wbyte
        Case 65300: registers.nr14 = wbyte
        Case 65302: registers.nr21 = wbyte
        Case 65303: registers.nr22 = wbyte
        Case 65304: registers.nr23 = wbyte
        Case 65305: registers.nr24 = wbyte
        Case 65306: registers.nr30 = wbyte
        Case 65307: registers.nr31 = wbyte
        Case 65308: registers.nr32 = wbyte
        Case 65309: registers.nr33 = wbyte
        Case 65310: registers.nr34 = wbyte
        Case 65312: registers.nr41 = wbyte
        Case 65313: registers.nr42 = wbyte
        Case 65314: registers.nr43 = wbyte
        Case 65315: registers.nr44 = wbyte
        Case 65316: registers.nr50 = wbyte
        Case 65317: registers.nr51 = wbyte
        Case 65318: registers.nr52 = wbyte
    End Select

End Sub

