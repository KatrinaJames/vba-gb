VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MBC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public cartridge As cartridge
Public romBank As Byte
Public ramBank As Byte
Public ramEnabled As Boolean
Public bankMode As Byte

Public Sub Class_Initialize()
    
    Me.romBank = 1
    Me.ramBank = 0
    Me.ramEnabled = False
    Me.bankMode = 0
    
End Sub

Public Function ReadByte(ByVal address As Long) As Byte

    Dim offset As Long

    ' ROM Bank 00
    If address >= 0 And address <= 16383 Then
        ReadByte = Me.cartridge.GetRomByteAt(address)
        Exit Function
    End If
    
    ' ROM Banks 01-7F
    If address >= 16384 And address <= 32767 Then
        offset = CLng(16384) * Me.romBank
        ReadByte = Me.cartridge.GetRomByteAt((address - 16384) + offset)
        Exit Function
    End If
    
    ' RAM
    If address >= 40960 And address <= 49151 Then
        If Not Me.cartridge.hasRam Then
            Err.Raise 50000, "MBC1", "Cartridge has no RAM but RAM access was attempted."
        End If
        
        offset = 8192 * Me.ramBank
        ReadByte = Me.cartridge.GetRamByteAt((address - 40960) + offset)
    End If

End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)

    Dim romBank2 As Byte
    Dim offset As Long

    ' RAM Enable
    If address >= 0 And address <= 8191 Then
        Me.ramEnabled = (wbyte And 15) = 10
        Exit Sub
    End If
    
    ' ROM Banking
    If address >= 8192 And address <= 16383 Then
        romBank2 = wbyte And 31 ' Mask for lower 5 bits.
        Me.romBank = Me.romBank And 224 ' Turn off lower 5 bits.
        Me.romBank = Me.romBank Or romBank2 ' Set lower 5 bits.
        If Me.romBank = 0 Then
            Me.romBank = Me.romBank + 1
        End If
        Exit Sub
    End If
    
    ' RAM/ROM Banking
    If address >= 16384 And address <= 24575 Then
        If Me.bankMode = 0 Then  ' ROM Banking
            romBank2 = Utils.LeftShift(wbyte, 5) ' Move bits into correct location.
            Me.romBank = Me.romBank And 96 ' Turn off bits 5 and 6.
            Me.romBank = Me.romBank Or romBank2 ' Set bits 5 and 6.
        ElseIf Me.bankMode = 1 Then ' RAM Banking
            Me.ramBank = wbyte And 3
        End If
        Exit Sub
    End If
    
    ' ROM/RAM Mode Select
    If address >= 24576 And address <= 32767 Then
        Me.bankMode = wbyte And 1
        Exit Sub
    End If
    
    If address >= 40960 And address <= 49151 Then
        If Not Me.ramEnabled Then ' RAM disabled.
            Exit Sub
        End If
        
        ' Mark for persistance at the end of the next frame.
        If Me.cartridge.hasBattery Then
            Me.cartridge.ramIsDirty = True
        End If
        
        If Me.bankMode = 0 Then ' ROM mode, only write to bank 0x00 of RAM.
            Me.cartridge.SetRamByteAt address - 40960, wbyte
            Exit Sub
        End If
        
        offset = Me.ramBank * 8192
        Me.cartridge.SetRamByteAt (address - 40960) + offset, wbyte
        
    End If
    
End Sub
