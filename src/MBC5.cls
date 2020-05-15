VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MBC5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public cartridge As cartridge
Public romBankLo As Byte
Public romBankHi As Byte
Public ramBank As Byte
Public ramEnabled As Boolean

Public Sub Class_Initialize()

    Me.romBankLo = 1  ' Actually goes up to 9-bits of banking.
    Me.romBankHi = 0
    Me.ramBank = 0
    Me.ramEnabled = False

End Sub

Public Function ReadByte(ByVal address As Long) As Byte
    Dim offset As Long

    ' ROM Bank 00
    If address >= 0 And address <= 16383 Then
        ReadByte = Me.cartridge.GetRomByteAt(address)
        Exit Function
    End If
    
    ' ROM Banks 00-1FF
    If address >= 16384 And address <= 32767 Then
        Dim romBank As Byte: romBank = Utils.LeftShift(Me.romBankHi, 8) + Me.romBankLo
        offset = 16384 * romBank
        ReadByte = Me.cartridge.GetRomByteAt((address - 16384) + offset)
        Exit Function
    End If
    
    ' RAM
    If address >= 40960 And address <= 49151 Then
        If Not Me.cartridge.hasRam Then
            Err.Raise 50000, "MBC5", "Cartridge has no RAM but RAM access was attempted."
        End If
        
        offset = 8192 * Me.ramBank
        ReadByte = Me.cartridge.GetRamByteAt((address - 40960) + offset)
        Exit Function
    End If
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)
    Dim offset As Long

    ' RAM Enable
    If address >= 0 And address <= 8191 Then
        Me.ramEnabled = (wbyte And 15) = 10
        Exit Sub
    End If
    
    ' ROM Banking Low
    If address >= 8192 And address <= 12287 Then
        Me.romBankLo = wbyte
        Exit Sub
    End If
    
    ' ROM Banking High
    If address >= 12288 And address <= 16383 Then
        Me.romBankHi = wbyte
        Exit Sub
    End If
    
    ' RAM Banking
    If address >= 16384 And address <= 24575 Then
        Me.ramBank = wbyte And 15
        Exit Sub
    End If
    
    ' RAM
    If address >= 40960 And address <= 49151 Then
        If Not Me.ramEnabled Then ' RAM disabled.
            Exit Sub
        End If
        
        ' Mark for persistance at the end of the next frame.
        If Me.cartridge.hasBattery Then
            Me.cartridge.ramIsDirty = True
        End If
        
        offset = Me.ramBank * 8192
        Me.cartridge.SetRamByteAt (address - 40960) + offset, wbyte
    End If
End Sub
