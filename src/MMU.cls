VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MMU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Private wram() As Byte
Private zram() As Byte

Private registers As MmuRegisters

Public Sub Class_Initialize()

    ReDim wram(0 To 32768 - 1)
    ReDim zram(0 To 128 - 1)
    
    ' Initialize RAM
    Dim i As Long
    For i = 0 To UBound(wram)
        wram(i) = Int(Rnd() * 256) ' Reset Working RAM (32KB)
    Next i
    
    For i = 0 To UBound(zram)
        zram(i) = Int(Rnd() * 256) ' Reset Zero-page RAM (128B)
    Next i
    
    registers.if = 0 ' $FF0F Interrupt Flag (R/W)
    registers.ie = 0 ' $FFFF Interrupt Enable (R/W)
    
    '  Color GB Only Registers
    registers.key1 = 126 ' $FF4D Prepare Speed Switch
    registers.tp = 0 ' $FF56 Infrared Communications Port
    registers.svbk = 0 ' $FF70 WRAM Bank

End Sub

Public Function ReadByte(ByVal address As Long) As Byte

    If address < 0 Or address > 65535 Then
        Err.Raise 50000, "MMU", "Segfault read @ " & address
    End If
    
    ' ROM Banks
    If address >= 0 And address <= 32767 Then
        ReadByte = Me.system.cartridge.ReadByte(address)
        Exit Function
    End If
    
    ' VRAM
    If address >= 32768 And address <= 40959 Then
        ReadByte = Me.system.Gpu.ReadByte(address)
        Exit Function
    End If
    
    ' External RAM
    If address >= 40960 And address <= 49151 Then
        ReadByte = Me.system.cartridge.ReadByte(address)
        Exit Function
    End If
    
    ' WRAM Bank 0
    If address >= 49152 And address <= 53247 Then
        ReadByte = wram(address - 49152)
        Exit Function
    End If
        
    ' WRAM Switchable Banks 1-7
    If address >= 53248 And address <= 57343 Then
        ReadByte = wram((address - 49152) + (registers.svbk * 4096))
        Exit Function
    End If
    
    ' WRAM Echo
    If address >= 57344 And address <= 65023 Then
        ReadByte = wram(address - 57344)
        Exit Function
    End If
     
    ' Sprite Attribute Table (OAM)
    If address >= 65024 And address <= 65183 Then
        ReadByte = Me.system.Gpu.ReadByte(address)
        Exit Function
    End If
     
    ' Joypad
    If address = 65280 Then
        ReadByte = Me.system.Joypad.ReadByte(address)
        Exit Function
    End If
    
    ' Serial
    If address >= 65281 And address <= 65282 Then
        ReadByte = Me.system.Serial.ReadByte(address)
        Exit Function
    End If
    
    ' Timer
    If address >= 65284 And address <= 65287 Then
        ReadByte = Me.system.Timer.ReadByte(address)
        Exit Function
    End If
    
    ' Interrupt Flag
    If address = 65295 Then
        ReadByte = registers.if Or 224
        Exit Function
    End If
    
    ' Audio
    If address >= 65296 And address <= 65343 Then
        ReadByte = Me.system.APU.ReadByte(address)
        Exit Function
    End If
    
    ' GPU
    If address >= 65344 And address <= 65355 Then
        ReadByte = Me.system.Gpu.ReadByte(address)
        Exit Function
    End If
    
    ' GBC Double Speed (GBC Only)
    If address = 65357 Then
        ReadByte = registers.key1
        Exit Function
    End If
    
    ' LCD VRAM Banking (GBC Only)
    If address = 65359 Then
        ReadByte = Me.system.Gpu.ReadByte(address)
        Exit Function
    End If
    
    ' LCD VRAM DMA (GBC Only)
    If address >= 65361 And address <= 65365 Then
        ReadByte = Me.system.Gpu.ReadByte(address)
        Exit Function
    End If
    
    ' WRAM Bank Select (GBC Only)
    If address = 65392 Then
        ReadByte = registers.svbk And 7
        Exit Function
    End If
    
    ' High RAM (stack)
    If address >= 65408 And address <= 65534 Then
        ReadByte = zram(address - 65408)
        Exit Function
    End If
    
    ' Interrupt Enable Register
    If address = 65535 Then
        ReadByte = registers.ie
        Exit Function
    End If
    
    Debug.Print "Warning: Read attempt @ " & address & " / instr: " & Me.system.cpu.instructionCode

End Function

Public Function ReadWord(ByVal address As Long) As Long
    ' Read byte + next byte shifted by 1 byte.
    ReadWord = (Utils.LeftShift(Me.ReadByte(address + 1), 8)) + Me.ReadByte(address)
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)
    If address < 0 Or address > 65535 Then
        Err.Raise 50000, "MMU", "Segfault write @ " & address & " / value: " & wbyte
    End If
    If IsEmpty(wbyte) Then
        Err.Raise 50000, "MMU", "Ins " & Me.system.cpu.instructionCode & " tried to write NaN to " & address
    End If
    
    ' Cartridge ROM
    If address >= 0 And address <= 32767 Then
        Me.system.cartridge.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' VRAM
    If address >= 32768 And address <= 40959 Then
        Me.system.Gpu.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' External RAM
    If address >= 40960 And address <= 49151 Then
        Me.system.cartridge.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' WRAM Bank 0
    If address >= 49152 And address <= 53247 Then
        wram(address - 49152) = wbyte
        Exit Sub
    End If
    
    ' WRAM Switchable Banks 1-7
    If address >= 53248 And address <= 57343 Then
        wram((address - 49152) + (registers.svbk * 4096)) = wbyte
        Exit Sub
    End If
    
    ' WRAM Echo
    If address >= 57344 And address <= 65023 Then
        wram(address - 57344) = wbyte
        Exit Sub
    End If
    
    ' Sprite Attribute Table (OAM)
    If address >= 65024 And address <= 65183 Then
        Me.system.Gpu.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' Unusable space
    If address >= 65184 And address <= 65279 Then
        Exit Sub
    End If
     
    ' Joypad
    If address = 65280 Then
        Me.system.Joypad.WriteByte address, wbyte
        Exit Sub
    End If
     
    ' Serial
    If address >= 65281 And address <= 65282 Then
        Me.system.Serial.WriteByte address, wbyte
        Exit Sub
    End If
     
    ' Timer
    If address >= 65284 And address <= 65287 Then
        Me.system.Timer.WriteByte address, wbyte
        Exit Sub
    End If
     
    ' Interrupt Flag
    If address = 65295 Then
        registers.if = wbyte
        Exit Sub
    End If
    
    ' Audio
    If address >= 65296 And address <= 65343 Then
        Me.system.APU.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' Graphics
    If address >= 65344 And address <= 65355 Then
        Me.system.Gpu.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' GBC Double Speed (GBC Only)
    If address = 65357 Then
        registers.key1 = wbyte
        Exit Sub
    End If
    
    ' LCD VRAM Banking (GBC Only)
    If address = 65359 Then
        Me.system.Gpu.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' LCD VRAM DMA (GBC Only)
    If address >= 65361 And address <= 65365 Then
        Me.system.Gpu.WriteByte address, wbyte
        Exit Sub
    End If
    
    ' WRAM Bank Select (GBC Only)
    If address = 65392 Then
        registers.svbk = wbyte
        If registers.svbk = 0 Then
            registers.svbk = 1
        End If
        Exit Sub
    End If
    
    ' Zero-page RAM
    If address >= 65408 And address <= 65534 Then
        zram(address - 65408) = wbyte
        Exit Sub
    End If
    
    ' Interrupt Enable Register
    If address = 65535 Then
        registers.ie = wbyte
        Exit Sub
    End If
    
    Debug.Print "MMU: Write attempt @ " & address & " value: " & wbyte
End Sub

Public Sub WriteWord(ByVal address As Long, ByVal word As Long)
    Me.WriteByte address, word And 255 ' LSB
    Me.WriteByte address + 1, Utils.RightShift(word, 8)
End Sub
