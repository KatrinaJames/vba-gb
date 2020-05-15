VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cartridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private rom() As Byte
Public headerChecksum As Byte
Public globalChecksum As Long
Public title As String
Public colorGameboyFlag As Boolean
Public superGameboyFlag As Boolean
Public cartridgeType As Byte
Public mbc As Object
Public hasRam As Boolean
Public hasBattery As Boolean
Public hasRumble As Boolean
Public romSize As Byte
Public totalRomBanks As Integer
Public totalRamBanks As Integer
Public RamSize As Byte
Private ram() As Byte
Public ramIsDirty As Boolean

Public Sub Class_Initialize()

End Sub

Public Sub RedimRom(ByVal lowerBound As Long, ByVal upperBound As Long)
    ReDim rom(lowerBound To upperBound)
End Sub

Public Function GetRomByteAt(ByVal index As Long) As Byte
    GetRomByteAt = rom(index)
End Function

Public Sub SetRomByteAt(ByVal index As Long, ByVal value As Byte)
    rom(index) = value
End Sub

Public Function GetRamByteAt(ByVal index As Long) As Byte
    GetRamByteAt = ram(index)
End Function

Public Sub SetRamByteAt(ByVal index As Long, ByVal value As Byte)
    ram(index) = value
End Sub

Public Function GetRamSize() As Long
    GetRamSize = UBound(ram)
End Function

Public Sub init()

    ' Read header checksum.
    Me.headerChecksum = rom(333)
    
    ' Read global checksum.
    Me.globalChecksum = Utils.LeftShift(rom(334), 8) + rom(335)
    
    ' Read ROM title.
    Me.title = ""
    
    Dim i As Long
    For i = 308 To 322
        If Not rom(i) = 0 Then
            Me.title = Me.title & Chr(rom(i))
        End If
    Next i
    
    ' Read Color GB flag.
    Me.colorGameboyFlag = rom(323) <> 0
    
    ' Read Super GB flag.
    Me.superGameboyFlag = rom(326) = 3
    
    ' Read cartridge type, determine memory bank controller and other cartridge properties.
    Me.cartridgeType = rom(327)
    Set Me.mbc = Nothing
    Me.hasRam = False
    Me.hasBattery = False
    Select Case Me.cartridgeType
        Case 1: Set Me.mbc = Create.NewMBC1(Me)
        Case 2: Set Me.mbc = Create.NewMBC1(Me): Me.hasRam = True
        Case 3: Set Me.mbc = Create.NewMBC1(Me): Me.hasRam = True: Me.hasBattery = True
        Case 5: Set Me.mbc = Create.NewMBC2(Me)
        Case 6: Set Me.mbc = Create.NewMBC2(Me): Me.hasBattery = True
        Case 8: Me.hasRam = True
        Case 9: Me.hasRam = True: Me.hasBattery = True
        Case 15: Set Me.mbc = Create.NewMBC3(Me): Me.hasBattery = True
        Case 16: Set Me.mbc = Create.NewMBC3(Me): Me.hasBattery = True
        Case 17: Set Me.mbc = Create.NewMBC3(Me)
        Case 18: Set Me.mbc = Create.NewMBC3(Me): Me.hasRam = True
        Case 19: Set Me.mbc = Create.NewMBC3(Me): Me.hasRam = True: Me.hasBattery = True
        Case 25: Set Me.mbc = Create.NewMBC5(Me)
        Case 26: Set Me.mbc = Create.NewMBC5(Me): Me.hasRam = True
        Case 27: Set Me.mbc = Create.NewMBC5(Me): Me.hasRam = True: Me.hasBattery = True
        Case 28: Set Me.mbc = Create.NewMBC5(Me): Me.hasRumble = True
        Case 29: Set Me.mbc = Create.NewMBC5(Me): Me.hasRumble = True: Me.hasRam = True
        Case 30: Set Me.mbc = Create.NewMBC5(Me): Me.hasRumble = True: Me.hasRam = True: Me.hasBattery = True
        Case Else: Err.Raise 50000, "Cartridge", "Cartridge: Unsupported cartridge type: " & Me.cartridgeType
    End Select
    
    ' Read ROM size.
    Me.romSize = rom(328)
    
    ' Determine total ROM banks.
    Select Case Me.romSize
        Case 0: Me.totalRomBanks = 1
        Case 1: Me.totalRomBanks = 4
        Case 2: Me.totalRomBanks = 8
        Case 3: Me.totalRomBanks = 16
        Case 4: Me.totalRomBanks = 32
        Case 5: Me.totalRomBanks = 64
        Case 6: Me.totalRomBanks = 128
        Case 7: Me.totalRomBanks = 256
        Case 8: Me.totalRomBanks = 512
        Case 82: Me.totalRomBanks = 72
        Case 83: Me.totalRomBanks = 80
        Case 84: Me.totalRomBanks = 96
    End Select
    
    ' Read RAM size.
    Me.RamSize = rom(329)
    
    ' Initialize RAM space.
    Dim totalRam As Long: totalRam = 0
    Select Case Me.RamSize
        Case 0: totalRam = 0: Me.totalRamBanks = 0
        Case 1: totalRam = 2048: Me.totalRamBanks = 1
        Case 2: totalRam = 8192: Me.totalRamBanks = 1
        Case 3: totalRam = 32768: Me.totalRamBanks = 4
        Case 4: totalRam = 131072: Me.totalRamBanks = 16
        Case 5: totalRam = 65536: Me.totalRamBanks = 8
    End Select
    
    ReDim ram(0 To totalRam - 1)
    If totalRam > 0 Then
        For i = 0 To totalRam - 1
            ram(i) = Int(Rnd() * 256)
        Next i
    End If
    
    ' Load "battery-backed" RAM for storage.
    Dim tempRam() As Byte
    Dim RamLoader As RamLoader
    
    If Me.hasBattery Then
        Me.ramIsDirty = False
        Set RamLoader = Utils.GetRam("RAM-" & Me.title & "-" & CStr(Me.globalChecksum))
        If RamLoader.wasSuccessful Then
        
            ReDim tempRam(0 To RamLoader.RamSize)
            For i = 0 To UBound(tempRam)
                tempRam(i) = RamLoader.GetRamByteAt(i)
            Next i
            
            ReDim ram(0 To UBound(tempRam))
            For i = 0 To UBound(ram)
                ram(i) = tempRam(i)
            Next i
            
        End If
    End If
    
End Sub

Public Function ReadByte(ByVal address As Long) As Byte
    ' ROM
    If address >= 0 And address <= 32767 Then
        If Me.mbc Is Nothing Then
            ReadByte = rom(address)
            Exit Function
        End If
        ReadByte = Me.mbc.ReadByte(address)
        Exit Function
    End If
    
    ' RAM
    If address >= 40960 And address <= 49151 Then
        If Not Me.hasRam Then
            ReadByte = 255
            Exit Function
        End If
        ReadByte = Me.mbc.ReadByte(address)
        Exit Function
    End If
    
    Err.Raise 50000, "Cartridge", "Cartridge: Unsupported read at " & address
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)
    ' ROM
    If address >= 0 And address <= 32767 Then
        If Me.mbc Is Nothing Then ' ROM is read-only
            Exit Sub
        End If
        Me.mbc.WriteByte address, wbyte
        Exit Sub
    End If

    ' RAM
    If address >= 40960 And address <= 49151 Then
        If Not Me.hasRam Then '  No RAM to write to.
            Exit Sub
        End If
        
        Me.mbc.WriteByte address, wbyte
        Exit Sub
    End If
End Sub
