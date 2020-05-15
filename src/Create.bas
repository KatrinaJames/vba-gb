Attribute VB_Name = "Create"
' Constructors

Option Explicit
Option Base 0

Public Function NewCartridge(ByRef rom() As Byte) As cartridge

    Dim obj As New cartridge
    Dim lowerBound As Long: lowerBound = LBound(rom)
    Dim upperBound As Long: upperBound = UBound(rom)
    obj.RedimRom lowerBound, upperBound
    
    Dim i As Long: For i = lowerBound To upperBound
        obj.SetRomByteAt i, rom(i)
    Next i
    
    obj.init
    Set NewCartridge = obj
    
End Function

Public Function NewLR35902(ByRef system As Gameboy) As LR35902

    Dim obj As New LR35902
    Set obj.system = system
    Set NewLR35902 = obj

End Function

Public Function NewMMU(ByRef system As Gameboy) As MMU

    Dim obj As New MMU
    Set obj.system = system
    Set NewMMU = obj

End Function

Public Function NewGpu(ByRef system As Gameboy) As Gpu

    Dim obj As New Gpu
    Set obj.system = system
    Set NewGpu = obj

End Function

Public Function NewSerial(ByRef system As Gameboy) As Serial

    Dim obj As New Serial
    Set obj.system = system
    Set NewSerial = obj

End Function

Public Function NewAPU(ByRef system As Gameboy) As APU

    Dim obj As New APU
    Set obj.system = system
    Set NewAPU = obj

End Function

Public Function NewTimer(ByRef system As Gameboy) As Timer

    Dim obj As New Timer
    Set obj.system = system
    Set NewTimer = obj

End Function

Public Function NewMBC1(ByRef cart As cartridge) As MBC1

    Dim obj As New MBC1
    Set obj.cartridge = cart
    Set NewMBC1 = obj

End Function

Public Function NewMBC2(ByRef cart As cartridge) As MBC2

    Dim obj As New MBC2
    Set obj.cartridge = cart
    Set NewMBC2 = obj

End Function

Public Function NewMBC3(ByRef cart As cartridge) As MBC3

    Dim obj As New MBC3
    Set obj.cartridge = cart
    Set NewMBC3 = obj

End Function

Public Function NewMBC5(ByRef cart As cartridge) As MBC5

    Dim obj As New MBC5
    Set obj.cartridge = cart
    Set NewMBC5 = obj

End Function

Public Function NewJoypad(ByRef system As Gameboy) As Joypad

    Dim obj As New Joypad
    Set obj.system = system
    Set NewJoypad = obj

End Function
