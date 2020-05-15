VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MBC3"
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
Public rtcEnabled As Boolean
Public mappedRegister As Byte
Public latchBuffer As Byte
Public latchedTime As Byte
Private rtcRegister As Mbc3Register
Public rtcEpoch As Long

Public Sub Class_Initialize()

    rtcRegister.seconds = 0
    rtcRegister.minutes = 0
    rtcRegister.hours = 0
    rtcRegister.days = 0
    rtcRegister.dayCarry = 0
    
    Me.rtcEpoch = DateDiff("s", #1/1/1970#, Now())

End Sub

Public Function ReadByte(ByVal address As Long) As Byte

    Dim offset As Long

    ' ROM Bank 0
    If address >= 0 And address <= 16383 Then
        ReadByte = Me.cartridge.GetRomByteAt(address)
        Exit Function
    End If
    
    ' ROM Bank 1 (Memory Bank Controlled)
    If address >= 16384 And address <= 32767 Then
        offset = CLng(16384) * Me.romBank
        ReadByte = Me.cartridge.GetRomByteAt((address - 16384) + offset)
        Exit Function
    End If
    
    ' RAM
    If address >= 40960 And address <= 49151 Then
        If Me.mappedRegister >= 0 And Me.mappedRegister <= 7 Then
            offset = 8192 * Me.ramBank
            ReadByte = Me.cartridge.GetRamByteAt((address - 40960) + offset)
            Exit Function
        End If
    End If
    
    ReadByte = 255

End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)

    Dim romBank2 As Byte
    Dim TimeSpan As TimeSpan
    Dim offset As Long

    ' RAM & RTC Enable
    If address >= 0 And address <= 8191 Then
        Me.ramEnabled = wbyte = 10
        Me.rtcEnabled = wbyte = 10
        Exit Sub
    End If
    
    ' ROM Banking
    If address >= 8192 And address <= 16383 Then
        romBank2 = wbyte And 127 ' Mask for lower 7 bits.
        Me.romBank = Me.romBank And 128 ' Turn off lower 7 bits.
        Me.romBank = Me.romBank Or romBank2 ' Set lower 7 bits.
        If Me.romBank = 0 Then
            Me.romBank = Me.romBank + 1
        End If
        Exit Sub
    End If
    
    ' RAM Bank & RTC Register Select
    If address >= 16384 And address <= 24575 Then
        Me.mappedRegister = wbyte
        Exit Sub
    End If
    
    ' Latch Clock Data
    If address >= 24576 And address <= 32767 Then
        ' Latch RTC if latch buffer is 0x00 and incoming byte is 0x01;
        If Me.latchBuffer = 0 And wbyte = 1 Then
            TimeSpan = Me.CalculateTimeSpan(Me.rtcEpoch, DateDiff("s", #1/1/1970#, Now()))
            
            If TimeSpan.days > 512 Then
                rtcRegister.dayCarry = 1
                TimeSpan.days = TimeSpan.days Mod 512
            End If
            
            rtcRegister.seconds = TimeSpan.seconds
            rtcRegister.minutes = TimeSpan.minutes
            rtcRegister.hours = TimeSpan.hours
            rtcRegister.days = TimeSpan.days
        End If
        
        Me.latchBuffer = wbyte
        Exit Sub
    End If
    
    If address >= 40960 And address <= 49151 Then
        If Me.mappedRegister >= 0 And Me.mappedRegister <= 7 Then
            If Not Me.ramEnabled Then ' RAM disabled.
                Exit Sub
            End If
            
            ' Mark for persistance at the end of the next frame.
            If Me.cartridge.hasBattery Then
                Me.cartridge.ramIsDirty = True
            End If
            
            offset = Me.ramBank * 8192
            Me.cartridge.SetRamByteAt (address - 40960) + offset, wbyte
            Exit Sub
        End If
        
        If Me.mappedRegister >= 8 And Me.mappedRegister <= 12 Then
            Select Case Me.mappedRegister
                Case 8: rtcRegister.seconds = wbyte
                Case 9: rtcRegister.minutes = wbyte
                Case 10: rtcRegister.hours = wbyte
                Case 11: rtcRegister.days = wbyte
                Case 12: Err.Raise 50000, "MBC3", "day high, carry, and halt"
            End Select
            Exit Sub
        End If
    End If

End Sub

Public Function CalculateTimeSpan(ByVal startTime As Long, ByVal endTime As Long) As TimeSpan

    Dim seconds As Long: seconds = Int((endTime - startTime) / 1000)
    Dim minutes As Long: minutes = 0
    Dim hours As Long: hours = 0
    Dim days As Long: days = 0
    
    If seconds >= 60 Then
        minutes = Int(seconds / 60)
        seconds = seconds Mod 60
    End If
    
     If minutes >= 60 Then
        hours = Int(minutes / 60)
        minutes = minutes Mod 60
     End If
     
     If hours >= 24 Then
        days = Int(hours / 24)
        hours = hours Mod 24
     End If
     
    Dim returnVal As TimeSpan

    returnVal.seconds = seconds
    returnVal.minutes = minutes
    returnVal.hours = hours
    returnVal.days = days
    
    CalculateTimeSpan = returnVal

End Function
