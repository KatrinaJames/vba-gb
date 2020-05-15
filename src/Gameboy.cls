VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gameboy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public clockSpeed As Double
Public clockMultiplier As Double
Public cpuClock As Long
Public frameIntervalDelay As Double
Public frameInterval As Long
Public cartridge As cartridge
Public systemInitialized As Boolean
Public cpu As LR35902
Public MMU As MMU
Public Gpu As Gpu
Public Serial As Serial
Public APU As APU
Public Timer As Timer
Public Joypad As Joypad
Public gbcModeEnabled As Boolean
Public shouldStop As Boolean

Public Sub Class_Initialize()

    Me.clockSpeed = 4194304 ' Hz, double speed for GBC mode.
    Me.clockMultiplier = 1 ' Default to 1x multiplier for DMG clock speed.
    Me.frameIntervalDelay = 16.67 ' 1/60s = ~16ms
    
    Me.gbcModeEnabled = False
    Me.systemInitialized = False
    
    Me.shouldStop = False
    Me.cpuClock = 9999999
End Sub

Public Sub LoadFile(ByVal filepath As String)

    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim tempByte As Byte
    Dim byteArray() As Byte
    
    Open filepath For Binary Access Read As fileNum
    ReDim byteArray(0 To LOF(fileNum) - 1)
    Get fileNum, , byteArray
    Close fileNum
    
    Set Me.cartridge = NewCartridge(byteArray)
    
End Sub

Public Sub Start(Optional ByVal runToInstruction As Long)

    If Me.cartridge Is Nothing Then
        Debug.Print "No cartridge loaded!"
        Exit Sub
    End If
    
    If Not Me.systemInitialized Then
        ' Initialize components.
        Set Me.cpu = Create.NewLR35902(Me)
        Set Me.MMU = Create.NewMMU(Me)
        Set Me.Gpu = Create.NewGpu(Me)
        Set Me.Serial = Create.NewSerial(Me)
        Set Me.APU = Create.NewAPU(Me)
        Set Me.Timer = Create.NewTimer(Me)
        Set Me.Joypad = Create.NewJoypad(Me)
        
        ' Set starting register values.
        If Me.cartridge.colorGameboyFlag Then
            Me.gbcModeEnabled = True
            Me.cpu.SetAF 4480
            Me.cpu.SetBC 0
            Me.cpu.SetDE 65366
            Me.cpu.SetHL 13
        Else
            Me.cpu.SetAF 432
            Me.cpu.SetBC 19
            Me.cpu.SetDE 216
            Me.cpu.SetHL 333
        End If
        
        Me.cpu.RegisterPC = 256
        Me.cpu.RegisterSP = 65534
        
        Me.Gpu.RegisterLCDC = 145
        
        Me.systemInitialized = True
    End If
    
    If Not CBool(Me.frameInterval) Then
        Main.RunFrame Me, runToInstruction
    End If

End Sub

Public Sub Stopp()
    Me.frameInterval = Empty
    Me.shouldStop = True
End Sub

Public Sub Frame(Optional ByVal runToInstruction As Long)
    Const baseFrameCycleLimit As Long = 70224
    Dim frameClockLimit As Long: frameClockLimit = (baseFrameCycleLimit * Me.clockMultiplier)
    
    Dim errorNumber As Long
    Dim errorDescription As String
    
    Do
        If CBool(runToInstruction) And Me.cpu.RegisterPC = runToInstruction Then
            Me.Stopp
            Debug.Print "DEBUG: Stop address reached: " & runToInstruction
            GoTo BREAK
        End If
        
        On Error GoTo CATCH
        Me.Step
        GoTo LOOP_END
CATCH:
        Debug.Print Err.Description
        Me.Stopp
        errorNumber = Err.Number
        errorDescription = Err.Description
        On Error GoTo 0
        Err.Raise errorNumber, "Gameboy", errorDescription
LOOP_END:
    Loop While Me.cpuClock < frameClockLimit
BREAK:

    ' Frame complete, reset CPU clock.
    Me.cpuClock = 0
    
    Dim i As Long
    Dim ram() As Byte
    ReDim ram(0 To Me.cartridge.GetRamSize)
    For i = 0 To UBound(ram)
        ram(i) = Me.cartridge.GetRamByteAt(i)
    Next i
    
    ' Save RAM to local storage if there's a battery in the cartridge.
    If Me.cartridge.hasBattery And Me.cartridge.ramIsDirty Then
        Utils.SetRam "RAM-" & Me.cartridge.title & "-" & CStr(Me.cartridge.globalChecksum), ram
        Me.cartridge.ramIsDirty = False
    End If
End Sub

Public Sub Step()
    Me.CheckInterupts
    Me.cpu.Step
End Sub

Public Sub ConsumeClockCycles(ByVal cycles As Integer)
    Me.cpuClock = Me.cpuClock + cycles
    
    Me.Gpu.Step cycles
    Me.Timer.tick cycles
End Sub

Public Sub RequestInterrupt(ByVal id As Long)
    Me.MMU.WriteByte 65295, Me.MMU.ReadByte(65295) Or Utils.LeftShift(1, id)
End Sub

Public Sub CheckInterupts()
    ' Check if interrupts are enabled.
    If Not Me.cpu.imeFlag Then
        Exit Sub
    End If
    
    If Not CBool(Me.MMU.ReadByte(65535)) Then ' Check if anything is allowed to interrupt.
        Exit Sub
    End If
    
    Dim interupts As Long: interupts = Me.MMU.ReadByte(65295) ' Get active interrupt flags.
    
    If Not CBool(interupts) Then ' Leave if nothing to handle.
        Exit Sub
    End If
    
    Dim i As Long
    For i = 0 To 4
        ' Check if the IE flag is set for the given interrupt.
        If CBool(interupts And Utils.LeftShift(1, i)) And CBool(Me.MMU.ReadByte(65535) And Utils.LeftShift(1, i)) Then
            Me.HandleInterupt i
        End If
    Next i
End Sub

Public Sub HandleInterupt(ByVal interrupt As Long)
    Me.cpu.imeFlag = False ' Disable interrupt handling.
    Me.cpu.haltFlag = False
    
    Me.cpu.RegisterSP = Me.cpu.RegisterSP - 2 ' Push program counter to stack.
    Me.MMU.WriteWord Me.cpu.RegisterSP, Me.cpu.RegisterPC
    
    Dim interupts As Long: interupts = Me.MMU.ReadByte(65295)
    interupts = interupts And Not Utils.LeftShift(1, interrupt) ' Reset interrupt flag.
    Me.MMU.WriteByte 65295, interupts
    
    Select Case interrupt
        Case 0: Me.cpu.RegisterPC = 64 ' V-blank
        Case 1: Me.cpu.RegisterPC = 72 ' LCD
        Case 2: Me.cpu.RegisterPC = 80 ' Timer
        Case 3: Me.cpu.RegisterPC = 88 ' Serial
        Case 4: Me.cpu.RegisterPC = 96 ' Joypad
    End Select
    
    Me.ConsumeClockCycles 20
End Sub
