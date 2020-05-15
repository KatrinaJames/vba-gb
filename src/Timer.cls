VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Public frequency As Long
Public counter As Long
Private registers As TimerRegisters

Public Sub Class_Initialize()

    Me.frequency = 4096
    Me.counter = 0

    registers.div = 0 ' 0xFF04 (r/w) Divider 16-bit MSB is actual value
    registers.tima = 0 ' 0xFF05 (r/w) Timer counter
    registers.tma = 0 ' 0xFF06 (r/w) Timer modulo
    registers.tac = 0 ' 0xFF07 (r/w) Timer control

End Sub

Public Function ReadByte(ByVal address As Long) As Integer

    Select Case address
        Case 65284: ReadByte = Utils.RightShift(registers.div, 8)
        Case 65285: ReadByte = registers.tima
        Case 65286: ReadByte = registers.tma
        Case 65287: ReadByte = registers.tac Or 248
    End Select

End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Integer)

    Select Case address
        Case 65284: registers.div = 0
        Case 65285: registers.tima = wbyte
        Case 65286: registers.tma = wbyte
        Case 65287: Me.UpdateFrequency wbyte
    End Select

End Sub

Public Function IsClockEnabled() As Boolean
    IsClockEnabled = CBool(registers.tac And 4)
End Function

Public Sub UpdateFrequency(ByVal data As Long)

    Dim currentFrequency As Long: currentFrequency = registers.tac And 3
    registers.tac = data
    Dim newFrequency As Long: newFrequency = registers.tac And 3
    
    If currentFrequency <> newFrequency Then
    
        Select Case newFrequency
            Case 0: Me.frequency = 4096
            Case 1: Me.frequency = 262144
            Case 2: Me.frequency = 65536
            Case 3: Me.frequency = 16386
        End Select
    
    End If

End Sub

Public Sub tick(ByVal cycles As Long)

    registers.div = (registers.div + cycles) And 65535
    
    If Not Me.IsClockEnabled() Then
        Exit Sub
    End If
    
    Me.counter = Me.counter + cycles
    Dim interval As Double: interval = Me.system.clockSpeed / Me.frequency
    
    Do While Me.counter >= interval
    
        Me.counter = Me.counter - interval
        
        ' Did timer overflow?
        If registers.tima = 255 Then
            registers.tima = registers.tma
        End If
        
    Loop
    
End Sub
