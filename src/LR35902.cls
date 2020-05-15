VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LR35902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Private register As CpuRegister
Public haltFlag As Boolean
Public stopFlag As Boolean
Public imeFlag As Boolean
Public pendingDisableInterrupts As Integer
Public pendingEnableInterrupts As Integer
Public instructionCode As Long

Public Sub Class_Initialize()

    ' 8-bit registers
    register.a = 0
    register.f = 0
    register.b = 0
    register.c = 0
    register.d = 0
    register.e = 0
    register.h = 0
    register.l = 0
    
    ' 16-bit registers
    register.pc = 0
    register.sp = 0
    
    Me.haltFlag = False
    Me.stopFlag = False
    Me.imeFlag = False
    
    Me.pendingDisableInterrupts = 0
    Me.pendingEnableInterrupts = 0
    
End Sub

Public Property Get RegisterA() As Long
    RegisterA = register.a
End Property

Public Property Get RegisterB() As Long
    RegisterB = register.b
End Property

Public Property Get RegisterC() As Long
    RegisterC = register.c
End Property

Public Property Get RegisterD() As Long
    RegisterD = register.d
End Property

Public Property Get RegisterE() As Long
    RegisterE = register.e
End Property

Public Property Get RegisterF() As Long
    RegisterF = register.f
End Property

Public Property Get RegisterH() As Long
    RegisterH = register.h
End Property

Public Property Get RegisterL() As Long
    RegisterL = register.l
End Property

Public Property Get RegisterSP() As Long
    RegisterSP = register.sp
End Property

Public Property Get RegisterPC() As Long
    RegisterPC = register.pc
End Property

Public Property Let RegisterA(ByVal value As Long)
    register.a = value
End Property

Public Property Let RegisterB(ByVal value As Long)
    register.b = value
End Property

Public Property Let RegisterC(ByVal value As Long)
    register.c = value
End Property

Public Property Let RegisterD(ByVal value As Long)
    register.d = value
End Property

Public Property Let RegisterE(ByVal value As Long)
    register.e = value
End Property

Public Property Let RegisterF(ByVal value As Long)
    register.f = value
End Property

Public Property Let RegisterH(ByVal value As Long)
    register.h = value
End Property

Public Property Let RegisterL(ByVal value As Long)
    register.l = value
End Property

Public Property Let RegisterSP(ByVal value As Long)
    register.sp = value
End Property

Public Property Let RegisterPC(ByVal value As Long)
    register.pc = value
End Property

Public Sub Step()
        
    If register.pc < 0 Or register.pc > 65535 Then
        Err.Raise 50000, "LR35902 Cpu", "Program counter out of range."
    End If
    
    If Me.haltFlag Then
        ' CPU is "powered down". Only wake up if there's an interrupt.
        Me.system.ConsumeClockCycles 4
    Else
        Dim programCounter As Long: programCounter = Me.RegisterPC
        Me.RegisterPC = Me.RegisterPC + 1
        Me.instructionCode = Me.system.MMU.ReadByte(programCounter)
        Select Case instructionCode
            '*****************************************************************
            ' Main Instructions
            '*****************************************************************
            Case 143, 136, 137, 138, 139, 140, 141, 142, 206
                ADC
            Case 135, 128, 129, 130, 131, 132, 133, 134, 198
                ADD
            Case 9, 25, 41, 57
                ADD_HL_n
            Case 232
                ADD_SP_n
            Case 167, 160, 161, 162, 163, 164, 165, 166, 230
                ANDD
            Case 196, 204, 212, 220
                CALL_cc_nn
            Case 205
                CALL_nn
            Case 63
                CCF
            Case 191, 184, 185, 186, 187, 188, 189, 190, 254
                CP
            Case 47
                CPL
            Case 39
                DAA
            Case 61, 5, 13, 21, 29, 37, 45, 53
                DEC
            Case 11, 27, 43, 59
                DEC_nn
            Case 243
                DI
            Case 251
                EI
            Case 118
                HALT
            Case 60, 4, 12, 20, 28, 36, 44, 52
                INC
            Case 3, 19, 35, 51
                INC_nn
            Case 194, 202, 210, 218
                JP_cc_nn
            Case 195
                JP_nn
            Case 233
                JP_HLmem
            Case 32, 40, 48, 56
                JR_cc_n
            Case 24
                JR_n
            Case 8
                LD_d16_SP
            Case 120, 121, 122, 123, 124, 125, 10, 26, 126, 62, 250
                LD_A_n
            Case 64, 65, 66, 67, 68, 69, 70
                LD_B_n
            Case 72, 73, 74, 75, 76, 77, 78
                LD_C_n
            Case 80, 81, 82, 83, 84, 85, 86
                LD_D_n
            Case 88, 89, 90, 91, 92, 93, 94
                LD_E_n
            Case 96, 97, 98, 99, 100, 101, 102
                LD_H_n
            Case 104, 105, 106, 107, 108, 109, 110
                LD_L_n
            Case 248
                LD_HL_SP_r8
            Case 112, 113, 114, 115, 116, 117, 54
                LD_HLmem_n
            Case 127, 71, 79, 87, 95, 103, 111, 2, 18, 119, 234
                LD_n_A
            Case 1, 17, 33, 49
                LD_n_nn
            Case 6, 14, 22, 30, 38, 46
                LD_nn_n
            Case 249
                LD_SP_HL
            Case 242
                LD_A_Cmem
            Case 226
                LD_Cmem_A
            Case 58
                LDD_A_HLmem
            Case 50
                LDD_HLmem_A
            Case 240
                LDH_A_d8mem
            Case 224
                LDH_d8mem_A
            Case 42
                LDI_A_HLmem
            Case 34
                LDI_HLmem_A
            Case 0
                NOP
            Case 183, 176, 177, 178, 179, 180, 181, 182, 246
                ORR
            Case 241, 193, 209, 225
                POP
            Case 245, 197, 213, 229
                PUSH
            Case 201
                RET
            Case 192, 200, 208, 216
                RET_cc
            Case 217
                RETI
            Case 23
                RLA
            Case 7
                RLCA
            Case 31
                RRA
            Case 15
                RRCA
            Case 199, 207, 215, 223, 231, 239, 247, 255
                RST
            Case 152, 153, 154, 155, 156, 157, 158, 159, 222
                SBC
            Case 55
                SCF
            Case 16
                Stopp
            Case 151, 144, 145, 146, 147, 148, 149, 150, 214
                SUBB
            Case 175, 168, 169, 170, 171, 172, 173, 174, 238
                XORR
            ' *****************************************************************
            '  CB Instructions
            ' *****************************************************************
            Case 203
                CB
            Case 51968, 51969, 51970, 51971, 51972, 51973, 51974, 51975
                RLC
            Case 51976, 51977, 51978, 51979, 51980, 51981, 51982, 51983
                RRC
            Case 51984, 51985, 51986, 51987, 51988, 51989, 51990, 51991
                RL
            Case 51992, 51993, 51994, 51995, 51996, 51997, 51998, 51999
                RR
            Case 52000, 52001, 52002, 52003, 52004, 52005, 52006, 52007
                SLA
            Case 52008, 52009, 52010, 52011, 52012, 52013, 52014, 52015
                SRA
            Case 52016, 52017, 52018, 52019, 52020, 52021, 52022, 52023
                SWAP
            Case 52024, 52025, 52026, 52027, 52028, 52029, 52030, 52031
                SRL
            Case 52032, 52033, 52034, 52035, 52036, 52037, 52038, 52039, _
                 52040, 52041, 52042, 52043, 52044, 52045, 52046, 52047, _
                 52048, 52049, 52050, 52051, 52052, 52053, 52054, 52055, _
                 52056, 52057, 52058, 52059, 52060, 52061, 52062, 52063, _
                 52064, 52065, 52066, 52067, 52068, 52069, 52070, 52071, _
                 52072, 52073, 52074, 52075, 52076, 52077, 52078, 52079, _
                 52080, 52081, 52082, 52083, 52084, 52085, 52086, 52087, _
                 52088, 52089, 52090, 52091, 52092, 52093, 52094, 52095
                bit
            Case 52096, 52097, 52098, 52099, 52100, 52101, 52102, 52103, _
                 52104, 52105, 52106, 52107, 52108, 52109, 52110, 52111, _
                 52112, 52113, 52114, 52115, 52116, 52117, 52118, 52119, _
                 52120, 52121, 52122, 52123, 52124, 52125, 52126, 52127, _
                 52128, 52129, 52130, 52131, 52132, 52133, 52134, 52135, _
                 52136, 52137, 52138, 52139, 52140, 52141, 52142, 52143, _
                 52144, 52145, 52146, 52147, 52148, 52149, 52150, 52151, _
                 52152, 52153, 52154, 52155, 52156, 52157, 52158, 52159
                RES
            Case 52160, 52161, 52162, 52163, 52164, 52165, 52166, 52167, _
                 52168, 52169, 52170, 52171, 52172, 52173, 52174, 52175, _
                 52176, 52177, 52178, 52179, 52180, 52181, 52182, 52183, _
                 52184, 52185, 52186, 52187, 52188, 52189, 52190, 52191, _
                 52192, 52193, 52194, 52195, 52196, 52197, 52198, 52199, _
                 52200, 52201, 52202, 52203, 52204, 52205, 52206, 52207, _
                 52208, 52209, 52210, 52211, 52212, 52213, 52214, 52215, _
                 52216, 52217, 52218, 52219, 52220, 52221, 52222, 52223
                SETT
            Case Else
                Err.Raise 50000, "LR35902", "Instruction " & Me.instructionCode & " not found"
        End Select
    End If
    
    If CBool(Me.pendingEnableInterrupts) Then
        If (Me.pendingEnableInterrupts And 15) > 0 Then
            Me.pendingEnableInterrupts = Me.pendingEnableInterrupts - 1
        Else
            Me.imeFlag = True
            Me.pendingEnableInterrupts = 0
        End If
    End If
    
End Sub

' Flag Helpers
Public Sub SetZ()
    register.f = register.f Or 128
End Sub
Public Sub SetN()
    register.f = register.f Or 64
End Sub
Public Sub SetH()
    register.f = register.f Or 32
End Sub
Public Sub SetC()
    register.f = register.f Or 16
End Sub
Public Sub ClearZ()
    register.f = register.f And Not 128
End Sub
Public Sub ClearN()
    register.f = register.f And Not 64
End Sub
Public Sub ClearH()
    register.f = register.f And Not 32
End Sub
Public Sub ClearC()
    register.f = register.f And Not 16
End Sub

' Register Helpers
Public Function GetAF() As Long
    GetAF = Utils.LeftShift(register.a, 8) + register.f
End Function
Public Function GetBC() As Long
    GetBC = Utils.LeftShift(register.b, 8) + register.c
End Function
Public Function GetDE() As Long
    GetDE = Utils.LeftShift(register.d, 8) + register.e
End Function
Public Function GetHL() As Long
    GetHL = Utils.LeftShift(register.h, 8) + register.l
End Function
Public Sub SetAF(ByVal value As Long)
    register.a = Utils.RightShift(value, 8) And 255
    register.f = value And 240
End Sub
Public Sub SetBC(ByVal value As Long)
    register.b = Utils.RightShift(value, 8) And 255
    register.c = value And 255
End Sub
Public Sub SetDE(ByVal value As Long)
    register.d = Utils.RightShift(value, 8) And 255
    register.e = value And 255
End Sub
Public Sub SetHL(ByVal value As Long)
    register.h = Utils.RightShift(value, 8) And 255
    register.l = value And 255
End Sub

' *************************************************************************
' Main Instructions
' *************************************************************************

Public Sub ADC()
    Dim cycles As Long
    Dim inputt As Long
    
    Select Case Me.instructionCode
        Case 143
            inputt = register.a: cycles = 4
        Case 136
            inputt = register.b: cycles = 4
        Case 137
            inputt = register.c: cycles = 4
        Case 138
            inputt = register.d: cycles = 4
        Case 139
            inputt = register.e: cycles = 4
        Case 140
            inputt = register.h: cycles = 4
        Case 141
            inputt = register.l: cycles = 4
        Case 142
            inputt = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 206
            inputt = Me.system.MMU.ReadByte(register.pc): cycles = 8
            register.pc = register.pc + 1
    End Select
    
    Dim a As Long: a = register.a
    Dim carry As Long: carry = IIf(CBool(register.f And 16), 1, 0)
    Dim result As Long: result = a + inputt + carry
    Me.ClearN
    If result > 255 Then
        Me.SetC
    Else
        Me.ClearC
    End If
    result = result And 255
    If result = 0 Then
        Me.SetZ
    Else
        Me.ClearZ
    End If
    If CBool((a Xor inputt Xor result) And 16) Then
        Me.SetH
    Else
        Me.ClearH
    End If
    register.a = result
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub ADD()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 135: value = register.a: cycles = 4
        Case 128: value = register.b: cycles = 4
        Case 129: value = register.c: cycles = 4
        Case 130: value = register.d: cycles = 4
        Case 131: value = register.e: cycles = 4
        Case 132: value = register.h: cycles = 4
        Case 133: value = register.l: cycles = 4
        Case 134: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 198
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    Dim a As Long: a = register.a
    register.a = register.a + value
    If register.a > 255 Then
        Me.SetC
    Else
        Me.ClearC
    End If
    register.a = register.a And 255
    Me.ClearN
    If register.a = 0 Then
        Me.SetZ
    Else
        Me.ClearZ
    End If
    If ((a And 15) + (value And 15)) > 15 Then
        Me.SetH
    Else
        Me.ClearH
    End If
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub ADD_HL_n()
    Dim value As Long
    
    Select Case Me.instructionCode
        Case 9: value = Me.GetBC()
        Case 25: value = Me.GetDE()
        Case 41: value = Me.GetHL()
        Case 57: value = register.sp
    End Select
    
    Dim hl As Long: hl = Me.GetHL()
    Me.ClearN
    If CBool(((hl And 4095) + (value And 4095)) And 4096) Then
        Me.SetH
    Else
        Me.ClearH
    End If
    If hl + value > 65535 Then
        Me.SetC
    Else
        Me.ClearC
    End If
    hl = hl + value
    Me.SetHL hl
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub ADD_SP_n()
    Dim n As Long: n = Me.system.MMU.ReadByte(register.pc)
    register.pc = register.pc + 1
    If n > 127 Then
        n = -(((Not n) + 1) And 255)
    End If
    Dim result As Long: result = register.sp + n
    Me.ClearZ
    Me.ClearN
    If (register.sp Xor n Xor result) And 16 Then
        Me.SetH
    Else
        Me.ClearH
    End If
    If (register.sp Xor n Xor result) And 256 Then
        Me.SetC
    Else
        Me.ClearC
    End If
    register.sp = result And 65535
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub ANDD()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 167: value = register.a: cycles = 4
        Case 160: value = register.b: cycles = 4
        Case 161: value = register.c: cycles = 4
        Case 162: value = register.d: cycles = 4
        Case 163: value = register.e: cycles = 4
        Case 164: value = register.h: cycles = 4
        Case 165: value = register.l: cycles = 4
        Case 166: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 230
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    register.a = (register.a And value) And 255
    If CBool(register.a) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    Me.ClearN
    Me.SetH
    Me.ClearC
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub CALL_cc_nn()
    Dim condition As Boolean
    Dim address As Long: address = Me.system.MMU.ReadWord(register.pc)
    register.pc = register.pc + 2
    
    Select Case Me.instructionCode
        Case 196: condition = Not CBool(register.f And 128)
        Case 204: condition = CBool(register.f And 128)
        Case 212: condition = Not CBool(register.f And 16)
        Case 220: condition = CBool(register.f And 16)
    End Select
    
    If condition Then
        register.sp = register.sp - 2
        Me.system.MMU.WriteWord register.sp, register.pc
        register.pc = address
        Me.system.ConsumeClockCycles 24
    Else
        Me.system.ConsumeClockCycles 12
    End If
End Sub

Public Sub CALL_nn()
    register.sp = register.sp - 2
    Me.system.MMU.WriteWord register.sp, register.pc + 2
    register.pc = Me.system.MMU.ReadWord(register.pc)
    Me.system.ConsumeClockCycles 24
End Sub

Public Sub CCF()
    Me.ClearN
    Me.ClearH
    If CBool(register.f And 16) Then
        Me.ClearC
    Else
        Me.SetC
    End If
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub CP()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 191: value = register.a: cycles = 4
        Case 184: value = register.b: cycles = 4
        Case 185: value = register.c: cycles = 4
        Case 186: value = register.d: cycles = 4
        Case 187: value = register.e: cycles = 4
        Case 188: value = register.h: cycles = 4
        Case 189: value = register.l: cycles = 4
        Case 190: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 254
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    Dim result As Long: result = register.a - value
    If (result And 255) = 0 Then
        Me.SetZ
    Else
        Me.ClearZ
    End If
    Me.SetN
    If register.a < value Then
        Me.SetC
    Else
        Me.ClearC
    End If
    If (register.a And 15) < (value And 15) Then
        Me.SetH
    Else
        Me.ClearH
    End If
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub CPL()
    Me.SetN
    Me.SetH
    register.a = register.a Xor 255
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub DAA()
    Dim correction As Long: correction = 0
    Dim flagN As Boolean: flagN = CBool(register.f And 64)
    Dim flagH As Boolean: flagH = CBool(register.f And 32)
    Dim flagc As Boolean: flagc = CBool(register.f And 16)
    
    If flagH Or (Not flagN And (register.a And 15) > 9) Then
        correction = 6
    End If
    
    If flagc Or (Not flagN And register.a > 153) Then
        correction = correction Or 96
        Me.SetC
    End If
    
    register.a = register.a + IIf(flagN, -correction, correction)
    register.a = register.a And 255
    Me.ClearH
    
    If Not CBool(register.a) Then
        Me.SetZ
    Else
        Me.ClearZ
    End If
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub DEC()
    Dim original As Long: original = Empty
    Dim result As Long: result = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 61
            original = register.a
            result = (original - 1) And 255
            register.a = result
            cycles = 4
        Case 5
            original = register.b
            result = (original - 1) And 255
            register.b = result
            cycles = 4
        Case 13
            original = register.c
            result = (original - 1) And 255
            register.c = result
            cycles = 4
        Case 21
            original = register.d
            result = (original - 1) And 255
            register.d = result
            cycles = 4
        Case 29
            original = register.e
            result = (original - 1) And 255
            register.e = result
            cycles = 4
        Case 37
            original = register.h
            result = (original - 1) And 255
            register.h = result
            cycles = 4
        Case 45
            original = register.l
            result = (original - 1) And 255
            register.l = result
            cycles = 4
        Case 53
            original = Me.system.MMU.ReadByte(Me.GetHL())
            result = (original - 1) And 255
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 12
    End Select
    
    If result = 0 Then
        Me.SetZ
    Else
        Me.ClearZ
    End If
    Me.SetN
    If (original And 15) - 1 < 0 Then
        Me.SetH
    Else
        Me.ClearH
    End If
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub DEC_nn()
    Select Case Me.instructionCode
        Case 11
            register.c = (register.c - 1) And 255
            If register.c = 255 Then
                register.b = (register.b - 1) And 255
            End If
        Case 27
            register.e = (register.e - 1) And 255
            If register.e = 255 Then
                register.d = (register.d - 1) And 255
            End If
        Case 43
            register.l = (register.l - 1) And 255
            If register.l = 255 Then
                register.h = (register.h - 1) And 255
            End If
        Case 59
            register.sp = (register.sp - 1) And 65535
    End Select
    
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub DI()
    Me.imeFlag = False
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub EI()
    Me.pendingEnableInterrupts = 17
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub HALT()
    Me.haltFlag = True
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub INC()
    Dim original As Long: original = Empty
    Dim result As Long: result = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 60
            original = register.a
            result = (original + 1) And 255
            register.a = result
            cycles = 4
        Case 4
            original = register.b
            result = (original + 1) And 255
            register.b = result
            cycles = 4
        Case 12
            original = register.c
            result = (original + 1) And 255
            register.c = result
            cycles = 4
        Case 20
            original = register.d
            result = (original + 1) And 255
            register.d = result
            cycles = 4
        Case 28
            original = register.e
            result = (original + 1) And 255
            register.e = result
            cycles = 4
        Case 36
            original = register.h
            result = (original + 1) And 255
            register.h = result
            cycles = 4
        Case 44
            original = register.l
            result = (original + 1) And 255
            register.l = result
            cycles = 4
        Case 52
            original = Me.system.MMU.ReadByte(Me.GetHL())
            result = (original + 1) And 255
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 12
    End Select
    
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    Me.ClearN
    If (original And 15) + 1 > 15 Then
        Me.SetH
    Else
        Me.ClearH
    End If
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub INC_nn()
    Select Case Me.instructionCode
        Case 3
            register.c = (register.c + 1) And 255
            If register.c = 0 Then
                register.b = (register.b + 1) And 255
            End If
        Case 19
            register.e = (register.e + 1) And 255
            If register.e = 0 Then
                register.d = (register.d + 1) And 255
            End If
        Case 35
            register.l = (register.l + 1) And 255
            If register.l = 0 Then
                register.h = (register.h + 1) And 255
            End If
        Case 51
            register.sp = (register.sp + 1) And 65535
    End Select
    
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub JP_d16()
    register.pc = Me.system.MMU.ReadWord(register.pc)
    ' Note: GB CPU manual says this instruction is 12 cycles. It's actually 16.
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub JP_cc_nn()
    Dim condition As Boolean
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 194: condition = Not CBool(register.f And 128)
        Case 202: condition = CBool(register.f And 128)
        Case 210: condition = Not CBool(register.f And 16)
        Case 218: condition = CBool(register.f And 16)
    End Select
    
    If condition Then
        register.pc = Me.system.MMU.ReadWord(register.pc)
        cycles = 16
    Else
        register.pc = register.pc + 2
        cycles = 12
    End If
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub JP_nn()
    register.pc = Me.system.MMU.ReadWord(register.pc)
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub JP_HLmem()
    register.pc = Me.GetHL()
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub JR_cc_n()
    Dim condition As Boolean
    Dim move As Long: move = Me.system.MMU.ReadByte(register.pc)
    register.pc = register.pc + 1
    
    Select Case Me.instructionCode
        Case 32: condition = Not CBool(register.f And 128)
        Case 40: condition = CBool(register.f And 128)
        Case 48: condition = Not CBool(register.f And 16)
        Case 56: condition = CBool(register.f And 16)
    End Select
    
    If condition Then
        If move > 127 Then
            move = -(((Not move) + 1) And 255)
        End If
        register.pc = register.pc + move
        Me.system.ConsumeClockCycles 12
    Else
        Me.system.ConsumeClockCycles 8
    End If
End Sub

Public Sub JR_n()
    Dim move As Long: move = Me.system.MMU.ReadByte(register.pc)
    register.pc = register.pc + 1
    If move > 127 Then
        move = -(((Not move) + 1) And 255)
    End If
    register.pc = register.pc + move
    Me.system.ConsumeClockCycles 12
End Sub

Public Sub LD_d16_SP()
    Dim address As Long: address = Me.system.MMU.ReadWord(register.pc)
    register.pc = register.pc + 2
    Me.system.MMU.WriteWord address, register.sp
    Me.system.ConsumeClockCycles 20
End Sub

Public Sub LD_A_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 120: value = register.b: cycles = 4
        Case 121: value = register.c: cycles = 4
        Case 122: value = register.d: cycles = 4
        Case 123: value = register.e: cycles = 4
        Case 124: value = register.h: cycles = 4
        Case 125: value = register.l: cycles = 4
        Case 10: value = Me.system.MMU.ReadByte(Me.GetBC()): cycles = 8
        Case 26: value = Me.system.MMU.ReadByte(Me.GetDE()): cycles = 8
        Case 126: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 62
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
        Case 250
            value = Me.system.MMU.ReadByte(Me.system.MMU.ReadWord(register.pc))
            register.pc = register.pc + 2
            cycles = 16
    End Select
    
    register.a = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_B_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 64: value = register.b: cycles = 4
        Case 65: value = register.c: cycles = 4
        Case 66: value = register.d: cycles = 4
        Case 67: value = register.e: cycles = 4
        Case 68: value = register.h: cycles = 4
        Case 69: value = register.l: cycles = 4
        Case 70: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
    End Select
    
    register.b = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_C_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 72: value = register.b: cycles = 4
        Case 73: value = register.c: cycles = 4
        Case 74: value = register.d: cycles = 4
        Case 75: value = register.e: cycles = 4
        Case 76: value = register.h: cycles = 4
        Case 77: value = register.l: cycles = 4
        Case 78: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
    End Select
    
    register.c = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_D_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 80: value = register.b: cycles = 4
        Case 81: value = register.c: cycles = 4
        Case 82: value = register.d: cycles = 4
        Case 83: value = register.e: cycles = 4
        Case 84: value = register.h: cycles = 4
        Case 85: value = register.l: cycles = 4
        Case 86: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
    End Select
    
    register.d = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_E_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 88: value = register.b: cycles = 4
        Case 89: value = register.c: cycles = 4
        Case 90: value = register.d: cycles = 4
        Case 91: value = register.e: cycles = 4
        Case 92: value = register.h: cycles = 4
        Case 93: value = register.l: cycles = 4
        Case 94: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
    End Select
    
    register.e = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_H_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 96: value = register.b: cycles = 4
        Case 97: value = register.c: cycles = 4
        Case 98: value = register.d: cycles = 4
        Case 99: value = register.e: cycles = 4
        Case 100: value = register.h: cycles = 4
        Case 101: value = register.l: cycles = 4
        Case 102: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
    End Select
    
    register.h = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_L_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 104: value = register.b: cycles = 4
        Case 105: value = register.c: cycles = 4
        Case 106: value = register.d: cycles = 4
        Case 107: value = register.e: cycles = 4
        Case 108: value = register.h: cycles = 4
        Case 109: value = register.l: cycles = 4
        Case 110: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
    End Select
    
    register.l = value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_HL_SP_r8()
    Dim n As Long: n = Me.system.MMU.ReadByte(register.pc)
    register.pc = register.pc + 1
    If n > 127 Then
        n = -(((Not n) + 1) And 255)
    End If
    Dim result As Long: result = (n + register.sp) And 65535
    register.h = Utils.RightShift(result, 8) And 255
    register.l = result And 255
    Me.ClearZ
    Me.ClearN
    If ((register.sp Xor n Xor result) And 256) = 256 Then
        Me.SetC
    Else
        Me.ClearC
    End If
    If ((register.sp Xor n Xor result) And 16) = 16 Then
        Me.SetH
    Else
        Me.ClearH
    End If
    Me.system.ConsumeClockCycles 12
End Sub

Public Sub LD_HLmem_n()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 112: value = register.b: cycles = 8
        Case 113: value = register.c: cycles = 8
        Case 114: value = register.d: cycles = 8
        Case 115: value = register.e: cycles = 8
        Case 116: value = register.h: cycles = 8
        Case 117: value = register.l: cycles = 8
        Case 54
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 12
    End Select
    
    Me.system.MMU.WriteByte Me.GetHL(), value
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub LD_n_nn()
    Dim word As Long: word = Me.system.MMU.ReadWord(register.pc)
    register.pc = register.pc + 2
    
    Select Case Me.instructionCode
        Case 1: register.b = Utils.RightShift(word, 8) And 255: register.c = word And 255
        Case 17: register.d = Utils.RightShift(word, 8) And 255: register.e = word And 255
        Case 33: register.h = Utils.RightShift(word, 8) And 255: register.l = word And 255
        Case 49: register.sp = word
    End Select
    
    Me.system.ConsumeClockCycles 12
End Sub

Public Sub LD_nn_n()
    Dim value As Long: value = Me.system.MMU.ReadByte(register.pc)
    register.pc = register.pc + 1
    
    Select Case Me.instructionCode
        Case 6: register.b = value
        Case 14: register.c = value
        Case 22: register.d = value
        Case 30: register.e = value
        Case 38: register.h = value
        Case 46: register.l = value
    End Select
    
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LD_n_A()
    Select Case Me.instructionCode
        Case 127
            register.a = register.a
            Me.system.ConsumeClockCycles 4
        Case 71
            register.b = register.a
            Me.system.ConsumeClockCycles 4
        Case 79
            register.c = register.a
            Me.system.ConsumeClockCycles 4
        Case 87
            register.d = register.a
            Me.system.ConsumeClockCycles 4
        Case 95
            register.e = register.a
            Me.system.ConsumeClockCycles 4
        Case 103
            register.h = register.a
            Me.system.ConsumeClockCycles 4
        Case 111
            register.l = register.a
            Me.system.ConsumeClockCycles 4
        Case 2
            Me.system.MMU.WriteByte Me.GetBC(), register.a
            Me.system.ConsumeClockCycles 8
        Case 18
            Me.system.MMU.WriteByte Me.GetDE(), register.a
            Me.system.ConsumeClockCycles 8
        Case 119
            Me.system.MMU.WriteByte Me.GetHL(), register.a
            Me.system.ConsumeClockCycles 8
        Case 234
            Me.system.MMU.WriteByte Me.system.MMU.ReadWord(register.pc), register.a
            register.pc = register.pc + 2
            Me.system.ConsumeClockCycles 16
    End Select
End Sub

Public Sub LD_SP_HL()
    register.sp = Utils.LeftShift(register.h, 8) + register.l
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LDD_A_HLmem()
    register.a = Me.system.MMU.ReadByte(Me.GetHL())
    register.l = (register.l - 1) And 255
    If register.l = 255 Then
        register.h = (register.h - 1) And 255
    End If
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LDD_HLmem_A()
    Me.system.MMU.WriteByte Me.GetHL(), register.a
    register.l = (register.l - 1) And 255
    If register.l = 255 Then
        register.h = (register.h - 1) And 255
    End If
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LD_A_Cmem()
    register.a = Me.system.MMU.ReadByte(65280 + register.c)
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LD_Cmem_A()
    Me.system.MMU.WriteByte 65280 + register.c, register.a
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LDH_A_d8mem()
    register.a = Me.system.MMU.ReadByte(65280 + Me.system.MMU.ReadByte(register.pc))
    register.pc = register.pc + 1
    Me.system.ConsumeClockCycles 12
End Sub

Public Sub LDH_d8mem_A()
    Me.system.MMU.WriteByte 65280 + Me.system.MMU.ReadByte(register.pc), register.a
    register.pc = register.pc + 1
    Me.system.ConsumeClockCycles 12
End Sub

Public Sub LDI_A_HLmem()
    register.a = Me.system.MMU.ReadByte(Me.GetHL())
    register.l = (register.l + 1) And 255
    If register.l = 0 Then
        register.h = (register.h + 1) And 255
    End If
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub LDI_HLmem_A()
    Me.system.MMU.WriteByte Me.GetHL(), register.a
    register.l = (register.l + 1) And 255
    If register.l = 0 Then
        register.h = (register.h + 1) And 255
    End If
    Me.system.ConsumeClockCycles 8
End Sub

Public Sub NOP()
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub ORR()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 183: value = register.a: cycles = 4
        Case 176: value = register.b: cycles = 4
        Case 177: value = register.c: cycles = 4
        Case 178: value = register.d: cycles = 4
        Case 179: value = register.e: cycles = 4
        Case 180: value = register.h: cycles = 4
        Case 181: value = register.l: cycles = 4
        Case 182: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 246
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    register.a = (register.a Or value) And 255
    If register.a Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    Me.ClearN
    Me.ClearH
    Me.ClearC
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub POP()
    Select Case Me.instructionCode
        Case 241
            register.f = Me.system.MMU.ReadByte(register.sp) And 240
            register.sp = register.sp + 1
            register.a = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
        Case 193
            register.c = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
            register.b = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
        Case 209
            register.e = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
            register.d = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
        Case 225
            register.l = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
            register.h = Me.system.MMU.ReadByte(register.sp)
            register.sp = register.sp + 1
    End Select
    
    Me.system.ConsumeClockCycles 12
End Sub

Public Sub PUSH()
    Dim value As Long: value = Empty
    
    Select Case Me.instructionCode
        Case 245: value = Me.GetAF()
        Case 197: value = Me.GetBC()
        Case 213: value = Me.GetDE()
        Case 229: value = Me.GetHL()
    End Select
    
    register.sp = register.sp - 2
    Me.system.MMU.WriteWord register.sp, value
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub RET()
    register.pc = Me.system.MMU.ReadWord(register.sp)
    register.sp = register.sp + 2
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub RET_cc()
    Dim condition As Boolean
    
    Select Case Me.instructionCode
        Case 192: condition = Not CBool(register.f And 128)
        Case 200: condition = CBool(register.f And 128)
        Case 208: condition = Not CBool(register.f And 16)
        Case 216: condition = CBool(register.f And 16)
    End Select
    
    If condition Then
        register.pc = Me.system.MMU.ReadWord(register.sp)
        register.sp = register.sp + 2
        Me.system.ConsumeClockCycles 20
    Else
        Me.system.ConsumeClockCycles 8
    End If
End Sub

Public Sub RETI()
    Dim address As Long: address = Me.system.MMU.ReadWord(register.sp)
    register.sp = register.sp + 2
    register.pc = address
    Me.imeFlag = True
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub RLA()
    Dim a As Long: a = register.a
    Dim carryOut As Long: carryOut = IIf(CBool(register.a And 128), 1, 0)
    Dim carryIn As Long: carryIn = IIf(CBool(register.f And 16), 1, 0)
    register.a = (Utils.LeftShift(a, 1) + carryIn) And 255
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    Me.ClearZ
    Me.ClearN
    Me.ClearH
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub RLCA()
    Me.ClearN
    Me.ClearH
    Me.ClearZ
    Dim carryOut As Long: carryOut = IIf(CBool(register.a And 128), 1, 0)
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    register.a = (Utils.LeftShift(register.a, 1) + carryOut) And 255
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub RRA()
    Dim carryIn As Long: carryIn = IIf(CBool(register.f And 16), 1, 0)
    Dim carryOut As Long: carryOut = IIf(CBool(register.a And 1), 1, 0)
    register.a = (Utils.RightShift(register.a, 1) + Utils.LeftShift(carryIn, 7)) And 255
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    Me.ClearZ
    Me.ClearN
    Me.ClearH
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub RRCA()
    Me.ClearN
    Me.ClearH
    Me.ClearZ
    Dim carryOut As Long: carryOut = IIf(CBool(register.a And 1), 1, 0)
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    register.a = (Utils.RightShift(register.a, 1) + Utils.LeftShift(carryOut, 7)) And 255
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub RST()
    Dim offset As Long: offset = Empty
    
    Select Case Me.instructionCode
        Case 199: offset = 0
        Case 207: offset = 8
        Case 215: offset = 16
        Case 223: offset = 24
        Case 231: offset = 32
        Case 239: offset = 40
        Case 247: offset = 48
        Case 255: offset = 56
    End Select
    
    register.sp = register.sp - 2
    Me.system.MMU.WriteWord register.sp, register.pc
    register.pc = 0 + offset
    Me.system.ConsumeClockCycles 16
End Sub

Public Sub SBC()
    Dim cycles As Long: cycles = Empty
    Dim inputt As Long: inputt = Empty
    Dim a As Long: a = register.a
    
    Select Case Me.instructionCode
        Case 159: inputt = register.a: cycles = 4
        Case 152: inputt = register.b: cycles = 4
        Case 153: inputt = register.c: cycles = 4
        Case 154: inputt = register.d: cycles = 4
        Case 155: inputt = register.e: cycles = 4
        Case 156: inputt = register.h: cycles = 4
        Case 157: inputt = register.l: cycles = 4
        Case 158: inputt = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 222
            inputt = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    Dim result As Long: result = a - inputt - IIf(CBool(register.f And 16), 1, 0)
    Me.SetN
    If CBool(result And 255) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    If CBool(result And 256) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    If ((a Xor inputt Xor result) And 16) <> 0 Then
        Me.SetH
    Else
        Me.ClearH
    End If
    register.a = result And 255
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub SCF()
    Me.SetC
    Me.ClearN
    Me.ClearH
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub Stopp()
    Me.stopFlag = True
    register.pc = register.pc + 1
    Me.system.ConsumeClockCycles 4
End Sub

Public Sub SUBB()
    Dim cycles As Long: cycles = Empty
    Dim inputt As Long: inputt = Empty
    
    Select Case Me.instructionCode
        Case 151: inputt = register.a: cycles = 4
        Case 144: inputt = register.b: cycles = 4
        Case 145: inputt = register.c: cycles = 4
        Case 146: inputt = register.d: cycles = 4
        Case 147: inputt = register.e: cycles = 4
        Case 148: inputt = register.h: cycles = 4
        Case 149: inputt = register.l: cycles = 4
        Case 150: inputt = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 214
            inputt = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    Dim a As Long: a = register.a
    register.a = register.a - inputt
    Me.SetN
    If register.a < 0 Then
        Me.SetC
    Else
        Me.ClearC
    End If
    register.a = register.a And 255
    If CBool(register.a) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    If CBool((register.a Xor inputt Xor a) And 16) Then
        Me.SetH
    Else
        Me.ClearH
    End If
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub XORR()
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 175 ' XOR A
            value = register.a: cycles = 4
        Case 168 ' XOR B
            value = register.b: cycles = 4
        Case 169 ' XOR C
            value = register.c: cycles = 4
        Case 170 ' XOR D
            value = register.d: cycles = 4
        Case 171 ' XOR E
            value = register.e: cycles = 4
        Case 172 ' XOR H
            value = register.h: cycles = 4
        Case 173 ' XOR L
            value = register.l: cycles = 4
        Case 174 ' XOR (HL)
            value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 8
        Case 238 ' XOR n
            value = Me.system.MMU.ReadByte(register.pc)
            register.pc = register.pc + 1
            cycles = 8
    End Select
    
    register.a = register.a Xor value
    register.a = register.a And 255
    
    If Not CBool(register.a) Then
        Me.SetZ
    Else
        Me.ClearZ
    End If
    Me.ClearN
    Me.ClearH
    Me.ClearC
    Me.system.ConsumeClockCycles cycles
End Sub

' *****************************************************************
' CB Instructions
' *****************************************************************

Public Sub CB()
    Me.instructionCode = CLng(51968) + Me.system.MMU.ReadByte(register.pc)
    register.pc = register.pc + 1
    Select Case instructionCode
        '*****************************************************************
        ' Main Instructions
        '*****************************************************************
        Case 143, 136, 137, 138, 139, 140, 141, 142, 206
            ADC
        Case 135, 128, 129, 130, 131, 132, 133, 134, 198
            ADD
        Case 9, 25, 41, 57
            ADD_HL_n
        Case 232
            ADD_SP_n
        Case 167, 160, 161, 162, 163, 164, 165, 166, 230
            ANDD
        Case 196, 204, 212, 220
            CALL_cc_nn
        Case 205
            CALL_nn
        Case 63
            CCF
        Case 191, 184, 185, 186, 187, 188, 189, 190, 254
            CP
        Case 47
            CPL
        Case 39
            DAA
        Case 61, 5, 13, 21, 29, 37, 45, 53
            DEC
        Case 11, 27, 43, 59
            DEC_nn
        Case 243
            DI
        Case 251
            EI
        Case 118
            HALT
        Case 60, 4, 12, 20, 28, 36, 44, 52
            INC
        Case 3, 19, 35, 51
            INC_nn
        Case 194, 202, 210, 218
            JP_cc_nn
        Case 195
            JP_nn
        Case 233
            JP_HLmem
        Case 32, 40, 48, 56
            JR_cc_n
        Case 24
            JR_n
        Case 8
            LD_d16_SP
        Case 120, 121, 122, 123, 124, 125, 10, 26, 126, 62, 250
            LD_A_n
        Case 64, 65, 66, 67, 68, 69, 70
            LD_B_n
        Case 72, 73, 74, 75, 76, 77, 78
            LD_C_n
        Case 80, 81, 82, 83, 84, 85, 86
            LD_D_n
        Case 88, 89, 90, 91, 92, 93, 94
            LD_E_n
        Case 96, 97, 98, 99, 100, 101, 102
            LD_H_n
        Case 104, 105, 106, 107, 108, 109, 110
            LD_L_n
        Case 248
            LD_HL_SP_r8
        Case 112, 113, 114, 115, 116, 117, 54
            LD_HLmem_n
        Case 127, 71, 79, 87, 95, 103, 111, 2, 18, 119, 234
            LD_n_A
        Case 1, 17, 33, 49
            LD_n_nn
        Case 6, 14, 22, 30, 38, 46
            LD_nn_n
        Case 249
            LD_SP_HL
        Case 242
            LD_A_Cmem
        Case 226
            LD_Cmem_A
        Case 58
            LDD_A_HLmem
        Case 50
            LDD_HLmem_A
        Case 240
            LDH_A_d8mem
        Case 224
            LDH_d8mem_A
        Case 42
            LDI_A_HLmem
        Case 34
            LDI_HLmem_A
        Case 0
            NOP
        Case 183, 176, 177, 178, 179, 180, 181, 182, 246
            ORR
        Case 241, 193, 209, 225
            POP
        Case 245, 197, 213, 229
            PUSH
        Case 201
            RET
        Case 192, 200, 208, 216
            RET_cc
        Case 217
            RETI
        Case 23
            RLA
        Case 7
            RLCA
        Case 31
            RRA
        Case 15
            RRCA
        Case 199, 207, 215, 223, 231, 239, 247, 255
            RST
        Case 152, 153, 154, 155, 156, 157, 158, 159, 222
            SBC
        Case 55
            SCF
        Case 16
            Stopp
        Case 151, 144, 145, 146, 147, 148, 149, 150, 214
            SUBB
        Case 175, 168, 169, 170, 171, 172, 173, 174, 238
            XORR
        ' *****************************************************************
        '  CB Instructions
        ' *****************************************************************
        Case 203
            CB
        Case 51968, 51969, 51970, 51971, 51972, 51973, 51974, 51975
            RLC
        Case 51976, 51977, 51978, 51979, 51980, 51981, 51982, 51983
            RRC
        Case 51984, 51985, 51986, 51987, 51988, 51989, 51990, 51991
            RL
        Case 51992, 51993, 51994, 51995, 51996, 51997, 51998, 51999
            RR
        Case 52000, 52001, 52002, 52003, 52004, 52005, 52006, 52007
            SLA
        Case 52008, 52009, 52010, 52011, 52012, 52013, 52014, 52015
            SRA
        Case 52016, 52017, 52018, 52019, 52020, 52021, 52022, 52023
            SWAP
        Case 52024, 52025, 52026, 52027, 52028, 52029, 52030, 52031
            SRL
        Case 52032, 52033, 52034, 52035, 52036, 52037, 52038, 52039, _
             52040, 52041, 52042, 52043, 52044, 52045, 52046, 52047, _
             52048, 52049, 52050, 52051, 52052, 52053, 52054, 52055, _
             52056, 52057, 52058, 52059, 52060, 52061, 52062, 52063, _
             52064, 52065, 52066, 52067, 52068, 52069, 52070, 52071, _
             52072, 52073, 52074, 52075, 52076, 52077, 52078, 52079, _
             52080, 52081, 52082, 52083, 52084, 52085, 52086, 52087, _
             52088, 52089, 52090, 52091, 52092, 52093, 52094, 52095
            bit
        Case 52096, 52097, 52098, 52099, 52100, 52101, 52102, 52103, _
             52104, 52105, 52106, 52107, 52108, 52109, 52110, 52111, _
             52112, 52113, 52114, 52115, 52116, 52117, 52118, 52119, _
             52120, 52121, 52122, 52123, 52124, 52125, 52126, 52127, _
             52128, 52129, 52130, 52131, 52132, 52133, 52134, 52135, _
             52136, 52137, 52138, 52139, 52140, 52141, 52142, 52143, _
             52144, 52145, 52146, 52147, 52148, 52149, 52150, 52151, _
             52152, 52153, 52154, 52155, 52156, 52157, 52158, 52159
            RES
        Case 52160, 52161, 52162, 52163, 52164, 52165, 52166, 52167, _
             52168, 52169, 52170, 52171, 52172, 52173, 52174, 52175, _
             52176, 52177, 52178, 52179, 52180, 52181, 52182, 52183, _
             52184, 52185, 52186, 52187, 52188, 52189, 52190, 52191, _
             52192, 52193, 52194, 52195, 52196, 52197, 52198, 52199, _
             52200, 52201, 52202, 52203, 52204, 52205, 52206, 52207, _
             52208, 52209, 52210, 52211, 52212, 52213, 52214, 52215, _
             52216, 52217, 52218, 52219, 52220, 52221, 52222, 52223
            SETT
        Case Else
            Err.Raise 50000, "LR35902", "Instruction " & Me.instructionCode & " not found"
    End Select
End Sub

Public Sub bit()
    Dim bit As Long: bit = Empty
    Dim value As Long: value = Empty
    Dim cycles As Long: cycles = Empty
    
    Dim codeHI As Long: codeHI = Utils.RightShift(Me.instructionCode, 4) And 15
    Dim codeLO As Long: codeLO = Me.instructionCode And 15
    
    Select Case codeLO
        Case 7, 15: value = register.a: cycles = 8
        Case 0, 8: value = register.b: cycles = 8
        Case 1, 9: value = register.c: cycles = 8
        Case 2, 10: value = register.d: cycles = 8
        Case 3, 11: value = register.e: cycles = 8
        Case 4, 12: value = register.h: cycles = 8
        Case 5, 13: value = register.l: cycles = 8
        Case 6, 14: value = Me.system.MMU.ReadByte(Me.GetHL()): cycles = 12
    End Select
    
    Select Case codeHI
        Case 4
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 0
            Else
                bit = 1
            End If
        Case 5
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 2
            Else
                bit = 3
            End If
        Case 6
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 4
            Else
                bit = 5
            End If
        Case 7
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 6
            Else
                bit = 7
            End If
    End Select
    
    If value And Utils.LeftShift(1, bit) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    Me.ClearN
    Me.SetH
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub RES()
    Dim bit As Long: bit = Empty
    Dim cycles As Long: cycles = Empty
    Dim hl As Long
    
    Dim codeHI As Long: codeHI = Utils.RightShift(Me.instructionCode, 4) And 15
    Dim codeLO As Long: codeLO = Me.instructionCode And 15
    
    Select Case codeHI
        Case 8
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 0
            Else
                bit = 1
            End If
        Case 9
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 2
            Else
                bit = 3
            End If
        Case 10
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 4
            Else
                bit = 5
            End If
        Case 11
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 6
            Else
                bit = 7
            End If
    End Select
    
    Select Case codeLO
        Case 7, 15
            register.a = register.a And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 0, 8
            register.b = register.b And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 1, 9
            register.c = register.c And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 2, 10
            register.d = register.d And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 3, 11
            register.e = register.e And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 4, 12
            register.h = register.h And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 5, 13
            register.l = register.l And (255) - Utils.LeftShift(1, bit)
            cycles = 8
        Case 6, 14
            hl = Me.system.MMU.ReadByte(Me.GetHL())
            Me.system.MMU.WriteByte Me.GetHL(), hl And (255) - Utils.LeftShift(1, bit)
            cycles = 16
    End Select
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub RL()
    Me.ClearN
    Me.ClearH
    Dim cycles As Long: cycles = Empty
    Dim result As Long: result = Empty
    
    Select Case Me.instructionCode
        Case 51991
            result = Me.p_RL(register.a)
            register.a = result
            cycles = 8
        Case 51984
            result = Me.p_RL(register.b)
            register.b = result
            cycles = 8
        Case 51985
            result = Me.p_RL(register.c)
            register.c = result
            cycles = 8
        Case 51986
            result = Me.p_RL(register.d)
            register.d = result
            cycles = 8
        Case 51987
            result = Me.p_RL(register.e)
            register.e = result
            cycles = 8
        Case 51988
            result = Me.p_RL(register.h)
            register.h = result
            cycles = 8
        Case 51989
            result = Me.p_RL(register.l)
            register.l = result
            cycles = 8
        Case 51990
            result = Me.p_RL(Me.system.MMU.ReadByte(Me.GetHL()))
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 16
    End Select
    
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Function p_RL(ByVal reg As Long) As Long
    Dim carryOut As Long: carryOut = IIf(CBool(reg And 128), 1, 0)
    Dim carryIn As Long: carryIn = IIf(CBool(register.f And 16), 1, 0)
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    p_RL = (Utils.LeftShift(reg, 1) + carryIn) And 255
End Function

Public Sub RLC()
    Me.ClearN
    Me.ClearH
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 51975
            register.a = Me.p_RLC(register.a)
            cycles = 8
        Case 51968
            register.b = Me.p_RLC(register.b)
            cycles = 8
        Case 51969
            register.c = Me.p_RLC(register.c)
            cycles = 8
        Case 51970
            register.d = Me.p_RLC(register.d)
            cycles = 8
        Case 51971
            register.e = Me.p_RLC(register.e)
            cycles = 8
        Case 51972
            register.h = Me.p_RLC(register.h)
            cycles = 8
        Case 51973
            register.l = Me.p_RLC(register.l)
            cycles = 8
        Case 51974
            Me.system.MMU.WriteByte Me.GetHL(), Me.p_RLC(Me.system.MMU.ReadByte(Me.GetHL()))
            cycles = 16
    End Select
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Function p_RLC(ByVal reg As Long) As Long
    Dim carryOut As Long: carryOut = IIf(CBool(reg And 128), 1, 0)
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    Dim result As Long: result = (Utils.LeftShift(reg, 1) + carryOut) And 255
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    p_RLC = result
End Function

Public Sub RR()
    Me.ClearN
    Me.ClearH
    Dim cycles As Long: cycles = Empty
    
    Select Case Me.instructionCode
        Case 51999
            register.a = Me.p_RR(register.a)
            cycles = 8
        Case 51992
            register.b = Me.p_RR(register.b)
            cycles = 8
        Case 51993
            register.c = Me.p_RR(register.c)
            cycles = 8
        Case 51994
            register.d = Me.p_RR(register.d)
            cycles = 8
        Case 51995
            register.e = Me.p_RR(register.e)
            cycles = 8
        Case 51996
            register.h = Me.p_RR(register.h)
            cycles = 8
        Case 51997
            register.l = Me.p_RR(register.l)
            cycles = 8
        Case 51998
            Me.system.MMU.WriteByte Me.GetHL(), Me.p_RR(Me.system.MMU.ReadByte(Me.GetHL()))
            cycles = 16
    End Select
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Function p_RR(ByVal reg As Long) As Long
    Dim carryOut As Long: carryOut = IIf(CBool(reg And 1), 1, 0)
    Dim carryIn As Long: carryIn = IIf(CBool(register.f And 16), 1, 0)
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    Dim result As Long: result = (Utils.RightShift(reg, 1) + Utils.LeftShift(carryIn, 7)) And 255
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    p_RR = result
End Function

Public Sub RRC()
    Dim cycles As Long: cycles = Empty
    
    Me.ClearN
    Me.ClearH
    
    Select Case Me.instructionCode
        Case 51983
            register.a = Me.p_RRC(register.a)
            cycles = 8
        Case 51976
            register.b = Me.p_RRC(register.b)
            cycles = 8
        Case 51977
            register.c = Me.p_RRC(register.c)
            cycles = 8
        Case 51978
            register.d = Me.p_RRC(register.d)
            cycles = 8
        Case 51979
            register.e = Me.p_RRC(register.e)
            cycles = 8
        Case 51980
            register.h = Me.p_RRC(register.h)
            cycles = 8
        Case 51981
            register.l = Me.p_RRC(register.l)
            cycles = 8
        Case 51982
            Me.system.MMU.WriteByte Me.GetHL(), Me.p_RRC(Me.system.MMU.ReadByte(Me.GetHL()))
            cycles = 16
    End Select
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Function p_RRC(ByVal reg As Long) As Long
    Dim carryOut As Long: carryOut = IIf(CBool(reg And 1), 1, 0)
    If CBool(carryOut) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    Dim result As Long: result = (Utils.RightShift(reg, 1) + Utils.LeftShift(carryOut, 7)) And 255
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    p_RRC = result
End Function

Public Sub SETT()
    Dim bit As Long: bit = Empty
    Dim cycles As Long: cycles = Empty
    
    Dim codeHI As Long: codeHI = Utils.RightShift(Me.instructionCode, 4) And 15
    Dim codeLO As Long: codeLO = Me.instructionCode And 15
    
    Select Case codeHI
        Case 12
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 0
            Else
                bit = 1
            End If
        Case 13
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 2
            Else
                bit = 3
            End If
        Case 14
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 4
            Else
                bit = 5
            End If
        Case 15
            If codeLO >= 0 And codeLO <= 7 Then
                bit = 6
            Else
                bit = 7
            End If
    End Select
    
    Dim hl As Long
    Select Case codeLO
        Case 7, 15
            register.a = register.a Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 0, 8
            register.b = register.b Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 1, 9
            register.c = register.c Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 2, 10
            register.d = register.d Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 3, 11
            register.e = register.e Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 4, 12
            register.h = register.h Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 5, 13
            register.l = register.l Or Utils.LeftShift(1, bit)
            cycles = 8
        Case 6, 14
            hl = Me.system.MMU.ReadByte(Me.GetHL())
            Me.system.MMU.WriteByte Me.GetHL(), hl Or Utils.LeftShift(1, bit)
            cycles = 16
    End Select
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub SLA()
    Me.ClearN
    Me.ClearH
    Dim cycles As Long: cycles = Empty
    Dim original As Long: original = Empty
    Dim result As Long: result = Empty
    
    Select Case Me.instructionCode
        Case 52007
            original = register.a
            result = Utils.LeftShift(original, 1) And 255
            register.a = result
            cycles = 8
        Case 52000
            original = register.b
            result = Utils.LeftShift(original, 1) And 255
            register.b = result
            cycles = 8
        Case 52001
            original = register.c
            result = Utils.LeftShift(original, 1) And 255
            register.c = result
            cycles = 8
        Case 52002
            original = register.d
            result = Utils.LeftShift(original, 1) And 255
            register.d = result
            cycles = 8
        Case 52003
            original = register.e
            result = Utils.LeftShift(original, 1) And 255
            register.e = result
            cycles = 8
        Case 52004
            original = register.h
            result = Utils.LeftShift(original, 1) And 255
            register.h = result
            cycles = 8
        Case 52005
            original = register.l
            result = Utils.LeftShift(original, 1) And 255
            register.l = result
            cycles = 8
        Case 52006
            original = Me.system.MMU.ReadByte(Me.GetHL())
            result = Utils.LeftShift(original, 1) And 255
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 16
    End Select
    
    If CBool(original And 128) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub SRA()
    Dim cycles As Long: cycles = Empty
    Dim original As Long: original = Empty
    Dim result As Long: result = Empty
    Me.ClearN
    Me.ClearH
    
    Select Case Me.instructionCode
        Case 52015
            original = register.a
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.a = result
            cycles = 8
        Case 52008
            original = register.b
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.b = result
            cycles = 8
        Case 52009
            original = register.c
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.c = result
            cycles = 8
        Case 52010
            original = register.d
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.d = result
            cycles = 8
        Case 52011
            original = register.e
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.e = result
            cycles = 8
        Case 52012
            original = register.h
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.h = result
            cycles = 8
        Case 52013
            original = register.l
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            register.l = result
            cycles = 8
        Case 52014
            original = Me.system.MMU.ReadByte(Me.GetHL())
            result = (Utils.RightShift(original, 1) + (original And 128)) And 255
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 16
    End Select
    
    If CBool(original And 1) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub SRL()
    Dim cycles As Long: cycles = Empty
    Dim original As Long: original = Empty
    Dim result As Long: result = Empty
    Me.ClearN
    Me.ClearH
    
    Select Case Me.instructionCode
        Case 52031
            original = register.a
            result = Utils.RightShift(original, 1) And 255
            register.a = result
            cycles = 8
        Case 52024
            original = register.b
            result = Utils.RightShift(original, 1) And 255
            register.b = result
            cycles = 8
        Case 52025
            original = register.c
            result = Utils.RightShift(original, 1) And 255
            register.c = result
            cycles = 8
        Case 52026
            original = register.d
            result = Utils.RightShift(original, 1) And 255
            register.d = result
            cycles = 8
        Case 52027
            original = register.e
            result = Utils.RightShift(original, 1) And 255
            register.e = result
            cycles = 8
        Case 52028
            original = register.h
            result = Utils.RightShift(original, 1) And 255
            register.h = result
            cycles = 8
        Case 52029
            original = register.l
            result = Utils.RightShift(original, 1) And 255
            register.l = result
            cycles = 8
        Case 52030
            original = Me.system.MMU.ReadByte(Me.GetHL())
            result = Utils.RightShift(original, 1) And 255
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 16
    End Select
    
    If CBool(original And 1) Then
        Me.SetC
    Else
        Me.ClearC
    End If
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Sub SWAP()
    Dim cycles As Long: cycles = Empty
    Dim result As Long: result = Empty
    
    Select Case Me.instructionCode
        Case 52023
            result = Me.p_SWAP(register.a)
            register.a = result
            cycles = 8
        Case 52016
            result = Me.p_SWAP(register.b)
            register.b = result
            cycles = 8
        Case 52017
            result = Me.p_SWAP(register.c)
            register.c = result
            cycles = 8
        Case 52018
            result = Me.p_SWAP(register.d)
            register.d = result
            cycles = 8
        Case 52019
            result = Me.p_SWAP(register.e)
            register.e = result
            cycles = 8
        Case 52020
            result = Me.p_SWAP(register.h)
            register.h = result
            cycles = 8
        Case 52021
            result = Me.p_SWAP(register.l)
            register.l = result
            cycles = 8
        Case 52022
            result = Me.p_SWAP(Me.system.MMU.ReadByte(Me.GetHL()))
            Me.system.MMU.WriteByte Me.GetHL(), result
            cycles = 16
    End Select
    
    If CBool(result) Then
        Me.ClearZ
    Else
        Me.SetZ
    End If
    
    Me.ClearN
    Me.ClearH
    Me.ClearC
    
    Me.system.ConsumeClockCycles cycles
End Sub

Public Function p_SWAP(ByVal value As Long) As Long
    p_SWAP = Utils.LeftShift((value And 15), 4) + Utils.RightShift(value, 4)
End Function
