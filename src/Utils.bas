Attribute VB_Name = "Utils"
Option Explicit
Option Base 0

Public Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Const VK_LEFT As Long = &H25 'LEFT ARROW key
Public Const VK_UP As Long = &H26 'UP ARROW key
Public Const VK_RIGHT As Long = &H27 'RIGHT ARROW key
Public Const VK_DOWN As Long = &H28 'DOWN ARROW key
Public Const VK_SPACE As Long = &H20 'SPACEBAR
Public Const VK_RETURN As Long = &HD 'ENTER key
Public Const VK_CONTROL As Long = &H11 'CTRL key
Public Const VK_MENU As Long = &H12 'ALT key

Type ApuRegisters
    nr10 As Long
    nr11 As Long
    nr12 As Long
    nr13 As Long
    nr14 As Long
    
    nr21 As Long
    nr22 As Long
    nr23 As Long
    nr24 As Long
    
    nr30 As Long
    nr31 As Long
    nr32 As Long
    nr33 As Long
    nr34 As Long
    
    nr41 As Long
    nr42 As Long
    nr43 As Long
    nr44 As Long
    
    nr50 As Long
    nr51 As Long
    nr52 As Long
End Type

Type JoypadRegisters
    p1 As Long
End Type

Type Mbc3Register
    seconds As Long
    minutes As Long
    hours As Long
    days As Long
    dayCarry As Long
End Type

Type TimeSpan
    seconds As Long
    minutes As Long
    hours As Long
    days As Long
End Type

Type SerialRegister
    sb As Long
    sc As Long
End Type

Type MmuRegisters
    if As Long
    ie As Long
    
    key1 As Long
    tp As Long
    svbk As Long
End Type

Type CpuRegister
    a As Long
    f As Long
    b As Long
    c As Long
    d As Long
    e As Long
    h As Long
    l As Long
    pc As Long
    sp As Long
End Type

Type TimerRegisters
    div As Long
    tima As Long
    tma As Long
    tac As Long
End Type

Type GpuRegisters
    lcdc As Long
    stat As Long
    scy As Long
    scx As Long
    ly As Long
    lyc As Long
    dma As Long
    bgp As Long
    obj0 As Long
    obj1 As Long
    wy As Long
    wx As Long
    
    vbk As Long
    
    hdma1 As Long
    hdma2 As Long
    hdma3 As Long
    hdma4 As Long
    hdma5 As Long
    
    bgpi As Long
    bgpd As Long
    obpi As Long
    obpd As Long
End Type

Public Function RightShift(ByVal value As Long, ByVal Shift As Byte) As Long
    
    RightShift = value
    
    If Shift > 0 Then
        RightShift = Int(RightShift / (2 ^ Shift))
    End If
    
End Function

Public Function LeftShift(ByVal value As Long, ByVal Shift As Integer) As Long
    LeftShift = value
    
    If Shift = -1 Then
        If value Mod 2 = 1 Then
            LeftShift = -2147483648#
            Exit Function
        Else
            LeftShift = 0
            Exit Function
        End If
    End If
    
    If Shift > 0 Then
        LeftShift = value * (CLng(2) ^ Shift)
    End If
End Function

Public Function GetRam(ByVal id As String) As RamLoader

    Dim returnObj As New RamLoader
    Dim ramSht As Worksheet: Set ramSht = ThisWorkbook.Sheets("ram")
    Dim colNum As Integer: colNum = 1
    Dim rowNum As Long
    Dim lastRow As Long
    
    Do While True
        If ramSht.Cells(1, colNum).value = "" Then
            GoTo NOT_SUCCESSFUL
        ElseIf ramSht.Cells(1, colNum).value = id Then
            lastRow = ramSht.Cells(1, colNum).End(xlDown).row
            returnObj.RedimRam 0, lastRow - 2
            
            For rowNum = 2 To lastRow
                returnObj.SetRamByteAt rowNum - 2, CByte(ramSht.Cells(rowNum, colNum).value)
            Next rowNum
            GoTo SUCCESSFUL
            
        End If
        colNum = colNum + 1
    Loop
    
SUCCESSFUL:
    returnObj.wasSuccessful = True
    Set GetRam = returnObj
    Exit Function
    
    
NOT_SUCCESSFUL:
    returnObj.wasSuccessful = False
    Set GetRam = returnObj
    Exit Function

End Function

Public Sub SetRam(ByVal id As String, ByRef ram() As Byte)

    Dim ramSht As Worksheet: Set ramSht = ThisWorkbook.Sheets("ram")
    Dim colNum As Integer: colNum = 1
    Dim rowNum As Long
    
    Do While True
        If ramSht.Cells(1, colNum).value = "" Or ramSht.Cells(1, colNum).value = id Then
            ramSht.Cells(1, colNum).value = id
            
            For rowNum = 2 To UBound(ram) + 2
                ramSht.Cells(rowNum, colNum).value = ram(rowNum - 2)
            Next rowNum
            
            Exit Sub
        End If
        colNum = colNum + 1
    Loop

End Sub
