VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gpu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public system As Gameboy
Public screenWidth As Long
Public screenHeight As Long
Private vram() As Byte
Private oam() As Byte
Private frameBuffer() As Pixel
Public oamDmaActive As Boolean
Public oamDmaByte As Byte
Private registers As GpuRegisters
Private bgColors() As Long
Private obj0Colors() As Long
Private obj1Colors() As Long
Public clock As LongLong
Public sht As Worksheet
Public frameCounter As Integer

Public Sub Class_Initialize()
    Me.frameCounter = 0

    Me.screenWidth = 160
    Me.screenHeight = 144
    ReDim vram(0 To 8191)
    ReDim oam(0 To 159)
    ReDim frameBuffer(0 To 255, 0 To 255)
    
    Dim i As Long
    
    ' Reset video RAM
    For i = LBound(vram) To UBound(vram)
        vram(i) = Int(Rnd() * 256)
    Next i
    
    ' Sprite Attribute Memory (OAM)
    For i = LBound(oam) To UBound(oam)
        oam(i) = Int(Rnd() * 256)
    Next i
    
    Me.oamDmaActive = False
    Me.oamDmaByte = 0
    
    registers.lcdc = 0 ' $FF40 (r/w) LCD control
    registers.stat = 0 ' $FF41 (r/w) LCDC Status
    registers.scy = 0 ' $FF42 (r/w) Scroll Y
    registers.scx = 0 ' $FF43 (r/w) Scroll X TODO: Fill in full memory address
    registers.ly = 0 ' $FF44 (r) LCDC Y-coordinate
    registers.lyc = 0 ' $FF45 (r/w) LY Compare
    registers.dma = 0 ' $FF46 (w) DM Transfer & Start Address
    registers.bgp = 0 ' $FF47 (r/w) BG & Window palette
    registers.obj0 = 0 ' $FF48 (r/w) OBJ 0 Palette
    registers.obj1 = 0 ' $FF49 (r/w) OBJ 1 Palette
    registers.wy = 0 ' $FF4A (r/w) Window Y position
    registers.wx = 0 ' $FF4B (r/w) Window X position
    
    registers.vbk = 0 ' $FF4F VRAM Bank
    
    registers.hdma1 = 0 ' $FF51 New DMA source, high
    registers.hdma2 = 0 ' $FF52 New DMA source, low
    registers.hdma3 = 0 ' $FF53 New DMA destination, high
    registers.hdma4 = 0 ' $FF54 New DMA destination, low
    registers.hdma5 = 0 ' $FF55 New DMA length/mode/start
    
    registers.bgpi = 0 ' $FF68 Background Palette Index
    registers.bgpd = 0 ' $FF69 Background Palette Data
    registers.obpi = 0 ' $FF6A Sprite Palette Index
    registers.obpd = 0 ' $FF6B Sprite Palette Data
    
    ReDim bgColors(0 To 3)
    bgColors(0) = 14743760 ' White
    bgColors(1) = 8962160 ' Light Grey
    bgColors(2) = 3434582 ' Dark Grey
    bgColors(3) = 530464  ' Black
    
    ReDim obj0Colors(0 To 3)
    obj0Colors(0) = 14743760 ' White
    obj0Colors(1) = 8962160 ' Light Grey
    obj0Colors(2) = 3434582 ' Dark Grey
    obj0Colors(3) = 530464  ' Black
    
    ReDim obj1Colors(0 To 3)
    obj1Colors(0) = 14743760 ' White
    obj1Colors(1) = 8962160 ' Light Grey
    obj1Colors(2) = 3434582 ' Dark Grey
    obj1Colors(3) = 530464  ' Black
    
    Me.clock = 0
    Set Me.sht = ThisWorkbook.Sheets("Gameboy")
    
    ' Initialize the screen
    Dim numPixels As Long: numPixels = Me.screenWidth * Me.screenHeight
    Dim row As Long: row = 0
    Dim col As Long: col = 0
    
    With sht
        For i = 0 To numPixels - 1
            col = i \ Me.screenHeight
            row = i Mod Me.screenHeight
            .Cells(row + 1, col + 1).Interior.color = RGB(239, 239, 239)
        Next i
    End With
    DoEvents
    
End Sub

Public Property Get RegistersLCDC() As Long
    RegistersLCDC = registers.lcdc
End Property

Public Property Get RegistersSTAT() As Long
    RegistersSTAT = registers.stat
End Property

Public Property Get RegistersSCY() As Long
    RegistersSCY = registers.scy
End Property

Public Property Get RegistersSCX() As Long
    RegistersSCX = registers.scx
End Property

Public Property Get RegistersLY() As Long
    RegistersLY = registers.ly
End Property

Public Property Get RegistersLYC() As Long
    RegistersLYC = registers.lyc
End Property

Public Property Get RegistersDMA() As Long
    RegistersDMA = registers.dma
End Property

Public Property Get RegistersBGP() As Long
    RegistersBGP = registers.bgp
End Property

Public Property Get RegistersOBJ0() As Long
    RegistersOBJ0 = registers.obj0
End Property

Public Property Get RegistersOBJ1() As Long
    RegistersOBJ1 = registers.obj1
End Property

Public Property Get RegistersWY() As Long
    RegistersWY = registers.wy
End Property

Public Property Get RegistersWX() As Long
    RegistersWX = registers.wx
End Property

Public Property Get RegistersVBK() As Long
    RegistersVBK = registers.vbk
End Property

Public Property Get RegistersHDMA1() As Long
    RegistersHDMA1 = registers.hdma1
End Property

Public Property Get RegistersHDMA2() As Long
    RegistersHDMA2 = registers.hdma2
End Property

Public Property Get RegistersHDMA3() As Long
    RegistersHDMA3 = registers.hdma3
End Property

Public Property Get RegistersHDMA4() As Long
    RegistersHDMA4 = registers.hdma4
End Property

Public Property Get RegistersHDMA5() As Long
    RegistersHDMA5 = registers.hdma5
End Property

Public Property Get RegistersBGPI() As Long
    RegistersBGPI = registers.bgpi
End Property

Public Property Get RegistersBGPD() As Long
    RegistersBGPD = registers.bgpd
End Property

Public Property Get RegistersOBPI() As Long
    RegistersOBPI = registers.obpi
End Property

Public Property Get RegistersOBPD() As Long
    RegistersOBPD = registers.obpd
End Property

Public Property Let RegisterLCDC(ByVal value As Long)
    registers.lcdc = value
End Property

Public Property Let RegisterSTAT(ByVal value As Long)
    registers.stat = value
End Property

Public Property Let RegisterSCY(ByVal value As Long)
    registers.scy = value
End Property

Public Property Let RegisterSCX(ByVal value As Long)
    registers.scx = value
End Property

Public Property Let RegisterLY(ByVal value As Long)
    registers.ly = value
End Property

Public Property Let RegisterLYC(ByVal value As Long)
    registers.lyc = value
End Property

Public Property Let RegisterDMA(ByVal value As Long)
    registers.dma = value
End Property

Public Property Let RegisterBGP(ByVal value As Long)
    registers.bgp = value
End Property

Public Property Let RegisterOBJ0(ByVal value As Long)
    registers.obj0 = value
End Property

Public Property Let RegisterOBJ1(ByVal value As Long)
    registers.obj1 = value
End Property

Public Property Let RegisterWY(ByVal value As Long)
    registers.wy = value
End Property

Public Property Let RegisterWX(ByVal value As Long)
    registers.wx = value
End Property

Public Property Let RegisterVBK(ByVal value As Long)
    registers.vbk = value
End Property

Public Property Let RegisterHDMA1(ByVal value As Long)
    registers.hdma1 = value
End Property

Public Property Let RegisterHDMA2(ByVal value As Long)
    registers.hdma2 = value
End Property

Public Property Let RegisterHDMA3(ByVal value As Long)
    registers.hdma3 = value
End Property

Public Property Let RegisterHDMA4(ByVal value As Long)
    registers.hdma4 = value
End Property

Public Property Let RegisterHDMA5(ByVal value As Long)
    registers.hdma5 = value
End Property

Public Property Let RegisterBGPI(ByVal value As Long)
    registers.bgpi = value
End Property

Public Property Let RegisterBGPD(ByVal value As Long)
    registers.bgpd = value
End Property

Public Property Let RegisterOBPI(ByVal value As Long)
    registers.obpi = value
End Property

Public Property Let RegisterOBPD(ByVal value As Long)
    registers.obpd = value
End Property

Public Function ReadByte(ByVal address As Long) As Byte
    
    Dim stat As Byte
    Dim ly As Byte

    If address >= 32768 And address <= 40959 Then
        ReadByte = vram((address - 32768) + (registers.vbk * 8192))
        Exit Function
    End If
    
    ' Sprite Attribute Memory
    If address >= 65024 And address <= 65183 Then
        ReadByte = oam(address - 65024)
        Exit Function
    End If
    
    Select Case address
        Case 65344: ReadByte = registers.lcdc: Exit Function
        Case 65345
            stat = registers.stat
            
            ' Bit 7 is unused and always returns 1.
            stat = stat Or 128
            
            ' Bits 0-2 return 0 when LCD is off.
            If Not Me.IsLcdEnabled Then
                stat = stat And (Not 7)
            End If
            
            ReadByte = stat
            Exit Function
        Case 65346: ReadByte = registers.scy: Exit Function
        Case 65347: ReadByte = registers.scx: Exit Function
        Case 65348
            ly = registers.ly
            
            ' When the LCD is off, LY is fixed at 0.
            If Not Me.IsLcdEnabled() Then
                ly = 0
            End If
            ReadByte = ly
            Exit Function
        Case 65349: ReadByte = registers.lyc: Exit Function
        Case 65350: ReadByte = registers.dma: Exit Function
        Case 65351: ReadByte = registers.bgp: Exit Function
        Case 65352: ReadByte = registers.obj0: Exit Function
        Case 65353: ReadByte = registers.obj1: Exit Function
        Case 65354: ReadByte = registers.wy: Exit Function
        Case 65355: ReadByte = registers.wx: Exit Function
        Case 65359: ReadByte = registers.vbk: Exit Function
    End Select
    
    Err.Raise 50000, "Gpu", "GPU: Unknown read at " & address
End Function

Public Sub WriteByte(ByVal address As Long, ByVal wbyte As Byte)
    ' Video RAM
    If address >= 32768 And address <= 40959 Then
        vram((address - 32768) + (registers.vbk * 8192)) = wbyte
        Exit Sub
    End If
    
    ' Sprite Attribute Memory
    If address >= 65024 And address <= 65183 Then
        oam(address - 65024) = wbyte
        Exit Sub
    End If
    
    ' GPU Registers
    Select Case address
        Case 65344: registers.lcdc = wbyte: Exit Sub
        Case 65345: registers.stat = wbyte: Exit Sub
        Case 65346: registers.scy = wbyte: Exit Sub
        Case 65347: registers.scx = wbyte: Exit Sub
        Case 65348: registers.ly = 0: Exit Sub ' Note: any outside write to LY resets the value to 0;
        Case 65349: registers.lyc = wbyte: Exit Sub
        Case 65350
            registers.dma = wbyte
            Me.InitializeOAM_DMA
            Exit Sub
        Case 65351: registers.bgp = wbyte: Exit Sub
        Case 65352: registers.obj0 = wbyte: Exit Sub
        Case 65353: registers.obj1 = wbyte: Exit Sub
        Case 65354: registers.wy = wbyte: Exit Sub
        Case 65355: registers.wx = wbyte: Exit Sub
        Case 65359: registers.vbk = wbyte: Exit Sub
    End Select
    
    Err.Raise 50000, "Gpu", "GPU: Unknown write at " & address & " / value: " & wbyte
End Sub

Public Sub InitializeOAM_DMA()
    Me.oamDmaActive = True
    Me.oamDmaByte = 0
End Sub

Public Function IsLcdEnabled() As Boolean
    IsLcdEnabled = CBool(registers.lcdc And 128)
End Function

Public Function GetLcdMode() As Byte
    GetLcdMode = registers.stat And 3
End Function

Public Sub SetLcdMode(ByVal mode As Byte)
    registers.stat = registers.stat And (Not 3)  ' Clear mode.
    registers.stat = registers.stat Or mode ' Set mode.
End Sub

Public Sub Step(ByVal cycles As Long)
    Dim cyclesToProcess As Integer
    Dim i As Long
    Dim currentMode As Byte
    Dim mode As Integer
    Dim interruptRequested As Boolean
    
    ' Process OAM DMA transfer.
    If Me.oamDmaActive Then
        Dim address As Long: address = Utils.LeftShift(registers.dma, 8)
        cyclesToProcess = cycles / 4
        
        For i = 0 To cyclesToProcess - 1
            Me.system.MMU.WriteByte 65024 + Me.oamDmaByte, Me.system.MMU.ReadByte(address + Me.oamDmaByte)
            Me.oamDmaByte = Me.oamDmaByte + 1
        Next i
        
        If Me.oamDmaByte >= 160 Then
            ' Transfer complete
            Me.oamDmaActive = False
            Me.oamDmaByte = 0
        End If
    End If
    
    If Me.IsLcdEnabled() Then
        Me.clock = (Me.clock + cycles) And 4294967295#
    Else
        Me.SetLcdMode 1
        Exit Sub
    End If
    
    ' Set LCD status.
    currentMode = Me.GetLcdMode()
    mode = Empty
    interruptRequested = False
    
    If registers.ly < 144 Then
        If Me.clock <= 80 Then
            ' Mode 2
            mode = 2
            interruptRequested = CBool(registers.stat And 32)
        ElseIf Me.clock >= 80 And Me.clock < 252 Then
            ' Mode 3
            mode = 3
        ElseIf Me.clock >= 252 And Me.clock < 456 Then
            ' Mode 0
            mode = 0
            interruptRequested = CBool(registers.stat And 8)
        End If
    Else
        ' Mode 1
        mode = 1
        interruptRequested = CBool(registers.stat And 16)
    End If
    
    Me.SetLcdMode mode
    
    ' Request interrupt if modes changed and interrupt requested for LCD stat.
    If currentMode <> mode And interruptRequested Then
        Me.system.RequestInterrupt 1
    End If
    
    ' Check for coincidence flag.
    If registers.ly = registers.lyc Then
        ' Set coincidence flag, lines match.
        registers.stat = registers.stat Or 4
        
        If CBool(registers.stat And 64) Then
            Me.system.RequestInterrupt 1
        End If
    Else
        ' Reset coincidence flag, lines don't match.
        registers.stat = registers.stat And Not 4
    End If
    
    If Me.clock >= 456 Then
    
        If registers.ly < 144 Then
            Me.RenderScanline
        End If
        
        Me.clock = 0
        registers.ly = registers.ly + 1
        
        If registers.ly = 144 Then
            Me.system.RequestInterrupt 0
        ElseIf registers.ly > 153 Then
            registers.ly = 0
            Me.DrawScreen
        End If
    End If
    
End Sub

Public Sub RenderScanline()
    
    Dim pixels() As Pixel: ReDim pixels(0 To 255)
    
    Dim sx As Integer: sx = registers.scx
    Dim sy As Integer: sy = registers.scy
    Dim wx As Integer: wx = registers.wx - 7
    Dim wy As Integer: wy = registers.wy
    Dim ly As Integer: ly = registers.ly
    Dim i As Long
    
    ' Check if window is enabled.
    Dim windowEnabled As Boolean: windowEnabled = CBool(registers.lcdc And 32) And wy <= ly
    Dim tilemapRegion As Long: tilemapRegion = 0
    
    If windowEnabled Then
        If CBool(registers.lcdc And 64) Then
            tilemapRegion = 39936 ' 39936 - 40959
        Else
            tilemapRegion = 38912 ' 38912 - 39935
        End If
    Else
        If CBool(registers.lcdc And 8) Then
            tilemapRegion = 39936 ' 39936 - 40959
        Else
            tilemapRegion = 38912 ' 38912 - 39935
        End If
    End If
    
    ' Get tileset region.
    Dim tilesetRegion As Long: tilesetRegion = 0
    Dim unsignedTiles As Boolean: unsignedTiles = True
    If CBool(registers.lcdc And 16) Then
        tilesetRegion = 32768 ' 32768 - 36863
    Else
        tilesetRegion = 34816 ' 34816 - 38911
        unsignedTiles = False
    End If
    
    ' Load color palette for background.
    Dim bgPalette As Byte: bgPalette = Me.ReadByte(65351)
    
    ' Calculate which scanline we're on.
    Dim yPos As Integer: yPos = 0
    If windowEnabled Then
        yPos = ly - wy
    Else
        yPos = (sy + ly) Mod 256
    End If
    
    ' Generate background / window pixels
    Dim x As Long
    Dim xPos As Integer
    Dim tx As Integer
    Dim ty As Integer
    Dim tileId As Integer
    Dim tileAddress As Long
    Dim px As Integer
    Dim py As Integer
    Dim pixelRow As Integer
    Dim lb As Integer
    Dim hb As Integer
    Dim l As Integer
    Dim h As Integer
    Dim colorCode As Long
    Dim pixelColor As Long
    
    For x = 0 To 159
    
        xPos = (sx + x) Mod 256
        
        If windowEnabled And x >= wx Then
            xPos = x - wx
        End If
        
        tx = Int(xPos / 8) And 255
        ty = Int(yPos / 8) And 255
        tileId = Me.ReadByte(tilemapRegion + (32 * ty + tx))
        
        If Not unsignedTiles Then
            ' Adjust for signed byte.
            If tileId > 127 Then
                tileId = -(((Not tileId) + 1) And 255)
            End If
            tileId = tileId + 128
        End If
        
        ' Find tile pixel data for color.
        tileAddress = tilesetRegion + (tileId * 16)
        px = 0
        py = 0
        If windowEnabled Then
            px = (x - wx) Mod 8
            py = (ly - wy) Mod 8
        Else
            px = (sx + x) Mod 8
            py = (sy + ly) Mod 8
        End If
        pixelRow = py * 2
        lb = Me.ReadByte(tileAddress + pixelRow)
        hb = Me.ReadByte(tileAddress + pixelRow + 1)
        
        l = IIf(CBool(lb And Utils.LeftShift(1, (7 - px))), 1, 0)
        h = IIf(CBool(hb And Utils.LeftShift(1, (7 - px))), 1, 0)
        colorCode = Utils.LeftShift(h, 1) + l
        pixelColor = 0
        
        Select Case colorCode
            Case 0: pixelColor = bgColors(bgPalette And 3)
            Case 1: pixelColor = bgColors(Utils.RightShift(bgPalette, 2) And 3)
            Case 2: pixelColor = bgColors(Utils.RightShift(bgPalette, 4) And 3)
            Case 3: pixelColor = bgColors(Utils.RightShift(bgPalette, 6) And 3)
        End Select
        
        Set pixels(x) = New Pixel
        pixels(x).code = colorCode
        pixels(x).color = pixelColor
        pixels(x).pixelType = "BG"
        
    Next x
    
    ' Load sprites
    Dim spriteId As Long
    Dim renderedSprites As Long
    Dim height As Long
    Dim Sprite As Sprite
    Dim palette() As Long
    Dim Pixel As Long
    Dim currentPixel As Pixel
    
    If CBool(registers.lcdc And 2) Then
        renderedSprites = 0
        
        For spriteId = 0 To 39
            height = IIf(CBool(registers.lcdc And 4), 16, 8)
            Set Sprite = Me.GetSprite(spriteId, height)
            
            If ly >= Sprite.y And ly < (Sprite.y + height) Then
                For tx = 0 To 7
                    ' Load color palette for background.
                    If Sprite.paletteId = 0 Then
                        ReDim palette(LBound(obj0Colors) To UBound(obj0Colors))
                        For i = LBound(obj0Colors) To UBound(obj0Colors)
                            palette(i) = obj0Colors(i)
                        Next i
                    Else
                        ReDim palette(LBound(obj1Colors) To UBound(obj1Colors))
                        For i = LBound(obj1Colors) To UBound(obj1Colors)
                            palette(i) = obj1Colors(i)
                        Next i
                    End If
                    py = ly - Sprite.y
                    
                    If Sprite.yflip Then
                        py = py - (height - 1)
                        py = py * -1
                    End If
                    
                    ' Find tile pixel data for color.
                    px = tx
                    If Sprite.xflip Then
                        px = px - 7
                        px = px * -1
                    End If
                    
                    colorCode = Sprite.GetPixelAt(py Mod height, px Mod 8)
                    If colorCode = 0 Then
                        GoTo CONTINUE_LOOP
                    End If
                    
                    pixelColor = 0
                    
                    Select Case colorCode
                        Case 0: pixelColor = palette(Sprite.palette And 3)
                        Case 1: pixelColor = palette(Utils.RightShift(Sprite.palette, 2) And 3)
                        Case 2: pixelColor = palette(Utils.RightShift(Sprite.palette, 4) And 3)
                        Case 3: pixelColor = palette(Utils.RightShift(Sprite.palette, 6) And 3)
                    End Select
                    
                    Pixel = Sprite.x + tx
                    
                    If Pixel < 0 Then
                        GoTo CONTINUE_LOOP
                    End If
                    
                    If Sprite.priority = 0 Then
                        ' Priority 0: sprite above background.
                        Set pixels(Pixel) = New Pixel
                        pixels(Pixel).code = colorCode
                        pixels(Pixel).color = pixelColor
                        pixels(Pixel).pixelType = "OBJ"
                        pixels(Pixel).objId = spriteId
                    ElseIf Sprite.priority = 1 Then
                        Set currentPixel = pixels(Pixel)
                        If Not currentPixel Is Nothing Then
                            If currentPixel.pixelType = "BG" And Not currentPixel.code = (bgPalette And 3) Then
                                GoTo CONTINUE_LOOP
                            End If
                        End If
                        
                        Set pixels(Pixel) = New Pixel
                        pixels(Pixel).code = colorCode
                        pixels(Pixel).color = pixelColor
                        pixels(Pixel).pixelType = "OBJ"
                        pixels(Pixel).objId = spriteId
                    Else
                        Err.Raise 50000, "Gpu", "Invalid sprite priority"
                    End If
                    
CONTINUE_LOOP:
                Next tx
                
                renderedSprites = renderedSprites + 1
                
                ' Limit 10 sprites per line.
                If renderedSprites = 10 Then
                    GoTo BREAK_LOOP
                End If
            End If
        Next spriteId
BREAK_LOOP:
    End If
    
    For i = LBound(pixels) To UBound(pixels)
        Set frameBuffer(ly, i) = pixels(i)
    Next i
End Sub

Public Sub DrawScreen()
    Dim Pixel As Pixel
    
    Dim numPixels As Long: numPixels = Me.screenWidth * Me.screenHeight
    Dim row As Long: row = 0
    Dim col As Long: col = 0
    
    Dim i As Long
    If Me.frameCounter Mod 2 = 0 Then
        With sht
            For i = 0 To numPixels - 1
                col = i \ Me.screenHeight
                row = i Mod Me.screenHeight
                Set Pixel = frameBuffer(row, col)
                .Cells(row + 1, col + 1).Interior.color = RGB( _
                    Utils.RightShift(Pixel.color, 16) And 255, _
                    Utils.RightShift(Pixel.color, 8) And 255, _
                    Pixel.color And 255)
            Next i
        End With

        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False
        
        If GetAsyncKeyState(VK_RIGHT) Then
            Me.system.Joypad.ButtonPressed 0
        Else
            Me.system.Joypad.ButtonReleased 0
        End If
        If GetAsyncKeyState(VK_LEFT) Then
            Me.system.Joypad.ButtonPressed 1
        Else
            Me.system.Joypad.ButtonReleased 1
        End If
        If GetAsyncKeyState(VK_UP) Then
            Me.system.Joypad.ButtonPressed 2
        Else
            Me.system.Joypad.ButtonReleased 2
        End If
        If GetAsyncKeyState(VK_DOWN) Then
            Me.system.Joypad.ButtonPressed 3
        Else
            Me.system.Joypad.ButtonReleased 3
        End If
        If GetAsyncKeyState(VK_CONTROL) Then
            Me.system.Joypad.ButtonPressed 4
        Else
            Me.system.Joypad.ButtonReleased 4
        End If
        If GetAsyncKeyState(VK_MENU) Then
            Me.system.Joypad.ButtonPressed 5
        Else
            Me.system.Joypad.ButtonReleased 5
        End If
        If GetAsyncKeyState(VK_SPACE) Then
            Me.system.Joypad.ButtonPressed 6
        Else
            Me.system.Joypad.ButtonReleased 6
        End If
        If GetAsyncKeyState(VK_RETURN) Then
            Me.system.Joypad.ButtonPressed 7
        Else
            Me.system.Joypad.ButtonReleased 7
        End If
        
        Me.frameCounter = 0
    End If
    Me.frameCounter = Me.frameCounter + 1
End Sub

Public Function GetSprite(ByVal spriteId As Long, ByVal height As Long) As Sprite
    Dim spriteAddress As Long: spriteAddress = CLng(65024) + (spriteId * 4)
    Dim spriteY As Integer: spriteY = Me.ReadByte(spriteAddress) - 16 ' Offset for display window.
    Dim spriteX As Integer: spriteX = Me.ReadByte(spriteAddress + 1) - 8 ' Offset for display window.
    Dim tileId As Long: tileId = Me.ReadByte(spriteAddress + 2)
    Dim attributes As Long: attributes = Me.ReadByte(spriteAddress + 3)
    
    ' TODO: Get sprite priority.
    
    Dim pixels() As Integer: ReDim pixels(0 To height - 1, 0 To 7)
    Dim tileAddress As Long: tileAddress = 32768 + (tileId * 16)
    
    Dim y As Long
    Dim x As Long
    Dim lb As Long
    Dim hb As Long
    Dim l As Integer
    Dim h As Integer
    Dim color As Integer
    For y = 0 To height - 1
        lb = Me.ReadByte(tileAddress + (y * 2))
        hb = Me.ReadByte(tileAddress + (y * 2) + 1)
        
        For x = 0 To 7
            l = IIf(CBool(lb And Utils.LeftShift(1, (7 - x))), 1, 0)
            h = IIf(CBool(hb And Utils.LeftShift(1, (7 - x))), 1, 0)
            color = Utils.LeftShift(h, 1) + l
            
            pixels(y, x) = color
        Next x
    Next y
    
    Dim returnObj As New Sprite
    
    With returnObj
        .id = spriteId
        .address = spriteAddress
        .tileId = tileId
        .x = spriteX
        .y = spriteY
        .xflip = CBool(attributes And 32)
        .yflip = CBool(attributes And 64)
        .RedimPixels LBound(pixels, 1), UBound(pixels, 1), LBound(pixels, 2), UBound(pixels, 2)
        For y = LBound(pixels, 1) To UBound(pixels, 1)
            For x = LBound(pixels, 2) To UBound(pixels, 2)
                .SetPixelAt y, x, pixels(y, x)
            Next x
        Next y
        .palette = IIf(attributes And 16, Me.ReadByte(65353), Me.ReadByte(65352))
        .paletteId = attributes And 16
        .priority = Utils.RightShift(attributes, 7)
    End With
    Set GetSprite = returnObj
End Function
