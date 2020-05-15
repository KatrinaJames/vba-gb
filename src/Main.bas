Attribute VB_Name = "Main"
Option Explicit
Option Base 0

Public gb As Gameboy
Public runTo As Long

Public Sub Main()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With

    Dim Gameboy As New Gameboy
    
    With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .Filters.ADD "Gameboy Files", "*.gb", 1
        'Show the dialog box
        .Show
        
        'Store in fullpath variable
        If .SelectedItems.Count <> 1 Then
            Exit Sub
        End If
        
        Dim fullpath As String: fullpath = .SelectedItems.Item(1)
    End With
    
    If InStr(fullpath, ".gb") = 0 Then
        Exit Sub
    End If
    
    Gameboy.LoadFile fullpath
    Gameboy.Start
    
End Sub

Public Sub RunFrame(ByRef game As Gameboy, Optional ByVal runToInstruction As Long)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With
    
    ThisWorkbook.Sheets("Gameboy").DisplayPageBreaks = False
    
    Set gb = game
    runTo = runToInstruction
    
    If gb.shouldStop Then
        Exit Sub
    End If
    gb.Frame runTo
    Dim eventTimer As Date: eventTimer = Now
    Application.OnTime eventTimer, "'Main.RunFrame gb, runTo'"
End Sub
