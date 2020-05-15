VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public id As Long
Public address As Long
Public tileId As Long
Public x As Long
Public y As Long
Public xflip As Boolean
Public yflip As Boolean
Private pixels() As Integer
Public palette As Byte
Public paletteId As Long
Public priority As Integer

Public Sub Class_Initialize()

End Sub

Public Sub RedimPixels( _
    ByVal lowerBound1 As Long, _
    ByVal upperBound1 As Long, _
    ByVal lowerBound2 As Long, _
    ByVal upperBound2 As Long)
        
    ReDim pixels(lowerBound1 To upperBound1, lowerBound2 To upperBound2)
End Sub

Public Function GetPixelAt( _
    ByVal index1 As Long, _
    ByVal index2 As Long) As Integer
    
    GetPixelAt = pixels(index1, index2)
End Function

Public Sub SetPixelAt( _
    ByVal index1 As Long, _
    ByVal index2 As Long, _
    ByVal value As Integer)
    
    pixels(index1, index2) = value
End Sub
