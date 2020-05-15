VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RamLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public wasSuccessful As Boolean
Private ram() As Byte

Public Sub Class_Initialize()

End Sub

Public Function GetRamByteAt(ByVal index As Long) As Byte
    GetRomByteAt = ram(index)
End Function

Public Sub SetRamByteAt(ByVal index As Long, ByVal value As Byte)
    ram(index) = value
End Sub

Public Sub RedimRam(ByVal lowerBound As Long, ByVal upperBound As Long)
    ReDim ram(lowerBound To upperBound)
End Sub

Public Function RamSize() As Long
    RamSize = UBound(ram)
End Function
