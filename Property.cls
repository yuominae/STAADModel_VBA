VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public ID As Long
Public Width As Double
Public Depth As Double
Public FlangeThickness As Double
Public WebThickness As Double
Public Ax As Double
Public Ay As Double
Public Az As Double
Public Ix As Double
Public Iy As Double
Public Iz As Double
Public AssignedBeams As Dictionary

Private Sub Class_Initialize()
    
    Set Me.AssignedBeams = New Dictionary
    
End Sub
