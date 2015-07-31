VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Public fields
Public X As Double
Public Y As Double
Public Z As Double

'' Public properties
Public Property Get Magnitude() As Double
    If Magnitude = 0 Then
        Magnitude = (Me.X ^ 2 + Me.Y ^ 2 + Me.Z ^ 2) ^ 0.5
    End If
End Property

'' Public methods
Public Sub Initialise(ByVal X1 As Double, ByVal Y1 As Double, ByVal Z1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal Z2 As Double)

    Me.X = X2 - X1
    Me.Y = Y2 - Y1
    Me.Z = Z2 - Z1

End Sub

Public Function Dot(ByVal OtherVector As Vector) As Double

    Dot = Me.X * OtherVector.X + Me.Y * OtherVector.Y + Me.Z * OtherVector.Z

End Function

Public Function angle(ByVal OtherVector As Vector)

    angle = Round(Helpers.Acos(Me.Dot(OtherVector) / (Me.Magnitude * OtherVector.Magnitude)), 6)

End Function
