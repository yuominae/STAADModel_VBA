VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Release"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Public fields
'' Force releases
Public Fx As Boolean
Public Fy As Boolean
Public Fz As Boolean
'' Moment releases
Public Mx As Boolean
Public My As Boolean
Public Mz As Boolean

'' Public properties
Public Property Get IsReleased() As Boolean

    IsReleased = False
    If Me.Fx Or Me.Fy Or Me.Fz Or Me.Mx Or Me.My Or Me.Mz Then
        IsReleased = True
    End If

End Property

