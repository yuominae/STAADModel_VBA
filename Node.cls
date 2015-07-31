VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Private fields
Private pNodeSupport As Support

'' Public fields
Public ID As Long
Public X As Double
Public Y As Double
Public Z As Double
Public ConnectedBeams As Dictionary

'' Public properties
Public Property Get IsSupport() As Boolean
    IsSupport = False
    If Not Me.NodeSupport Is Nothing Then
        IsSupport = True
    End If
End Property

Public Property Get NodeSupport() As Support
    Set NodeSupport = pNodeSupport
End Property
Public Property Let NodeSupport(ByVal value As Support)
    Set pNodeSupport = value
    Call pNodeSupport.Nodes.Add(Me.ID, Me)
End Property

'' Methods
Private Sub Class_Initialize()
    
    Set Me.ConnectedBeams = New Dictionary
    
End Sub

Public Sub Initialise(ByVal ID As Long, ByVal X As Double, ByVal Y As Double, ByVal Z As Double)

    Me.ID = ID
    Me.X = X
    Me.Y = Y
    Me.Z = Z

End Sub
