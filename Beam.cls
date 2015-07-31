VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Beam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Private fields
Private pNode1 As Node
Private pNode2 As Node
Private pSectionMaterial As Material
Private pSectionProperty As Property

'' Public fields
Public ID As Long
Public Spec As STAADBEAMSPEC
Public BetaAngle As Double
Public Release1 As Release
Public Release2 As Release
Public Forces1 As Dictionary
Public Forces2 As Dictionary
Public BeamType As MODELBEAMTYPE

'' Public properties
Public Property Get DX() As Double
    If DX = 0 Then
        DX = Me.Node2.X - Me.Node1.X
    End If
End Property

Public Property Get DY() As Double
    If DY = 0 Then
        DY = Me.Node2.Y - Me.Node1.Y
    End If
End Property

Public Property Get DZ() As Double
    If DZ = 0 Then
        DZ = Me.Node2.Z - Me.Node1.Z
    End If
End Property

Public Property Get Length() As Double
    If Length = 0 Then
        Length = (Me.DX ^ 2 + Me.DY ^ 2 + Me.DZ ^ 2) ^ 0.5
    End If
End Property

Public Property Get IsParallelToX() As Boolean
    IsParallelToX = (Me.DX <> 0 And Me.DY = 0 And Me.DZ = 0)
End Property

Public Property Get IsParallelToY() As Boolean
    IsParallelToY = (Me.DX = 0 And Me.DY <> 0 And Me.DZ = 0)
End Property

Public Property Get IsParallelToZ() As Boolean
    IsParallelToZ = (Me.DX = 0 And Me.DY = 0 And Me.DZ <> 0)
End Property

Public Property Get HasReleases() As Boolean
    HasReleases = (Me.Release1.IsReleased Or Me.Release2.IsReleased)
End Property

Public Property Get HasResults() As Boolean
    HasResults = (Me.Forces1.count > 0 Or Me.Forces2.count > 0)
End Property

Public Property Get Node1() As Node
    If Node1 Is Nothing Then
        Set Node1 = pNode1
    End If
End Property
Private Property Let Node1(value As Node)
    Set pNode1 = value
    Call pNode1.ConnectedBeams.Add(Me.ID, Me)
End Property

Public Property Get Node2() As Node
    If Node2 Is Nothing Then
        Set Node2 = pNode2
    End If
End Property
Private Property Let Node2(value As Node)
    Set pNode2 = value
    Call pNode2.ConnectedBeams.Add(Me.ID, Me)
End Property

Public Property Get ConnectedBeams1() As Dictionary
    If ConnectedBeams1 Is Nothing Then
        Set ConnectedBeams1 = GetConnectedBeams(Me.Node1)
    End If
End Property

Public Property Get ConnectedBeams2() As Dictionary
    If ConnectedBeams2 Is Nothing Then
        Set ConnectedBeams2 = GetConnectedBeams(Me.Node1)
    End If
End Property

Public Property Get SectionMaterial() As Material
    If SectionMaterial Is Nothing Then
        Set SectionMaterial = pSectionMaterial
    End If
End Property
Public Property Let SectionMaterial(ByVal value As Material)
    Set pSectionMaterial = value
    Call pSectionMaterial.AssignedBeams.Add(Me.ID, Me)
End Property

Public Property Get SectionProperty() As Property
    If SectionProperty Is Nothing Then
        Set SectionProperty = pSectionProperty
    End If
End Property
Public Property Let SectionProperty(ByVal value As Property)
    Set pSectionProperty = value
    Call pSectionProperty.AssignedBeams.Add(Me.ID, Me)
End Property

Public Property Get AxialVector() As Vector
    If AxialVector Is Nothing Then
        Set AxialVector = New Vector
        Call AxialVector.Initialise(Me.Node1.X, Me.Node1.Y, Me.Node1.Z, Me.Node2.X, Me.Node2.Y, Me.Node2.Z)
    End If
End Property

'' Methods
Private Sub Class_Initialize()
    
    Set Me.Release1 = New Release
    Set Me.Release2 = New Release
    Set Me.Forces1 = New Dictionary
    Set Me.Forces2 = New Dictionary
    Me.BeamType = MODELBEAMTYPE.UNKNOWN
    
End Sub

Public Sub Initialise(ByVal ID As Long, ByVal StartNode As Node, ByVal EndNode As Node)
    
    Me.ID = ID
    Node1 = StartNode
    Node2 = EndNode
    
End Sub

Public Function GetAngleToBeam(ByVal OtherBeam As Beam) As Double
    
    GetAngleToBeam = Me.AxialVector.angle(OtherBeam.AxialVector)
    
End Function

Public Function DetermineRelationship(ByVal OtherBeam As Beam) As BEAMRELATIONSHIP

    Dim angle As Double
    Dim output As BEAMRELATIONSHIP
    
    angle = Math.Abs((Me.GetAngleToBeam(OtherBeam) + 90) Mod 180 - 90)
    
    If angle = 0 Then
        output = BEAMRELATIONSHIP.parallel
    ElseIf angle = 90 Then
        output = BEAMRELATIONSHIP.PERPENDICULAR
    Else
        output = BEAMRELATIONSHIP.UNKNOWN
    End If
    
    DetermineRelationship = output

End Function

Private Function GetConnectedBeams(ByVal Node As Node) As Dictionary

    Dim thisBeam As Variant
    Dim output As Dictionary
    
    Set output = New Dictionary
    For Each thisBeam In Node.ConnectedBeams
        If Not thisBeam Is Me Then
            Call output.Add(thisBeam.ID, thisBeam)
        End If
    Next
    
    Set GetConnectedBeams = output

End Function
