Attribute VB_Name = "ModelAnalyser"
Option Explicit

Private Model As StaadModel

Public Sub AnalyseModel(ByVal StaadModel As StaadModel)

    'Dim currentNode As Node
    'Dim currentBeam As Beam
    Dim currentNode As Variant
    Dim currentBeam As Variant
    Dim supportNodes As Collection
    
    Set Model = StaadModel
    
    '' Classify vertical members connected to support nodes
    Set supportNodes = New Collection
    For Each currentNode In Model.Nodes.Items
        If currentNode.IsSupport Then
            Call CheckSupportNode(currentNode)
        End If
    Next
    
    For Each currentBeam In Model.beams.Items
        If Not currentBeam.BeamType = MODELBEAMTYPE.UNKNOWN Then
            GoTo SkipToNext
        End If
        
        If currentBeam.Spec = STAADBEAMSPEC.UNSPECIFIED Then
            If currentBeam.IsParallelToY Then
                currentBeam.BeamType = MODELBEAMTYPE.Post
            Else
                currentBeam.BeamType = MODELBEAMTYPE.Beam
            End If
        Else
            currentBeam.BeamType = MODELBEAMTYPE.Brace
        End If
SkipToNext:
    Next
    
End Sub

Private Function CheckSupportNode(ByVal SupportNode As Node) As Collection
    
    Dim currentBeam As Variant
    Dim columnBeam As Variant
    Dim columnBeams As Collection
    
    Set columnBeams = New Collection
    For Each currentBeam In SupportNode.ConnectedBeams.Items
        If currentBeam.IsParallelToY Then
            For Each columnBeam In GatherParallelBeams(currentBeam)
                columnBeam.BeamType = MODELBEAMTYPE.Column
                Call columnBeams.Add(columnBeam)
            Next
        End If
    Next
    
    Set CheckSupportNode = columnBeams
    
End Function

Private Function GatherParallelBeams(ByVal currentBeam As Beam, Optional ByVal PreviousNode As Node = Nothing, Optional ByVal MoveDownStream = False) As Collection
    
    Dim currentNode As Node
    Dim nextBeam As Beam
    Dim connectedBeam As Variant
    Dim connectedParallelBeams As Collection
    Dim parallelBeams As Collection
    
    Set parallelBeams = New Collection
    If MoveDownStream Then
        Call parallelBeams.Add(currentBeam)
    End If
    
    '' Determine the correct next node
    If MoveDownStream Then
        Set currentNode = currentBeam.Node2
        If Not PreviousNode Is Nothing And PreviousNode Is currentBeam.Node2 Then
            currentNode = currentBeam.Node1
        End If
    Else
        Set currentNode = currentBeam.Node1
        If Not PreviousNode Is Nothing And PreviousNode Is currentBeam.Node1 Then
            currentNode = currentBeam.Node2
        End If
    End If
    
    '' Get connected parallel beams and continue the chain from there
    Set connectedParallelBeams = New Collection
    For Each connectedBeam In currentNode.ConnectedBeams.Items
        If Not connectedBeam Is currentBeam And currentBeam.DetermineRelationship(connectedBeam) = BEAMRELATIONSHIP.parallel Then
            Call connectedParallelBeams.Add(connectedBeam)
        End If
    Next
    If connectedParallelBeams.count > 0 Then
        For Each connectedBeam In GatherParallelBeams(connectedParallelBeams(1), currentNode, MoveDownStream)
            Call parallelBeams.Add(connectedBeam)
        Next
    Else
        If Not MoveDownStream Then
            For Each connectedBeam In GatherParallelBeams(currentBeam, currentNode, True)
                Call parallelBeams.Add(connectedBeam)
            Next
        End If
    End If
    
    Set GatherParallelBeams = parallelBeams
    
End Function

