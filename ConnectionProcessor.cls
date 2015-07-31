VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Public events
Public Event StatusUpdate(ByVal StatusMessage As String)

Public Model As StaadModel
'' LIst of material names to ignore. Connecting members cannot be made of this material (usually concrete)
Public MaterialsToIgnore As Collection
Public PropertiesToIgnore As Collection
Public Connections As Dictionary

'''
''' Initialise the ConnectionProcessor using the specified StaadModel
'''
Public Sub Initialise(ByVal Model As StaadModel)

    Set Me.Model = Model
    Set Me.MaterialsToIgnore = New Collection
    Set Me.PropertiesToIgnore = New Collection
    Set Me.Connections = New Dictionary

End Sub

'''
''' Save all connection data to the specified file
'''
Public Function SaveConnectionData(ByVal TargetFilePath As String, Optional ByVal TargetConnections As Collection = Nothing) As Boolean
    
    Dim sb As StringBuilder
    Dim currentConnection As Variant
    
    On Error GoTo SaveError
    Set sb = New StringBuilder
    
    '' Check if any specific connections have been specified by the user
    If TargetConnections Is Nothing Then
        Set TargetConnections = New Collection
        For Each currentConnection In Me.Connections.Items
            Call TargetConnections.Add(currentConnection)
        Next
    End If
    
    
    '' Convert all connections to their string representation
    For Each currentConnection In TargetConnections
        Call sb.AppendLine(ConvertConnectionToString(currentConnection))
    Next
    
    '' Output the string to file
    With New FileSystemObject
        With .OpenTextFile(TargetFilePath, ForWriting, Not .FileExists(TargetFilePath))
            Call .Write(sb.Text)
        End With
    End With
    
    SaveConnectionData = True
SaveError:
    
End Function

'''
''' Load connection data from the specified file
'''
Public Function LoadConnectionData(ByVal TargetFilePath As String) As Boolean
    
    Dim added As Long
    Dim skipped As Long
    Dim currentConnection As Connection
    
    Call Me.Connections.RemoveAll
    
    On Error GoTo LoadError
    With New FileSystemObject
        If Not .FileExists(TargetFilePath) Then
            Exit Function
        End If
        With .OpenTextFile(TargetFilePath, ForReading)
            Do While Not .AtEndOfStream
                '' Parse the current line and try to instantiate a new connection from it
                Set currentConnection = ParseSavedConnectionString(.ReadLine)
                '' If the connection is nothing then it means it didn't exist
                If Not currentConnection Is Nothing Then
                    Call Me.Connections.Add(currentConnection.ConnectionNode.ID, currentConnection)
                    added = added + 1
                Else
                    skipped = skipped + 1
                End If
                '' inform the user of the progress
                RaiseEvent StatusUpdate("Loading entries: " & added + skipped & " (" & added & " added, " & skipped & " skipped")
            Loop
        End With
    End With
    
    LoadConnectionData = True
LoadError:
    
    '' inform the user of the final result
    If LoadConnectionData Then
        RaiseEvent StatusUpdate("Successfully loaded " & added & " connections (" & skipped & " connections skipped)")
    Else
        RaiseEvent StatusUpdate("An error occured whilst loading the connections")
    End If

End Function

'''
''' Auto generate connections for the current model using the default rules
'''
Public Function AutoGenerateConnections(Optional ByVal nodesToCheck As Collection = Nothing) As Dictionary

    Dim i As Long
    'Dim currentNode As Node
    Dim currentNode As Variant
    Dim newConnection As Connection
    
    '' Remove any existing connections
    Call Me.Connections.RemoveAll
    
    If nodesToCheck Is Nothing Then
        Set nodesToCheck = New Collection
        For Each currentNode In Me.Model.Nodes.Items
            Call nodesToCheck.Add(currentNode)
        Next
    End If
    
    '' Loop through the nodes and analyse each in turn
    '' Only nodes with at least two connecting beams are considered
    For Each currentNode In nodesToCheck
        If currentNode.ConnectedBeams.count >= 2 Then
            Set newConnection = AnalyseNode(currentNode)
            If Not newConnection Is Nothing Then
                Call Me.Connections.Add(newConnection.ConnectionNode.ID, newConnection)
            End If
        End If
        '' Update on analysis status
        i = i + 1
        RaiseEvent StatusUpdate("Checking ConnectionNode " & i & " of " & Me.Model.Nodes.count & " (Node " & currentNode.ID & ")...")
    Next
    
    '' Inform the user that the analysis has completed
    RaiseEvent StatusUpdate("Connection generation completed")
    Set AutoGenerateConnections = Me.Connections

End Function

'''
''' Select the
'''
Public Sub SelectConnections()
    
    Dim nodeCount As Long
    Dim beamCount As Long
    Dim nodeIDs() As Long
    Dim beamIDs() As Long
    Dim currentConnection As Variant
    Dim currentBeam As Variant
    
    ReDim nodeIDs(Me.Model.Nodes.count - 1)
    ReDim beamIDs(Me.Model.beams.count - 1)
    
    '' Get node and beam ids
    For Each currentConnection In Me.Connections.Items
        nodeIDs(nodeCount) = currentConnection.ConnectionNode.ID
        For Each currentBeam In currentConnection.ConnectingBeams
            beamIDs(beamCount) = currentBeam.ID
            beamCount = beamCount + 1
        Next
        nodeCount = nodeCount + 1
    Next
    
    '' Deselect current selection
    Call Me.Model.Geometry.ClearMemberSelection
    Call Me.Model.Geometry.ClearNodeSelection
    '' Select beams
    RaiseEvent StatusUpdate("Selecting beams...")
    Call Me.Model.Geometry.SelectMultipleBeams(beamIDs)
    '' Select nodes
    RaiseEvent StatusUpdate("Selecting nodes...")
    Call Me.Model.Geometry.SelectMultipleNodes(nodeIDs)
    
    RaiseEvent StatusUpdate("Selection completed.")
    
End Sub

'''
''' Extract forces for all connections loaded in the connection processor for the specified load cases
'''
Public Function ExtractConnectionForces(Optional ByVal TargetLoadCases As Collection = Nothing, Optional ByVal TargetConnections As Collection = Nothing) As Boolean
    
    Dim progress As Integer
    Dim currentConnection As Variant
    Dim currentBeam As Variant
    Dim currentLoadCase As Variant
    
    '' Verify that load cases have been specified
    If TargetLoadCases Is Nothing Then
        Set TargetLoadCases = New Collection
        For Each currentLoadCase In Me.Model.LoadCases.Items
            Call TargetLoadCases.Add(currentLoadCase)
        Next
        For Each currentLoadCase In Me.Model.LoadCombinations.Items
            Call TargetLoadCases.Add(currentLoadCase)
        Next
    End If
    
    '' Verify that connections have been specified
    If TargetConnections Is Nothing Then
        Set TargetConnections = New Collection
        For Each currentConnection In Me.Connections.Items
            Call TargetConnections.Add(currentConnection)
        Next
    End If
    
    '' Go through each connection in the processor and extract forces
    'On Error GoTo ExitPoint
    For Each currentConnection In TargetConnections
        progress = progress + 1
        '' Go through each beam in the connecting beams and extract forces at the node of the connection
        For Each currentBeam In currentConnection.ConnectingBeams
            With IIf(currentConnection.ConnectionNode Is currentBeam.Node1, currentBeam.Forces1, currentBeam.Forces2)
                Call .RemoveAll
            End With
            If Not Me.Model.LoadBeamForcesAtNode(currentBeam, currentConnection.ConnectionNode, TargetLoadCases) Then
                GoTo ExitPoint
            End If
            '' Notify the user of the progress
            RaiseEvent StatusUpdate(Format(progress / Me.Connections.count, "0%") & " Extracting forces at node " & currentConnection.ConnectionNode.ID & " for beam " & currentBeam.ID)
        Next
    Next
    
    ExtractConnectionForces = True

ExitPoint:
    '' Notify the user of the outcome of the extraction process
    If ExtractConnectionForces Then
        RaiseEvent StatusUpdate("Force extraction complete")
    Else
        RaiseEvent StatusUpdate("Error during force extraction")
    End If
    
End Function

Private Function AnalyseNode(ByVal NodeToCheck As Node) As Connection
    
    Dim i As Integer
    'Dim currentBeam As Beam
    Dim currentBeam As Variant
    Dim currentMaterial As Variant
    Dim pinnedBeams As Collection
    Dim fixedBeams As Collection
    Dim newConnection As Connection
        
    Set newConnection = New Connection
    
    Set pinnedBeams = New Collection
    Set fixedBeams = New Collection
    
    '' Gather all pinned beams first
    For Each currentBeam In NodeToCheck.ConnectedBeams.Items
        
        If currentBeam.Spec <> STAADBEAMSPEC.UNSPECIFIED Then
            Call pinnedBeams.Add(currentBeam)
            GoTo SkipNext
        End If
        
        If currentBeam.HasReleases Then
            If NodeToCheck Is currentBeam.Node1 Then
                If currentBeam.Release1.My And currentBeam.Release1.Mz Then
                    Call pinnedBeams.Add(currentBeam)
                    GoTo SkipNext
                End If
            Else
                If currentBeam.Release2.My And currentBeam.Release2.Mz Then
                    Call pinnedBeams.Add(currentBeam)
                    GoTo SkipNext
                End If
            End If
        End If
        
        Call fixedBeams.Add(currentBeam)
SkipNext:
    Next
    
    '' Resolve fixed beam intersections if required
    Select Case fixedBeams.count
        Case 1
            Call newConnection.Initialise(NodeToCheck, fixedBeams)
        Case Is > 1
            Set newConnection = ResolveBeamIntersection(NodeToCheck, fixedBeams)
        Case Else
            '' If all beams are released at the node then we have a mechanism...
            Set newConnection = Nothing
    End Select
    
    '' Add the pinned beams to the connection, but only if they are not made of a material to ignore
    For Each currentBeam In pinnedBeams
        Call newConnection.ConnectingBeams.Add(currentBeam)
    Next
    
    '' Exclude the connecting beams that are made of a material that was to be ignored
    If Me.MaterialsToIgnore.count > 0 And newConnection.ConnectingBeams.count > 0 Then
        For i = newConnection.ConnectingBeams.count To 1 Step -1
            For Each currentMaterial In Me.MaterialsToIgnore
                If StrComp(newConnection.ConnectingBeams(i).SectionMaterial.Name, currentMaterial, vbBinaryCompare) = 0 Then
                    Call newConnection.ConnectingBeams.Remove(i)
                End If
            Next
        Next
    End If
    
    '' Verify that the connection actually has any beams connecting to it or set it to nothing as this is not a connection
    If Not newConnection Is Nothing Then
        If newConnection.ConnectingBeams.count = 0 Then
            Set newConnection = Nothing
        End If
    End If
    
    Set AnalyseNode = newConnection
    
End Function

Private Function ResolveBeamIntersection(ByVal NodeToCheck As Node, ByVal BeamsToCheck As Collection) As Connection
        
        Dim i As Integer
        Dim beamAdded As Boolean
        Dim currentBeam1 As Variant
        Dim currentBeam2 As Variant
        Dim primaryBeams As Collection
        Dim primaryBeams_2 As Collection
        Dim secondaryBeams As Collection
        Dim newConnection As Connection
        
        Set primaryBeams = New Collection
        Set primaryBeams_2 = New Collection
        Set secondaryBeams = New Collection
        
        Set primaryBeams = GetContinuousBeams(BeamsToCheck)
        
        '' Check what was found
        Select Case primaryBeams.count
            Case 0
                '' if no beams were found then check for vertical members
                Set primaryBeams = New Collection
                For Each currentBeam1 In BeamsToCheck
                    If currentBeam1.IsParallelToY Then
                        Call primaryBeams.Add(currentBeam1)
                    End If
                Next
                '' If there are no vertical members then take the biggest beam
                If primaryBeams.count = 0 Then
                    Call primaryBeams.Add(SelectBiggestBeam(BeamsToCheck))
                End If
            Case Is > 2
                '' If more than two beams were found, check for columns
                Set primaryBeams_2 = New Collection
                For Each currentBeam1 In primaryBeams
                    If currentBeam1.BeamType = MODELBEAMTYPE.Column Then
                        Call primaryBeams_2.Add(currentBeam1)
                    End If
                Next
                '' if no columns were found then simply take the biggest beam and any beam that might be parallel to it
                If primaryBeams_2.count = 0 Then
                    Call primaryBeams_2.Add(SelectBiggestBeam(primaryBeams))
                    For Each currentBeam1 In primaryBeams
                        If currentBeam1.DetermineRelationship(primaryBeams_2(1)) = BEAMRELATIONSHIP.parallel And currentBeam1.SectionProperty.Name = primaryBeams_2(1).SectionProperty.Name Then
                            Call primaryBeams_2.Add(currentBeam1)
                            Exit For
                        End If
                    Next
                End If
                Set primaryBeams = primaryBeams_2
        End Select
        
        '' Initialise the new connection wit the primary beams
        Set newConnection = New Connection
        Call newConnection.Initialise(NodeToCheck, primaryBeams)
        
        '' Split out the secondary beams
        If primaryBeams.count < BeamsToCheck.count Then
            For Each currentBeam1 In BeamsToCheck
                Call secondaryBeams.Add(currentBeam1)
            Next
            For i = secondaryBeams.count To 1 Step -1
                For Each currentBeam1 In primaryBeams
                    If secondaryBeams(i) Is currentBeam1 Then
                        Call secondaryBeams.Remove(i)
                        Exit For
                    End If
                Next
            Next
        End If
        
        If secondaryBeams.count > 0 Then
            Set newConnection.ConnectingBeams = secondaryBeams
        End If
        
        Set ResolveBeamIntersection = newConnection
    
End Function

Private Function GetContinuousBeams(ByVal TargetBeams) As Collection
    
    Dim beamAdded As Boolean
    Dim currentBeam1 As Variant
    Dim currentBeam2 As Variant
    Dim output As Dictionary
    
    Set output = New Dictionary
    
    For Each currentBeam1 In TargetBeams
        beamAdded = False
        For Each currentBeam2 In TargetBeams
            If Not currentBeam1 Is currentBeam2 Then
                If currentBeam1.DetermineRelationship(currentBeam2) = BEAMRELATIONSHIP.parallel And currentBeam1.SectionProperty.Name = currentBeam2.SectionProperty.Name Then
                    If Not output.exists(currentBeam1.ID) Then
                        Call output.Add(currentBeam1.ID, currentBeam1)
                    End If
                    If Not output.exists(currentBeam2.ID) Then
                        Call output.Add(currentBeam2.ID, currentBeam2)
                    End If
                    beamAdded = True
                    Exit For
                End If
            End If
        Next
        If Not beamAdded And currentBeam1.BeamType = MODELBEAMTYPE.Column Then
            If Not output.exists(currentBeam1.ID) Then
                Call output.Add(currentBeam1.ID, currentBeam1)
            End If
        End If
    Next
    
    Set GetContinuousBeams = New Collection
    For Each currentBeam1 In output.Items
        Call GetContinuousBeams.Add(currentBeam1)
    Next
    
End Function

''''
''' Select beam with the biggest cross-sectional area among the specified collection of beams
'''
Private Function SelectBiggestBeam(ByVal TargetBeams As Collection) As Beam

    Dim maxArea As Double
    Dim biggestBeam As Beam
    Dim currentBeam As Variant
    
    maxArea = 0
    For Each currentBeam In TargetBeams
        If currentBeam.SectionProperty.Ax > maxArea Then
            Set biggestBeam = currentBeam
        End If
    Next
    
    Set SelectBiggestBeam = biggestBeam
    
End Function

'''
''' Convert the specified connection to an (arbitrary) string representation
'''
Private Function ConvertConnectionToString(ByVal TargetConnection As Connection) As String

    Dim sb As New StringBuilder
    Dim currentBeam As Variant
    
    '' Add the node
    Call sb.Append("CN" & TargetConnection.ConnectionNode.ID & ",")
    
    '' Add the main beams
    For Each currentBeam In TargetConnection.MainBeams
        Call sb.Append("MB" & currentBeam.ID & ",")
    Next
    
    '' Add the connecting beams
    For Each currentBeam In TargetConnection.ConnectingBeams
        Call sb.Append("CB" & currentBeam.ID & ",")
    Next
    
    '' Remove the last comma and return
    ConvertConnectionToString = Mid(sb.Text, 1, sb.Length - 1)

End Function

'''
''' Convert the specified string in to a connection. If the connection doesn't exist, then return nothing
'''
Private Function ParseSavedConnectionString(ByVal SavedConnectionString As String) As Connection
    
    Dim Prefix As String
    Dim Suffix As Long
    Dim i As Integer
    Dim entries() As String
    Dim newConnectionNode As Node
    Dim currentBeam As Beam
    Dim newConnectionMainBeams As Collection
    Dim newConnectionConnectingBeams As Collection
    Dim newConnection As Connection
    
    Set newConnection = Nothing
    Set newConnectionMainBeams = New Collection
    Set newConnectionConnectingBeams = New Collection
    
    '' Check the string entries
    entries = Strings.Split(SavedConnectionString, ",", , vbBinaryCompare)
    For i = 0 To UBound(entries)
    
        '' Verify that the current is valid
        If Not CheckSavedConnectionStringEntry(entries(i), Prefix, Suffix) Then
            GoTo SkipEntry
        End If
        
        '' Classify the entry
        Select Case Prefix
            '' Connecting node
            Case "CN"
                If Me.Model.Nodes.exists(Suffix) Then
                    Set newConnectionNode = Me.Model.Nodes(Suffix)
                End If
            '' Main beam
            Case "MB"
                If Me.Model.beams.exists(Suffix) Then
                    Call newConnectionMainBeams.Add(Me.Model.beams(Suffix))
                End If
            '' Connecting beam
            Case "CB"
                If Me.Model.beams.exists(Suffix) Then
                    Call newConnectionConnectingBeams.Add(Me.Model.beams(Suffix))
                End If
            Case Else
                GoTo SkipEntry
        End Select
        
SkipEntry:
    Next
    
    '' Instantiate a new connection if possible
    '' Verify that the node is not null
    If Not newConnectionNode Is Nothing And newConnectionMainBeams.count > 0 And newConnectionConnectingBeams.count > 0 Then
        '' Verify that each of the connected beams contains the node
        For i = newConnectionMainBeams.count To 1 Step -1
            Set currentBeam = newConnectionMainBeams(i)
            If Not (currentBeam.Node1 Is newConnectionNode Or currentBeam.Node2 Is newConnectionNode) Then
                Call newConnectionMainBeams.Remove(i)
            End If
        Next
        For i = newConnectionConnectingBeams.count To 1 Step -1
            Set currentBeam = newConnectionConnectingBeams(i)
            If Not (currentBeam.Node1 Is newConnectionNode Or currentBeam.Node2 Is newConnectionNode) Then
                Call newConnectionConnectingBeams.Remove(i)
            End If
        Next
        
        '' If all checks have been passed instantiate a new connection
        If newConnectionMainBeams.count > 0 And newConnectionConnectingBeams.count > 0 Then
            Set newConnection = New Connection
            Call newConnection.Initialise(newConnectionNode, newConnectionMainBeams)
            Set newConnection.ConnectingBeams = newConnectionConnectingBeams
        End If
    End If
    
    Set ParseSavedConnectionString = newConnection
    
End Function

'''
''' Check if a single entry form a saved connection string conform that the format [CN,MB,CB][0-9]*
'''
Private Function CheckSavedConnectionStringEntry(ByVal entry As String, ByRef Prefix As String, ByRef Suffix As Long) As Boolean

    Dim output As Boolean
    
    Prefix = ""
    Suffix = 0
    output = True
    
    '' The entry must be at least three characters long
    If Len(entry) < 3 Then
        output = False
        GoTo ExitPoint
    End If
    
    '' Separate out the prefix and check validity
    Prefix = UCase(Trim(Left(entry, 2)))
    If IsNumeric(Prefix) Then
        output = False
        GoTo ExitPoint
    End If
    
    '' Separate out the suffix and varify that it is numeric
    On Error GoTo SuffixparseError
    Suffix = CLng(Trim(Mid(entry, 3)))
    If Suffix <= 0 Then
        output = False
        GoTo ExitPoint
    End If
    
ExitPoint:
    CheckSavedConnectionStringEntry = output
    Exit Function
    
SuffixparseError:
    output = False
    Resume ExitPoint
    
End Function
