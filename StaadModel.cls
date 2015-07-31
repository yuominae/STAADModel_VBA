VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaadModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Public events
Public Event StatusUpdate(ByVal StatusMessage As String)

'' Private fields
Private pPath As String
Private pOpenStaad As OpenStaad

'' Public fields
Public Nodes As Dictionary
Public Supports As Dictionary
Public beams As Dictionary
Public Materials As Dictionary
Public properties As Dictionary
Public LoadCases As Dictionary
Public LoadCombinations As Dictionary

'' Properties
Public Property Get Name() As String
    Name = Mid(Me.Path, InStrRev(Me.Path, "\", -1, vbBinaryCompare) + 1)
End Property

Public Property Get Path() As String
    Call Me.OpenStaad.GetSTAADFile(pPath, True)
    Path = pPath
End Property

Public Property Get HasResults() As Boolean
    HasResults = Me.output.AreResultsAvailable
End Property

Public Property Get BaseUnitSystem() As STAADBASEUNITSYSTEM
    BaseUnitSystem = Me.OpenStaad.GetBaseUnit
End Property

Public Property Get ForceInputUnit() As STAADFORCEINPUTUNIT
    ForceInputUnit = Me.OpenStaad.GetInputUnitForForce(vbNullString)
End Property

Public Property Get LengthInputUnit() As STAADLENGTHINPUTUNIT
    LengthInputUnit = Me.OpenStaad.GetInputUnitForLength(vbNullString)
End Property

Public Property Get OpenStaad() As OpenStaad
    If OpenStaad Is Nothing Then
        Set OpenStaad = pOpenStaad
    End If
End Property
Public Property Set OpenStaad(ByVal value As OpenStaad)
    Set pOpenStaad = value
End Property

Public Property Get Geometry() As OSGeometryUI
    If Geometry Is Nothing Then
        Set Geometry = Me.OpenStaad.Geometry
    End If
End Property

Public Property Get Load() As OSLoadUI
    If Load Is Nothing Then
        Set Load = Me.OpenStaad.Load
    End If
End Property

Public Property Get output() As OSOutputUI
    If output Is Nothing Then
        Set output = Me.OpenStaad.output
    End If
End Property

Public Property Get Property() As OSPropertyUI
    If Property Is Nothing Then
        Set Property = Me.OpenStaad.Property
    End If
End Property

Public Property Get Support() As OSSupportUI
    If Support Is Nothing Then
        Set Support = Me.OpenStaad.Support
    End If
End Property

'''
''' Instantiate a new STAAD model object from the first open STAAD model found
'''
Public Function Instantiate() As Boolean
    
    Dim message As String
    
    '' Get the running staad instance
    Instantiate = True
    If Not Me.OpenStaad Is Nothing Then
        Set Me.OpenStaad = Nothing
    End If
    On Error GoTo InstantiationError
    Set Me.OpenStaad = GetObject(, "StaadPro.OpenSTAAD")
    On Error GoTo 0
    
    If Instantiate Then
        message = "Acquired model " & Me.Name & "."
    Else
        message = "Could not get hold of STAAD."
    End If
    RaiseEvent StatusUpdate(message)
    
    Exit Function
    
InstantiationError:
    Instantiate = False
    Resume Next
    
End Function

'''
''' Build the model in memory
'''
Public Function Build() As Boolean
    
    Dim message As String
    Dim success As Boolean
    
    success = True
    'On Error GoTo BuildError
    '' Acquire node data
    Call GetNodes
    '' Acquire support data
    Call GetSupports
    '' Acquire beam data
    Call GetBeams
    '' Acquire material data
    Call GetMaterials
    '' Acquire property data
    Call Getproperties
    '' Acquire load case and combination data
    Call GetLoadCases
    Call GetLoadCombinations
    
    ' Classify members
    RaiseEvent StatusUpdate("Classifying members")
    Call ModelAnalyser.AnalyseModel(Me)
    
ErrorRecoveryPoint:
    If success Then
        message = "Model build completed."
    Else
        message = "An error occured during the build."
    End If
    RaiseEvent StatusUpdate(message)
    Build = success
    
    Exit Function
    
BuildError:
    success = False
    Resume ErrorRecoveryPoint
    
End Function

'''
''' Load the forces for the specified beam at both start and end nodes for the specified load cases
'''
Public Function LoadBeamForces(Optional ByVal TargetBeams As Collection = Nothing, Optional ByVal TargetLoadCases As Collection = Nothing) As Boolean
    
    Dim countBeams As Long
    Dim currentBeamForces(1) As BeamForces
'    Dim currentBeam As Beam
'    Dim currentLoadCase As ILoadCase
    Dim currentBeam As Variant
    Dim currentLoadCase As Variant

    '' Verify that the model actually has results
    If Not Me.HasResults Then
        Exit Function
    End If
    
    '' If no beams have been specified then load all beams in the model
    If TargetBeams Is Nothing Then
        For Each currentBeam In Me.beams.Items
            Call TargetBeams.Add(currentBeam)
        Next
    End If
    
    '' If no load cases have been specified then load all available load cases
    If TargetLoadCases Is Nothing Then
        For Each currentLoadCase In Me.LoadCases.Items
            Call TargetLoadCases.Add(currentLoadCase)
        Next
        For Each currentLoadCase In Me.LoadCombinations.Items
            Call TargetLoadCases.Add(currentLoadCase)
        Next
    End If
    
    ''Loop through all the beams and load the forces
    On Error GoTo ExitPoint
    countBeams = 1
    For Each currentBeam In TargetBeams
        If Not Me.LoadBeamForcesAtNode(currentBeam, currentBeam.Node1, TargetLoadCases) Or Not Me.LoadBeamForcesAtNode(currentBeam, currentBeam.Node2, TargetLoadCases) Then
            GoTo ExitPoint
        End If
        '' Notify the user of the progress
        RaiseEvent StatusUpdate("Acquiring forces for beam " & countBeams & " of " & TargetBeams.count)
        countBeams = countBeams + 1
    Next
    
    LoadBeamForces = True
    
ExitPoint:
End Function

'''
''' Load the forces for the specified beam at the specified node
'''
Public Function LoadBeamForcesAtNode(ByVal TargetBeam As Beam, ByVal TargetNode As Node, Optional ByVal TargetLoadCases As Collection = Nothing) As Boolean
    
    Dim currentLoadCase As Variant
    Dim targetBeamForces As BeamForces
    
    '' If no load cases have been specified then load all available load cases
    If TargetLoadCases Is Nothing Then
        For Each currentLoadCase In Me.LoadCases.Items
            Call TargetLoadCases.Add(currentLoadCase)
        Next
        For Each currentLoadCase In Me.LoadCombinations.Items
            Call TargetLoadCases.Add(currentLoadCase)
        Next
    End If
    
    '' Loop through load cases and retrieve forces for the beam at the specified node
    'On Error GoTo ExitPoint
    For Each currentLoadCase In TargetLoadCases
        Set targetBeamForces = InstantiateBeamForces(TargetBeam, TargetNode, currentLoadCase)
        If TargetNode Is TargetBeam.Node1 Then
            Call TargetBeam.Forces1.Add(targetBeamForces.ParentLoadCase.ID, targetBeamForces)
        Else
            Call TargetBeam.Forces2.Add(targetBeamForces.ParentLoadCase.ID, targetBeamForces)
        End If
    Next
    
    LoadBeamForcesAtNode = True
    
ExitPoint:
End Function

Private Sub GetNodes()
    
    Dim i As Long
    Dim count As Long
    Dim nodeIDs() As Long
    
    '' Instantiate the nodes dictionary
    Set Me.Nodes = New Dictionary
    
    '' Get node count and ids
    count = Me.Geometry.GetNodeCount - 1
    If count < 0 Then
        Exit Sub
    End If
    ReDim nodeIDs(count)
    Call Me.Geometry.GetNodeList(nodeIDs)
    
    '' Assign nodes
    For i = 0 To count
        Call Me.Nodes.Add(nodeIDs(i), InstantiateNode(nodeIDs(i)))
        RaiseEvent StatusUpdate("Acquiring node " & i & " of " & count + 1)
    Next

End Sub

Private Sub GetSupports()

    Dim i As Long
    Dim count As Long
    Dim supportNodeIDs() As Long
    Dim currentNode As Node
    
    '' Instantiate supports dictionary
    Set Me.Supports = New Dictionary
    
    '' Get support count and ids
    count = Me.Support.GetSupportCount - 1
    If count < 0 Then
        Exit Sub
    End If
    ReDim supportNodeIDs(count)
    Call Me.Support.GetSupportNodes(supportNodeIDs)
        
    '' Assign supports
    For i = 0 To count
        Set currentNode = Me.Nodes(supportNodeIDs(i))
        currentNode.NodeSupport = InstantiateSupport(currentNode.ID)
        RaiseEvent StatusUpdate("Acquiring support " & i & " of " & count + 1)
    Next

End Sub

Private Sub GetBeams()

    Dim i As Long
    Dim count As Long
    Dim beamIDs() As Long
    
    '' Instantiate the beams dictionary
    Set Me.beams = New Dictionary
    
    '' Get beam count and ids
    count = Me.Geometry.GetMemberCount - 1
    If count < 0 Then
        Exit Sub
    End If
    ReDim beamIDs(count)
    Call Me.Geometry.GetBeamList(beamIDs)
    
    '' Assign beams
    For i = 0 To count
        Call Me.beams.Add(beamIDs(i), InstantiateBeam(beamIDs(i)))
        RaiseEvent StatusUpdate("Acquiring beam " & i & " of " & count + 1)
    Next
    
End Sub

Private Sub GetMaterials()

    Dim materialName As String
    Dim i As Long
    Dim iMax As Long
    Dim b() As Variant
    Dim newMaterial As Material
    
    '' Instantiate materials dictionary
    Set Me.Materials = New Dictionary
    
    RaiseEvent StatusUpdate("Acquiring materials...")
    
    '' Build materials
    b = Me.beams.Items
    iMax = UBound(b)
    For i = 0 To iMax
        If b(i).SectionMaterial Is Nothing Then
            materialName = Me.Property.GetBeamMaterialName(b(i).ID)
            Call InstantiateMaterial(materialName)
        End If
    Next

End Sub

Private Sub Getproperties()

    Dim i As Long
    Dim count As Long
    Dim propertyIDs() As Long
    
    '' Instantiate the properties dictionary
    Set Me.properties = New Dictionary
    
    '' Get property count and ids
    count = Me.Property.GetSectionPropertyCount - 1
    If count < 0 Then
        Exit Sub
    End If
    ReDim propertyIDs(count)
    Call Me.Property.GetSectionPropertyList(propertyIDs)
    
    '' Intantiate the properties and assign to beams
    For i = 0 To count
        Call Me.properties.Add(propertyIDs(i), InstantiateProperty(propertyIDs(i)))
        RaiseEvent StatusUpdate("Acquiring section property" & i & " of " & count + 1)
    Next

End Sub

Private Sub GetLoadCases()

    Dim i As Long
    Dim count As Long
    Dim loadCaseIDs() As Long
    
    '' Instantiate load cases dictionary
    Set Me.LoadCases = New Dictionary
    
    '' Get load case count and ids
    count = Me.Load.GetPrimaryLoadCaseCount - 1
    If count < 0 Then
        Exit Sub
    End If
    ReDim loadCaseIDs(count)
    Call Me.Load.GetPrimaryLoadCaseNumbers(loadCaseIDs)
    
    '' Get load cases
    For i = 0 To count
        Call Me.LoadCases.Add(loadCaseIDs(i), InstantiateLoadCase(loadCaseIDs(i)))
        RaiseEvent StatusUpdate("Acquiring load case " & i & " of " & count + 1)
    Next
    
End Sub

'''
''' Retrieve and instantiate all load combinations in the model
'''
Private Sub GetLoadCombinations()
    
    Dim i As Long
    Dim count As Long
    Dim loadCombinationIDs() As Long
    
    '' Instantiate load combinations dictionary
    Set Me.LoadCombinations = New Dictionary
    
    '' Get load combination count and ids
    count = Me.Load.GetLoadCombinationCaseCount - 1
    If count < 0 Then
        Exit Sub
    End If
    ReDim loadCombinationIDs(count)
    Call Me.Load.GetLoadCombinationCaseNumbers(loadCombinationIDs)
    
    '' Build load combinations
    For i = 0 To count
        Call Me.LoadCombinations.Add(loadCombinationIDs(i), InstantiateLoadCombination(loadCombinationIDs(i)))
        RaiseEvent StatusUpdate("Acquiring load combination " & i & " of " & count)
    Next
    
End Sub

'''
''' Instantiate a new node for the specified node ID
'''
Private Function InstantiateNode(ByVal NodeID As Long) As Node
    
    Dim nodeX As Double
    Dim nodeY As Double
    Dim nodeZ As Double
    Dim newNode As Node
    
    '' Get the coordinates of the node
    Call Me.Geometry.GetNodeCoordinates(NodeID, nodeX, nodeY, nodeZ)
    
    '' Instantiate the new node and return
    Set newNode = New Node
    Call newNode.Initialise(NodeID, nodeX, nodeY, nodeZ)
    
    Set InstantiateNode = newNode
    
End Function

'''
''' Instantiate a new support for the specified node ID
'''
Private Function InstantiateSupport(ByVal NodeID As Long) As Support
    
    Dim supportID As Long
    Dim forceReleases(5) As Integer
    Dim springReleases(5) As Double
    Dim exists As Boolean
    Dim newRelease As Release
    Dim newSupport As Support
    Dim SupportType As STAADSUPPORTTYPE
    
    '' Get support information
    Call Me.Support.GetSupportInformationEx(NodeID, supportID, SupportType, forceReleases, springReleases)
    
    '' Determing whether this type of support already exists
    If Me.Supports.exists(supportID) Then
        Set newSupport = Me.Supports(supportID)
        exists = True
    Else
        Set newSupport = New Support
        Call newSupport.Initialise(supportID)
        Call Me.Supports.Add(newSupport.ID, newSupport)
    End If
    
    If Not exists Then
        newSupport.SupportType = SupportType
        Set newRelease = New Release
        With newRelease
            .Fx = forceReleases(0) > 0
            .Fy = forceReleases(1) > 0
            .Fz = forceReleases(2) > 0
            .Mx = forceReleases(3) > 0
            .My = forceReleases(4) > 0
            .Mz = forceReleases(5) > 0
        End With
        Set newSupport.SupportRelease = newRelease
    End If
    
    Set InstantiateSupport = newSupport
    
End Function

'''
''' Instantiate a new beam from the specied ID
'''
Private Function InstantiateBeam(ByVal BeamID As Long) As Beam
    
    Dim specCode As Integer
    Dim node1ID As Long
    Dim node2ID As Long
    Dim newBeam As Beam
    
    '' Get nodes
    Call Me.Geometry.GetMemberIncidence(BeamID, node1ID, node2ID)
    Set newBeam = New Beam
    Call newBeam.Initialise(BeamID, Me.Nodes(node1ID), Me.Nodes(node2ID))
    
    '' Get beam releases
    Set newBeam.Release1 = GetBeamReleaseSpecs(newBeam.ID, 0)
    Set newBeam.Release2 = GetBeamReleaseSpecs(newBeam.ID, 1)
    
    '' Get beam type
    specCode = 0
    Call Me.Property.GetMemberSpecCode(newBeam.ID, specCode)
    Select Case specCode
        Case 0
            newBeam.Spec = STAADBEAMSPEC.MEMBERTRUSS
        Case 1
            newBeam.Spec = STAADBEAMSPEC.TENSIONMEMBER
        Case 2
            newBeam.Spec = STAADBEAMSPEC.COMPRESSIONMEMBER
        Case 3
            newBeam.Spec = STAADBEAMSPEC.CABLE
        Case 4
            newBeam.Spec = STAADBEAMSPEC.JOIST
        Case Else
            newBeam.Spec = STAADBEAMSPEC.UNSPECIFIED
    End Select

    '' Get beam beta angle
    newBeam.BetaAngle = Me.Property.GetBetaAngle(newBeam.ID)

    Set InstantiateBeam = newBeam

End Function

'''
''' Instantiate a new set of beam releases for the specified beam at the specified end (0 for start, 1 to end)
'''
Private Function GetBeamReleaseSpecs(ByVal BeamID As Long, ByVal BeamEnd As Integer) As Release

    Dim forceReleases(5) As Integer
    Dim springReleases(5) As Double
    
    Call Me.Property.GetMemberReleaseSpec(BeamID, BeamEnd, forceReleases, springReleases)
    
    Set GetBeamReleaseSpecs = New Release
    With GetBeamReleaseSpecs
        .Fx = forceReleases(0) > 0
        .Fy = forceReleases(1) > 0
        .Fz = forceReleases(2) > 0
        .Mx = forceReleases(3) > 0
        .My = forceReleases(4) > 0
        .Mz = forceReleases(5) > 0
    End With

End Function

'''
''' Instantiate a new material with the speciefied ID
'''
Private Sub InstantiateMaterial(ByVal materialName As String)
    
    Dim i As Long
    Dim count As Long
    Dim beamIDs() As Long
    Dim exists As Boolean
    Dim newMaterial As Material
    
    '' Check if the material already exists, if not instantiate a new one
    If Me.Materials.exists(materialName) Then
        Set newMaterial = Me.Materials(materialName)
        exists = True
    Else
        Set newMaterial = New Material
        newMaterial.Name = materialName
        Call Me.Materials.Add(newMaterial.Name, newMaterial)
    End If
    
    '' If the material doesn't yet exist, then assign it to the beams which are made of it
    If Not exists Then
        count = Me.Property.GetIsotropicMaterialAssignedBeamCount(newMaterial.Name) - 1
        ReDim beamIDs(count)
        Call Me.Property.GetIsotropicMaterialAssignedBeamList(newMaterial.Name, beamIDs)
        For i = 0 To count
            Me.beams(beamIDs(i)).SectionMaterial = newMaterial
        Next
    End If
    
End Sub

'''
''' Instantiate a new section property with the specified ID
'''
Private Function InstantiateProperty(ByVal PropertyID As Long) As Property

    Dim propertyName As String
    Dim i As Long
    Dim count As Long
    Dim beamIDs() As Long
    Dim sectionWidth As Double
    Dim sectionDepth As Double
    Dim sectionWebThickness As Double
    Dim sectionFlangeThickness As Double
    Dim sectionAx As Double
    Dim sectionAy As Double
    Dim sectionAz As Double
    Dim sectionIx As Double
    Dim sectionIy As Double
    Dim sectionIz As Double
    Dim newProperty As Property
    
    '' Get the section property name and properties from the API
    Call Me.Property.GetSectionPropertyName(PropertyID, propertyName)
    Call Me.Property.GetSectionPropertyValues(PropertyID, sectionWidth, sectionDepth, sectionAx, sectionAy, sectionAz, sectionIx, sectionIy, sectionIz, sectionFlangeThickness, sectionWebThickness)
    Set newProperty = New Property
    With newProperty
        .ID = PropertyID
        .Name = propertyName
        .Width = sectionWidth
        .Depth = sectionDepth
        .FlangeThickness = sectionFlangeThickness
        .WebThickness = sectionWebThickness
        .Ax = sectionAx
        .Ay = sectionAy
        .Az = sectionAz
        .Ix = sectionIx
        .Iy = sectionIy
        .Iz = sectionIz
    End With
    
    '' Get assigned beams and assign property to beams and beams to property
    count = Me.Property.GetSectionPropertyAssignedBeamCount(PropertyID) - 1
    ReDim beamIDs(count)
    Call Me.Property.GetSectionPropertyAssignedBeamList(PropertyID, beamIDs)
    For i = 0 To count
        Me.beams(beamIDs(i)).SectionProperty = newProperty
    Next
    
    Set InstantiateProperty = newProperty
    
End Function

'''
''' Instantiate a new load case from the specified load case ID
'''
Private Function InstantiateLoadCase(ByVal LoadCaseID As Long) As LoadCase

    Dim newLoadCaseName As String
    Dim newLoadCase As LoadCase
    
    '' Create new load case
    newLoadCaseName = Me.Load.GetLoadCaseTitle(LoadCaseID)
    Set newLoadCase = New LoadCase
    Call newLoadCase.Initialise(LoadCaseID, newLoadCaseName)
    newLoadCase.LoadCaseType = Me.Load.GetLoadType(newLoadCase.ILoadCase_ID)
    
    Set InstantiateLoadCase = newLoadCase
    
End Function

'''
''' Instantiate a new load combination from the specified load combination ID
'''
Private Function InstantiateLoadCombination(ByVal LoadCombinationID As Long) As LoadCombination
    
    Dim newLoadCombinationName As String
    Dim i As Long
    Dim loadCaseCount As Long
    Dim loadCaseIDs() As Long
    Dim loadCaseFactors() As Long
    Dim currentLoadCase As LoadCase
    Dim newLoadCombination As LoadCombination
    
    '' Create new load combination
    newLoadCombinationName = Me.Load.GetLoadCaseTitle(LoadCombinationID)
    Set newLoadCombination = New LoadCombination
    Call newLoadCombination.Initialise(LoadCombinationID, newLoadCombinationName)
    
    '' Get the load cases and factors forming the combination
    loadCaseCount = Me.Load.GetNoOfLoadAndFactorPairsForCombination(newLoadCombination.ILoadCase_ID) - 1
    ReDim loadCaseIDs(loadCaseCount)
    ReDim loadCaseFactors(loadCaseCount)
    Call Me.Load.GetLoadAndFactorForCombination(newLoadCombination.ILoadCase_ID, loadCaseIDs, loadCaseFactors)
    '' Assign the load cases to the combination and the combination to the load cases
    For i = 0 To loadCaseCount
        Set currentLoadCase = Me.LoadCases(loadCaseIDs(i))
        Call newLoadCombination.LoadCasesAndFactors.Add(currentLoadCase, loadCaseFactors(i))
        Call currentLoadCase.Combinations.Add(newLoadCombination.ILoadCase_ID, newLoadCombination)
    Next
    
    Set InstantiateLoadCombination = newLoadCombination
    
End Function

'''
''' Rerieve forces for the specified Beam at the specified node for the specified load case
'''
Private Function InstantiateBeamForces(ByVal TargetBeam As Beam, ByVal TargetNode As Node, ByVal TargetLoadCase As ILoadCase) As BeamForces
    
    Dim forces(5) As Double
    Dim currentBeamForces As BeamForces
    
    '' Get the forces from the API and assign to a new BeamForces object
    Call Me.output.GetMemberEndForces(TargetBeam.ID, IIf(TargetNode Is TargetBeam.Node1, 0, 1), TargetLoadCase.ID, forces, 0)
    Set currentBeamForces = New BeamForces
    With currentBeamForces
        .Fx = forces(0)
        .Fy = forces(1)
        .Fz = forces(2)
        .Mx = forces(3)
        .My = forces(4)
        .Mz = forces(5)
        Set .ParentLoadCase = TargetLoadCase
    End With
    
    Set InstantiateBeamForces = currentBeamForces
    
End Function
