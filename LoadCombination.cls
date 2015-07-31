VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadCombination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ILoadCase

'' Private fields
Private pName As String
Private pID As Long

'' Public fields
Public LoadCasesAndFactors As Dictionary

'' Public properties
Public Property Get ILoadCase_Name() As String
    ILoadCase_Name = pName
End Property
Private Property Let ILoadCase_Name(ByVal value As String)
    pName = value
End Property

Public Property Get ILoadCase_ID() As Long
    ILoadCase_ID = pID
End Property
Public Property Let ILoadCase_ID(ByVal value As Long)
    pID = value
End Property

'' public methods
Private Sub Class_Initialize()
    
    Set Me.LoadCasesAndFactors = New Dictionary
    
End Sub

Public Sub Initialise(ByVal ID As Long, ByVal Name As String)

    ILoadCase_ID = ID
    ILoadCase_Name = Name

End Sub
