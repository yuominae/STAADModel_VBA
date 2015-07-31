VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Public fields
Public ConnectionNode As Node
Public MainBeams As Collection
Public ConnectingBeams As Collection

Private Sub Class_Initialize()
    
    Set Me.MainBeams = New Collection
    Set Me.ConnectingBeams = New Collection
    
End Sub

Public Sub Initialise(ByVal ConnectionNode As Node, ByVal MainBeams As Collection)

    Set Me.ConnectionNode = ConnectionNode
    Set Me.MainBeams = MainBeams

End Sub
