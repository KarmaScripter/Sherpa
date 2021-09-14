VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectFilterSlice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type vars
    segmentacao     As String
    nomeSegmentacao As String
    valueFilter     As String
End Type
Private this As vars
Property Let segmentacao(ByVal value As String)
    this.segmentacao = value
End Property
Property Let nomeSegmentacao(ByVal value As String)
    this.nomeSegmentacao = value
End Property
Property Let valueFilter(ByVal value As String)
    this.valueFilter = value
End Property
Property Get segmentacao() As String
    segmentacao = this.segmentacao
End Property
Property Get nomeSegmentacao() As String
    segmentacao = this.nomeSegmentacao
End Property
Property Get valueFilter() As String
    segmentacao = this.valueFilter
End Property
Public Function filterProperties() As ObjectFilterSlice
    Set filterProperties = Me
End Function



