VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SlicerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type vars
    pastaDeTrabalho As Workbook
    nomeDoSlicer    As String
    valorDoFiltro   As String
    slicerCache     As slicerCache
End Type
Private this As vars
Property Let pastaDeTrabalho(ByVal value As Workbook)
    Set this.pastaDeTrabalho = value
End Property
Property Let nomeDoSlicer(ByVal value As Workbook)
    Set this.nomeDoSlicer = value
End Property
Public Function segmentacaoes(ParamArray segmentacoes() As Variant)
    Dim segmentacao As Variant
    Dim novoSlicerCache As slicerCache
    Dim controleParametros As Byte
    For segmentacao = 0 To UBound(segmentacoes) Step 3
        controleParametros = segmentacao + 1
        Set novoSlicerCache = this.pastaDeTrabalho.SlicerCaches(segmentacoes(segmentacao))
        filtraSegmentacao novoSlicerCache, segmentacoes(controleParametros), segmentacoes(controleParametros + 1)
    Next segmentacao
End Function
Private Function filtraSegmentacao(ByVal segmentacao As slicerCache, ByVal nomeDoSlicer As String, ByVal valueFilter As String)
    Dim stringDeFiltro As String
    segmentacao.VisibleSlicerItemsList = Array(nomeDoSlicer & ".&[" & valueFilter & "]")
End Function
