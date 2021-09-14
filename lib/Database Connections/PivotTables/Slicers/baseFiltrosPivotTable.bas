Attribute VB_Name = "baseFiltrosPivotTable"
Option Explicit


Sub filtrarSegmentacao(ByVal funcionario As String, ByVal status As String)
    Dim meuSlicer As New SlicerControl
    Dim teste As Variant
    
    meuSlicer.pastaDeTrabalho = ActiveWorkbook
    meuSlicer.segmentacaoes "SegmentaçãodeDados_baseRepresentantesAtendimento.Funcionario", _
                            "[relatorioCompleto].[baseRepresentantesAtendimento.Funcionario]", _
                            "BIANKA", _
                            "SegmentaçãodeDados_BaseDeStatus.STATUS_FINAL", _
                            "[relatorioCompleto].[BaseDeStatus.STATUS FINAL]", _
                            "RESERVADO"
End Sub


