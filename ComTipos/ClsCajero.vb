Option Explicit On
<Serializable()> Public Class clsCajero
    Public cboBanco As ComTipos.clsDataset
    Public cboModoTransporte As ComTipos.clsDataset
    Public Sub New()
        cboBanco = New ComTipos.clsDataset
        cboModoTransporte = New ComTipos.clsDataset
    End Sub
End Class
