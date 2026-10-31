<Serializable()> Public Class clsMicroSeguros

    Public Data As ComTipos.clsDataset
    Public Beneficiarios As ComTipos.clsDataset
    Public Sub New()
        Data = New ComTipos.clsDataset
        Beneficiarios = New ComTipos.clsDataset
    End Sub
End Class
