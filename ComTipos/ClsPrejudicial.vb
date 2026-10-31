<Serializable()> Public Class ClsPrejudicial

    Public GestorCobranza As ComTipos.clsDataset
    Public CalendariActual As ComTipos.clsDataset
    Public CalendariNegociaciones As ComTipos.clsDataset
    Public ClientesxGestorCobranza As ComTipos.clsDataset


    Public Sub New()
        GestorCobranza = New ComTipos.clsDataset
        CalendariActual = New ComTipos.clsDataset
        CalendariNegociaciones = New ComTipos.clsDataset
        ClientesxGestorCobranza = New ComTipos.clsDataset
    End Sub

End Class
