Option Explicit On

<Serializable()> Public Class ClsCredDupCal
    Public oDato As ComTipos.clsDataset
    Public oCalend As ComTipos.clsDataset
    Public oCalenDesemb As ComTipos.clsDataset

    Public Sub New()
        oDato = New ComTipos.clsDataset
        oCalend = New ComTipos.clsDataset
        oCalenDesemb = New ComTipos.clsDataset
    End Sub

End Class
