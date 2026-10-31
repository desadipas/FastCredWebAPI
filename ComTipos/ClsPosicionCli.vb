
<Serializable()> Public Class ClsPosicionCli
    Public RCC As ComTipos.clsDataset
    Public RCCDet As ComTipos.clsDataset
    Public Datos As ComTipos.clsDataset
    Public RCCTemp As ComTipos.clsDataset
    Public Comen As ComTipos.clsDataset
    Public BaseNegativa As ComTipos.clsDataset

    Public TamizGenesys As ComTipos.clsDataset 'ARCV 21-11-2010
    Public SCORD As ComTipos.clsDataset 'ARCV 28-02-2012

    Public Sub New()
        RCC = New ComTipos.clsDataset
        RCCDet = New ComTipos.clsDataset
        Datos = New ComTipos.clsDataset
        RCCTemp = New ComTipos.clsDataset
        Comen = New ComTipos.clsDataset
        BaseNegativa = New ComTipos.clsDataset
        TamizGenesys = New ComTipos.clsDataset
        SCORD = New ComTipos.clsDataset
    End Sub
End Class
