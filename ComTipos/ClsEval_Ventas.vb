<Serializable()> Public Class ClsEval_Ventas

    Public oCboMeses As ComTipos.clsDataset
    Public oCboDestino As ComTipos.clsDataset
    Public oCboPlazo As ComTipos.clsDataset
    Public oDataCicAnuNeg As ComTipos.clsDataset
    Public oDataVentas As ComTipos.clsDataset
    Public oDeudaFinan As ComTipos.clsDataset
    Public oDeudaFinanCony As ComTipos.clsDataset
    Public oInventario As ComTipos.clsDataset
    Public oCboTipoProd As ComTipos.clsDataset
    Public psProd1 As String
    Public psProd2 As String
    Public psProd3 As String
    Public psProd4 As String
    Public psProd5 As String
    Public psProd6 As String
    Public oDataMatPri As ComTipos.clsDataset
    Public oDataMatPri2 As ComTipos.clsDataset
    Public oDataMatPri3 As ComTipos.clsDataset
    Public oDataMatPri4 As ComTipos.clsDataset
    Public oDataMatPri5 As ComTipos.clsDataset
    Public oDataMatPri6 As ComTipos.clsDataset
    Public oDataInsumo As ComTipos.clsDataset
    Public oDataInsumo2 As ComTipos.clsDataset
    Public oDataInsumo3 As ComTipos.clsDataset
    Public oDataInsumo4 As ComTipos.clsDataset
    Public oDataInsumo5 As ComTipos.clsDataset
    Public oDataInsumo6 As ComTipos.clsDataset
    Public oDataResumen As ComTipos.clsDataset

    Public Sub New()
        oCboMeses = New ComTipos.clsDataset
        oDataCicAnuNeg = New ComTipos.clsDataset
        oDataVentas = New ComTipos.clsDataset
        oDeudaFinan = New ComTipos.clsDataset
        oInventario = New ComTipos.clsDataset
        oCboDestino = New ComTipos.clsDataset
        oDeudaFinanCony = New ComTipos.clsDataset
        oDataMatPri = New ComTipos.clsDataset
        oDataMatPri2 = New ComTipos.clsDataset
        oDataMatPri3 = New ComTipos.clsDataset
        oDataMatPri4 = New ComTipos.clsDataset
        oDataMatPri5 = New ComTipos.clsDataset
        oDataMatPri6 = New ComTipos.clsDataset
        oDataInsumo = New ComTipos.clsDataset
        oDataInsumo2 = New ComTipos.clsDataset
        oDataInsumo3 = New ComTipos.clsDataset
        oDataInsumo4 = New ComTipos.clsDataset
        oDataInsumo5 = New ComTipos.clsDataset
        oDataInsumo6 = New ComTipos.clsDataset
        oCboPlazo = New ComTipos.clsDataset
        oCboTipoProd = New ComTipos.clsDataset
        oDataResumen = New ComTipos.clsDataset
    End Sub
End Class
