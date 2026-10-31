<Serializable()> Public Class ClsWAprobCred

    Public DatosRelaPers As ComTipos.clsDataset
    Public cboProd As ComTipos.clsDataset
    Public cboSubProd As ComTipos.clsDataset
    Public cboLinea As ComTipos.clsDataset
    Public cboDestino As ComTipos.clsDataset
    Public cboCondicion As ComTipos.clsDataset
    Public cboAct As ComTipos.clsDataset
    Public Datos As ComTipos.clsDataset
    Public CalDesParcial As ComTipos.clsDataset
    '--------vvillafranca 05-11-2012------
    Public oMesesDoble As ComTipos.clsDataset
    '------------------------

    Public Sub New()
        DatosRelaPers = New ComTipos.clsDataset
        cboProd = New ComTipos.clsDataset
        cboSubProd = New ComTipos.clsDataset
        cboLinea = New ComTipos.clsDataset
        cboDestino = New ComTipos.clsDataset
        cboAct = New ComTipos.clsDataset
        Datos = New ComTipos.clsDataset
        cboCondicion = New ComTipos.clsDataset
        CalDesParcial = New ComTipos.clsDataset
        '--------vvillafranca 05-11-2012------
        oMesesDoble = New ComTipos.clsDataset
        '-------------------------------------
    End Sub

End Class
