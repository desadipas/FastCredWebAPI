<Serializable()> Public Class ClsWRegistroCred

    Public DatosRelaPers As ComTipos.clsDataset
    Public cboProd As ComTipos.clsDataset
    Public cboSubProd As ComTipos.clsDataset
    Public cboGrupo As ComTipos.clsDataset
    Public cboOficinas As ComTipos.clsDataset
    Public cboLinea As ComTipos.clsDataset
    Public cboDestino As ComTipos.clsDataset
    Public cboAnalista As ComTipos.clsDataset
    Public cboCondicion As ComTipos.clsDataset
    Public Datos As ComTipos.clsDataset
    Public CboCampanias As ComTipos.clsDataset
    Public CboAct As ComTipos.clsDataset
    Public CalDesParcial As ComTipos.clsDataset
    Public cboPromotores As ComTipos.clsDataset
    Public cboRecaudadores As ComTipos.clsDataset
    '--vvillafranca 06-08-2012
    Public oComentarios As ComTipos.clsDataset
    '--
    ''--------vvillafranca 05-11-2012------
    'Public oMesesDoble As ComTipos.clsDataset
    ''------------------------


    Public Sub New()
        DatosRelaPers = New ComTipos.clsDataset
        cboProd = New ComTipos.clsDataset
        cboSubProd = New ComTipos.clsDataset
        cboGrupo = New ComTipos.clsDataset
        cboLinea = New ComTipos.clsDataset
        cboDestino = New ComTipos.clsDataset
        cboAnalista = New ComTipos.clsDataset
        Datos = New ComTipos.clsDataset
        cboCondicion = New ComTipos.clsDataset
        CboCampanias = New ComTipos.clsDataset
        CboAct = New ComTipos.clsDataset
        cboOficinas = New ComTipos.clsDataset
        CalDesParcial = New ComTipos.clsDataset
        cboPromotores = New ComTipos.clsDataset
        cboRecaudadores = New ComTipos.clsDataset
        '--vvillafranca 06-08-2012
        oComentarios = New ComTipos.clsDataset
        ''--
        ''--------vvillafranca 05-11-2012------
        'oMesesDoble = New ComTipos.clsDataset
        ''-------------------------------------

    End Sub
End Class
