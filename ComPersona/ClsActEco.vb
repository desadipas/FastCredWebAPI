Public Class ClsActEco
    Public Sub DevuelveTipoActividad(ByRef oTipo As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveTipoActividad ")
        oTipo = oData
        oData = Nothing
    End Sub

    Public Sub DevuelveSubTipoActividad(ByRef oTipo As ComTipos.clsDataset, ByVal pnTipo As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveSubTipoActividad '" & pnTipo & "'")
        oTipo = oData
        oData = Nothing
    End Sub

    Public Sub DevuelveActividad(ByRef oTipo As ComTipos.clsDataset, ByVal pnSubTipo As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveActividad '" & pnSubTipo & "'")
        oTipo = oData
        oData = Nothing
    End Sub

    Public Sub DevuelveTipoCIIU(ByRef oTipo As ComTipos.clsDataset)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveTipoCIIU ")
        oTipo = oData
        oData = Nothing
    End Sub

    Public Sub DevuelveSubTipoCIIU(ByRef oTipo As ComTipos.clsDataset, ByVal pnTipo As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveSubTipoCIIU '" & pnTipo & "'")
        oTipo = oData
        oData = Nothing
    End Sub

    Public Sub DevuelveCIIU(ByRef oTipo As ComTipos.clsDataset, ByVal pnSubTipo As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_DevuelveCIIU '" & pnSubTipo & "'")
        oTipo = oData
        oData = Nothing
    End Sub
End Class
