Public Class clsZonas

    Public Sub Departamento_Recupera(ByRef oCred As ComTipos.clsDataset, _
                                    ByVal sCodZona As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Departamento_Recupera '" & sCodZona & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub Zonas_Recupera20(ByRef oCred As ComTipos.clsDataset, _
                                ByVal nGrupo As Integer, _
                                ByVal sBuscar As String, _
                                ByVal nTipoBusq As Integer, _
                                ByVal sPadre As String, _
                                    ByVal sCodZona As String)
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Zonas_Recupera20 " & nGrupo & ",'" & sBuscar & "'," & nTipoBusq & ",'" & sPadre & "','" & sCodZona & "'")
        oCred = oData
        oData = Nothing
    End Sub
End Class
