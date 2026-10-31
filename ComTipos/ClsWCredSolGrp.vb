<Serializable()> Public Class ClsWCredSolGrp

    Public Datos As clsDataset
    Public Integrantes As clsDataset
    Public Producto As clsDataset
    Public SubProducto As clsDataset
    Public Condicion As clsDataset
    Public Grupo As clsDataset
    Public Campania As clsDataset
    Public Lineas As clsDataset
    Public Destinos As clsDataset
    Public Analistas As clsDataset
    Public Actividades As clsDataset
    Public Oficinas As clsDataset

    Public Function SeparaActividad(ByVal psActiv As String)
        Dim i As Integer
        Dim sCad As String
        sCad = ""
        For i = Len(psActiv) To 1 Step -1
            If Mid(psActiv, i, 1) = "\" Then
                Exit For
            End If
            sCad = Mid(psActiv, i, 1) & sCad
        Next i

        SeparaActividad = sCad
    End Function

    Public Sub New()
        Dim C(26) As String
        Dim I(12) As String 'ARCV 08-11-2010 (12)
        I(1) = "nCodPers"
        I(2) = "cPersNombre"
        I(3) = "nPrestamo"
        I(4) = "nCodAct"
        I(5) = "cActividad"
        I(6) = "nCoord"
        I(7) = "nCodCred"
        I(8) = "nTasaComp"
        I(9) = "nMontoCuota"
        I(10) = "nLinea"
        I(11) = "nPrestamoAnt"
        I(12) = "cDocumento" 'ARCV 08-11-2010

        C(1) = "nProd"
        C(2) = "nSubProd"
        C(3) = "nPrestamoTot"
        C(4) = "nCuotas"
        C(5) = "nPlazo"
        C(6) = "nDiaVen"
        C(7) = "dFecDes"
        C(8) = "nCond"
        C(9) = "nDiasGra"
        C(10) = "nTipoGra"
        C(11) = "nCodGrp"
        C(12) = "nCiclo"
        C(13) = "nCamp"
        C(14) = "nLinea"
        C(15) = "nTasaComp"
        C(16) = "nCuota"
        C(17) = "nDestino"
        C(18) = "nAnalista"
        C(19) = "nActiv"
        C(20) = "bCoord"
        C(21) = "nCodAge"
        C(22) = "nTipoPeriodo"
        C(23) = "nMoneda"
        C(24) = "cNomGrupo"
        C(25) = "nOficina"
        C(26) = "nCodModalidad"

        Datos = New clsDataset
        Datos.CrearVacio(C)

        Integrantes = New clsDataset
        Integrantes.CrearVacio(I)

        Producto = New clsDataset
        SubProducto = New clsDataset
        Condicion = New clsDataset
        Grupo = New clsDataset
        Campania = New clsDataset
        Lineas = New clsDataset
        Destinos = New clsDataset
        Analistas = New clsDataset
        Actividades = New clsDataset
        Oficinas = New clsDataset
        
    End Sub
End Class
