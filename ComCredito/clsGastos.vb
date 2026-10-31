Imports System.Data
Imports System.Data.SqlClient

Public Class clsGastos

    Public Function CredGastos_Listar() As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CredGastos_Listar ")
        Return oData
    End Function

    Public Sub CredGastos_Insertar(ByVal sDescrip As String, _
                                ByVal nTipoGasto As Integer, _
                                ByVal nValor As Double, _
                                ByVal nTipoValor As Integer, _
                                ByVal nAplicado As Integer, _
                                ByVal nRangoIni As Double, _
                                ByVal nRangoFin As Double, _
                                ByVal nMoneda As Integer, _
                                ByVal nProd As Integer, _
                                ByVal nSubProd As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredGastos_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cDescrip", SqlDbType.VarChar, 150).Value = sDescrip
        oComm.Parameters.Add("@nTipoGasto", SqlDbType.Int).Value = nTipoGasto
        oComm.Parameters.Add("@nValor", SqlDbType.Money).Value = nValor
        oComm.Parameters.Add("@nTipoValor", SqlDbType.Int).Value = nTipoValor
        oComm.Parameters.Add("@nAplicado", SqlDbType.Int).Value = nAplicado
        oComm.Parameters.Add("@nRangoIni", SqlDbType.Money).Value = nRangoIni
        oComm.Parameters.Add("@nRangoFin", SqlDbType.Money).Value = nRangoFin
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = nProd
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = nSubProd

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredGastos_Modificar(ByVal nCodGasto As Integer, _
                                ByVal sDescrip As String, _
                                 ByVal nTipoGasto As Integer, _
                                 ByVal nValor As Double, _
                                 ByVal nTipoValor As Integer, _
                                 ByVal nAplicado As Integer, _
                                 ByVal nRangoIni As Double, _
                                 ByVal nRangoFin As Double, _
                                 ByVal nMoneda As Integer, _
                                 ByVal nProd As Integer, _
                                 ByVal nSubProd As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredGastos_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGasto", SqlDbType.Int).Value = nCodGasto
        oComm.Parameters.Add("@cDescrip", SqlDbType.VarChar, 150).Value = sDescrip
        oComm.Parameters.Add("@nTipoGasto", SqlDbType.Int).Value = nTipoGasto
        oComm.Parameters.Add("@nValor", SqlDbType.Money).Value = nValor
        oComm.Parameters.Add("@nTipoValor", SqlDbType.Int).Value = nTipoValor
        oComm.Parameters.Add("@nAplicado", SqlDbType.Int).Value = nAplicado
        oComm.Parameters.Add("@nRangoIni", SqlDbType.Money).Value = nRangoIni
        oComm.Parameters.Add("@nRangoFin", SqlDbType.Money).Value = nRangoFin
        oComm.Parameters.Add("@nMoneda", SqlDbType.Int).Value = nMoneda
        oComm.Parameters.Add("@nProd", SqlDbType.Int).Value = nProd
        oComm.Parameters.Add("@nSubProd", SqlDbType.Int).Value = nSubProd

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredGastos_Eliminar(ByVal nCodGasto As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredGastos_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodGasto", SqlDbType.Int).Value = nCodGasto
        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing

    End Sub

    Public Sub CredGastos_CargaInterfazMan(ByRef oMoneda As ComTipos.clsDataset, _
                                        ByRef oTipo As ComTipos.clsDataset, _
                                        ByRef oTValor As ComTipos.clsDataset, _
                                        ByRef oProducto As ComTipos.clsDataset, _
                                        ByRef oSubProducto As ComTipos.clsDataset)

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        oMoneda = C.RecuperaConstantes(2090)
        oProducto = C.RecuperaConstantes(1500)
        oSubProducto = C.RecuperaConstantes(1600)
        oTipo = C.RecuperaConstantes(2164)
        oTValor = C.RecuperaConstantes(2165)
        C = Nothing
    End Sub

    Public Function CredGastos_CargaInterfazAsig() As ComTipos.clsDataset

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        Return C.RecuperaConstantes(2164)
        C = Nothing
    End Function

    Public Function CredGastos_Buscar_x_Tipo(ByVal nTipoGasto As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CredGastos_Buscar_x_Tipo " & nTipoGasto)
        Return oData
    End Function

    Public Function CargarAsignacionGastos(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CargarAsignacionGastos " & nCodCred & "," & nCodAge)
        Return oData
    End Function


    Public Function CargarGastosCuotas(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer, _
                                ByVal nNroCuota As Integer, _
                                ByVal nTipoGasto As Integer) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_CargarGastosCuotas " & nCodCred & "," & nCodAge & "," & nNroCuota & "," & nTipoGasto)
        Return oData
    End Function

    Public Sub CredCalenGastos_Eliminar(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer, _
                                ByVal nNroCuota As Integer, _
                                ByVal nCodGasto As Integer, _
                                ByVal nNroCalen As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredCalenGastos_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nNroCuota", SqlDbType.Int).Value = nNroCuota
        oComm.Parameters.Add("@nCodGasto", SqlDbType.Int).Value = nCodGasto
        oComm.Parameters.Add("@nNroCalen", SqlDbType.Int).Value = nNroCalen

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredCalenGastos_Modificar(ByVal nCodCred As Integer, _
                            ByVal nCodAge As Integer, _
                            ByVal nNroCuota As Integer, _
                            ByVal nCodGasto As Integer, _
                            ByVal nMonto As Double, _
                            ByVal dFecAsig As Date, _
                            ByVal nNroCalen As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredCalenGastos_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@nNroCuota", SqlDbType.Int).Value = nNroCuota
        oComm.Parameters.Add("@nCodGasto", SqlDbType.Int).Value = nCodGasto
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@dFecAsig", SqlDbType.DateTime).Value = dFecAsig
        oComm.Parameters.Add("@nNroCalen", SqlDbType.Int).Value = nNroCalen

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredCalenGastos_Insertar(ByVal nCodCred As Integer, _
                                ByVal nCodAge As Integer, _
                                ByVal oNroCuota As ComTipos.clsDataset, _
                                ByVal nCodGasto As Integer, _
                                ByVal nMonto As Double, _
                                ByVal dFecAsig As Date, _
                                ByVal nTipo As Integer, _
                                ByVal nNroCalen As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredCalenGastos_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@oNroCuota", SqlDbType.Xml).Value = oNroCuota.DevuelveXML
        oComm.Parameters.Add("@nCodGasto", SqlDbType.Int).Value = nCodGasto
        oComm.Parameters.Add("@nMonto", SqlDbType.Money).Value = nMonto
        oComm.Parameters.Add("@dFecAsig", SqlDbType.DateTime).Value = dFecAsig
        oComm.Parameters.Add("@nTipo", SqlDbType.Int).Value = nTipo
        oComm.Parameters.Add("@nNroCalen", SqlDbType.Int).Value = nNroCalen

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub
End Class
