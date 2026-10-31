Imports System.Data
Imports System.Data.SqlClient

Public Class clsRCD

    Public Sub RCDGeneraDataRCD(ByVal sFecha As String, _
                            ByVal nTipoCambio As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraDataRCD", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha
        oComm.Parameters.Add("@nTipoCambio", SqlDbType.Money).Value = nTipoCambio

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RCDGeneraDataRCA(ByVal sFecha As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraDataRCA", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RCDGeneraDataRCM(ByVal sFecha As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraDataRCM", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RCDGeneraNumSecuencia(ByVal sFecha As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraNumSecuencia", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RecuperaRCDParametro(ByRef oCred As ComTipos.clsDataset, _
        ByVal dFecha As Date)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RecuperaRCDParametro '" & dFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCDReporte_Totales(ByRef oCred As ComTipos.clsDataset, _
    ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCDReporte_Totales '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCDReporte_Detalle(ByRef oCred As ComTipos.clsDataset, _
    ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCDReporte_Detalle '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCMReporte_Detalle(ByRef oCred As ComTipos.clsDataset, _
                    ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCMReporte_Detalle '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCMReporte_Mancomunos(ByRef oCred As ComTipos.clsDataset, _
                    ByVal sFecha As String, _
                    ByVal sNumSec As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCMReporte_Mancomunos '" & sNumSec & "','" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCAReporte_Detalle(ByRef oCred As ComTipos.clsDataset, _
                ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCAReporte_Detalle '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCDGeneraDataRCD_11356(ByVal sFecha As String, _
                        ByVal nTipoCambio As Double)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraDataRCD_11356", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha
        oComm.Parameters.Add("@nTipoCambio", SqlDbType.Money).Value = nTipoCambio

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RCDGeneraDataRCA_11356(ByVal sFecha As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraDataRCA_11356", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RCDGeneraDataRCM_11356(ByVal sFecha As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_RCDGeneraDataRCM_11356", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.Parameters.Add("@cFecha", SqlDbType.Char, 6).Value = sFecha

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub RCMReporte_Detalle_11356(ByRef oCred As ComTipos.clsDataset, _
                ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCMReporte_Detalle_11356 '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCAReporte_Detalle_11356(ByRef oCred As ComTipos.clsDataset, _
            ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCAReporte_Detalle_11356 '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCDReporte_Totales_11356(ByRef oCred As ComTipos.clsDataset, _
ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCDReporte_Totales '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RCDReporte_Detalle_11356(ByRef oCred As ComTipos.clsDataset, _
    ByVal sFecha As String)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RCDReporte_Detalle_11356 '" & sFecha & "'")
        oCred = oData
        oData = Nothing
    End Sub

    Public Sub RO_DevuelveListaOperaciones(ByRef oCred As ComTipos.clsDataset, _
            ByVal dFecha As Date, ByVal bSoloEfectivo As Integer)

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_RO_DevuelveListaOperaciones '" & dFecha & "'," & bSoloEfectivo, , , 0)
        oCred = oData
        oData = Nothing
    End Sub


    '******************************* CRO Central de Riespo Operaciones  *********************************************************************************************
    Public Sub CRO_CargarDatos_P1()

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CRO_CargarDatos_P1", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CRO_CargarDatos_G1()

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CRO_CargarDatos_G1", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CRO_CargarDatos_G4()

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CRO_CargarDatos_G4", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CRO_CargarDatos_T1()

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand
        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("CRO_CargarDatos_T1", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.CommandTimeout = 0

        oComm.ExecuteNonQuery()

        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function CRO_Recuperar_P1() As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CRO_Recuperar_P1")
        Return oData
    End Function

    Public Function CRO_Recuperar_G1() As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CRO_Recuperar_G1")
        Return oData
    End Function

    Public Function CRO_Recuperar_G4() As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CRO_Recuperar_G4")
        Return oData
    End Function

    Public Function CRO_Recuperar_T1() As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CRO_Recuperar_T1")
        Return oData
    End Function

    Sub CRO_Insertar_TransfCartera(ByVal nTipCart As Integer, _
                            ByVal nTipTxAdqCart As Integer, _
                            ByVal nCodPersEmp As Integer, _
                            ByVal nTipCliEmp As Integer, _
                            ByVal nTipContrTxCart As Integer, _
                            ByVal nCodPersRL As Integer, _
                            ByVal nRelEmpInf As Integer, _
                            ByVal nRespEmpRep As Integer, _
                            ByVal dFecTxAdq As Date, _
                            ByVal nValIniTxAdq As Decimal, _
                            ByVal nMtoTxAdq As Decimal, _
                            ByVal sNroCliTxAdq As String, _
                            ByVal sResAutoriz As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("CRO_Insertar_TransfCartera ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nTipCart", SqlDbType.Int).Value = nTipCart
        oComm.Parameters.Add("@nTipTxAdqCart", SqlDbType.Int).Value = nTipTxAdqCart
        oComm.Parameters.Add("@nCodPersEmp", SqlDbType.Int).Value = nCodPersEmp
        oComm.Parameters.Add("@nTipCliEmp", SqlDbType.Int).Value = nTipCliEmp
        oComm.Parameters.Add("@nTipContrTxCart", SqlDbType.Int).Value = nTipContrTxCart
        oComm.Parameters.Add("@nCodPersRL", SqlDbType.Int).Value = nCodPersRL
        oComm.Parameters.Add("@nRelEmpInf", SqlDbType.Int).Value = nRelEmpInf
        oComm.Parameters.Add("@nRespEmpRep", SqlDbType.Int).Value = nRespEmpRep
        oComm.Parameters.Add("@dFecTxAdq", SqlDbType.SmallDateTime).Value = dFecTxAdq
        oComm.Parameters.Add("@nValIniTxAdq", SqlDbType.Decimal, 18).Value = nValIniTxAdq
        oComm.Parameters.Add("@nMtoTxAdq", SqlDbType.Decimal, 18).Value = nMtoTxAdq
        oComm.Parameters.Add("@cNroCliTxAdq", SqlDbType.VarChar, 20).Value = sNroCliTxAdq
        oComm.Parameters.Add("@cResAutoriz", SqlDbType.VarChar, 20).Value = sResAutoriz

        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Sub CRO_Modificar_TransfCartera(ByVal nCodCart As Integer, _
                            ByVal nTipCart As Integer, _
                            ByVal nTipTxAdqCart As Integer, _
                            ByVal nCodPersEmp As Integer, _
                            ByVal nTipCliEmp As Integer, _
                            ByVal nTipContrTxCart As Integer, _
                            ByVal nCodPersRL As Integer, _
                            ByVal nRelEmpInf As Integer, _
                            ByVal nRespEmpRep As Integer, _
                            ByVal dFecTxAdq As Date, _
                            ByVal nValIniTxAdq As Decimal, _
                            ByVal nMtoTxAdq As Decimal, _
                            ByVal sNroCliTxAdq As String, _
                            ByVal sResAutoriz As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("CRO_Modificar_TransfCartera ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCart", SqlDbType.Int).Value = nCodCart
        oComm.Parameters.Add("@nTipCart", SqlDbType.Int).Value = nTipCart
        oComm.Parameters.Add("@nTipTxAdqCart", SqlDbType.Int).Value = nTipTxAdqCart
        oComm.Parameters.Add("@nCodPersEmp", SqlDbType.Int).Value = nCodPersEmp
        oComm.Parameters.Add("@nTipCliEmp", SqlDbType.Int).Value = nTipCliEmp
        oComm.Parameters.Add("@nTipContrTxCart", SqlDbType.Int).Value = nTipContrTxCart
        oComm.Parameters.Add("@nCodPersRL", SqlDbType.Int).Value = nCodPersRL
        oComm.Parameters.Add("@nRelEmpInf", SqlDbType.Int).Value = nRelEmpInf
        oComm.Parameters.Add("@nRespEmpRep", SqlDbType.Int).Value = nRespEmpRep
        oComm.Parameters.Add("@dFecTxAdq", SqlDbType.SmallDateTime).Value = dFecTxAdq
        oComm.Parameters.Add("@nValIniTxAdq", SqlDbType.Decimal, 18).Value = nValIniTxAdq
        oComm.Parameters.Add("@nMtoTxAdq", SqlDbType.Decimal, 18).Value = nMtoTxAdq
        oComm.Parameters.Add("@cNroCliTxAdq", SqlDbType.VarChar, 20).Value = sNroCliTxAdq
        oComm.Parameters.Add("@cResAutoriz", SqlDbType.VarChar, 20).Value = sResAutoriz

        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Sub CRO_Eliminar_TransfCartera(ByVal nCodCart As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlClient.SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlClient.SqlCommand("CRO_Eliminar_TransfCartera ", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCart", SqlDbType.Int).Value = nCodCart
        oComm.ExecuteNonQuery()
        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function CRO_Recuperar_TransfCartera(ByVal dFechaI As Date, ByVal dFechaF As Date) As ComTipos.clsDataset

        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "CRO_Recuperar_TransfCartera '" & dFechaI & "','" & dFechaF & "'")
        Return oData
    End Function

    Public Function RecuperaConstantes(ByVal nCodigo As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Constante_Recupera " & nCodigo)
        Return oData
    End Function
End Class
