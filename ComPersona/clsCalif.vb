Imports System.Data
Imports System.Data.SqlClient

Public Class clsCalif

    Public Sub PersonaCalif_Insertar(ByVal nCodPers As Integer, _
                        ByVal dFecha As Date, _
                        ByVal sCalif As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PersonaCalif_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@cCalif", SqlDbType.Char, 1).Value = sCalif

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersonaCalif_Modificar(ByVal nCodPers As Integer, _
                        ByVal dFecha As Date, _
                        ByVal sCalif As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PersonaCalif_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@cCalif", SqlDbType.Char, 1).Value = sCalif

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub PersonaCalif_Eliminar(ByVal nCodPers As Integer, _
                    ByVal dFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_PersonaCalif_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodPers", SqlDbType.Int).Value = nCodPers
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function PersonaCalif_Listar(ByVal nCodPers As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_PersonaCalif_Listar " & nCodPers)
        Return oData

    End Function

    Public Function PersonaCalif_CargaInterfaz() As ComTipos.clsDataset

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        Return C.RecuperaConstantes(2166)
        C = Nothing
    End Function

    Public sub TablaCalif_CargaInterfaz(byref oCalif as ComTipos.clsDataset , _
                        byref oProd as ComTipos.clsDataset , _
                        byref oGarant as ComTipos.clsDataset ) 

        Dim C As ComConstante.clsConstante
        C = New ComConstante.clsConstante
        oCalif = C.RecuperaConstantes(2166)
        oGarant = C.RecuperaConstantes(2150)
        oProd = C.RecuperaConstantes(1500)
        C = Nothing
    End Sub

    Public Function Listar_TablaCalif() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Listar_TablaCalif ")
        Return oData
    End Function

    Public Sub CredTablaCalif_Insertar(ByVal nCalif As Integer, _
                    ByVal nProducto As Integer, _
                    ByVal nGarant As Integer, _
                    ByVal bRefinanciado As Integer, _
                    ByVal nPorcen As Double, _
                    ByVal nRangoIni As Integer, _
                    ByVal nRangoFin As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredTablaCalif_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCalif", SqlDbType.Int).Value = nCalif
        oComm.Parameters.Add("@nProducto", SqlDbType.Int).Value = nProducto
        oComm.Parameters.Add("@nGarant", SqlDbType.Int).Value = nGarant
        oComm.Parameters.Add("@bRefinanciado", SqlDbType.Bit).Value = bRefinanciado
        oComm.Parameters.Add("@nPorcen", SqlDbType.Money).Value = nPorcen
        oComm.Parameters.Add("@nRangoIni", SqlDbType.Int).Value = nRangoIni
        oComm.Parameters.Add("@nRangoFin", SqlDbType.Int).Value = nRangoFin

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredTablaCalif_Modificar(ByVal nCalif As Integer, _
                ByVal nProducto As Integer, _
                ByVal nGarant As Integer, _
                ByVal bRefinanciado As Integer, _
                ByVal nPorcen As Double, _
                ByVal nRangoIni As Integer, _
                ByVal nRangoFin As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredTablaCalif_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCalif", SqlDbType.Int).Value = nCalif
        oComm.Parameters.Add("@nProducto", SqlDbType.Int).Value = nProducto
        oComm.Parameters.Add("@nGarant", SqlDbType.Int).Value = nGarant
        oComm.Parameters.Add("@bRefinanciado", SqlDbType.Bit).Value = bRefinanciado
        oComm.Parameters.Add("@nPorcen", SqlDbType.Money).Value = nPorcen
        oComm.Parameters.Add("@nRangoIni", SqlDbType.Int).Value = nRangoIni
        oComm.Parameters.Add("@nRangoFin", SqlDbType.Int).Value = nRangoFin

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub CredTablaCalif_Eliminar(ByVal nCalif As Integer, _
            ByVal nProducto As Integer, _
            ByVal nGarant As Integer, _
            ByVal bRefinanciado As Integer)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_CredTablaCalif_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCalif", SqlDbType.Int).Value = nCalif
        oComm.Parameters.Add("@nProducto", SqlDbType.Int).Value = nProducto
        oComm.Parameters.Add("@nGarant", SqlDbType.Int).Value = nGarant
        oComm.Parameters.Add("@bRefinanciado", SqlDbType.Bit).Value = bRefinanciado

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Function ProvisionVoluntaria_Listar(ByVal nCodCred As Integer, ByVal nCodAge As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_ProvisionVoluntaria_Listar " & nCodCred & "," & nCodAge)
        Return oData
    End Function

    Public Sub ProvisionVoluntaria_Insertar(ByVal nCodCred As Integer, _
                    ByVal nCodAge As Integer, _
                    ByVal dFecha As Date, _
                    ByVal nProvision As Double, _
                    ByVal sUser As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ProvisionVoluntaria_Insertar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nProvision", SqlDbType.Money).Value = nProvision
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ProvisionVoluntaria_Modificar(ByVal nCodCred As Integer, _
                    ByVal nCodAge As Integer, _
                    ByVal dFecha As Date, _
                    ByVal nProvision As Double, _
                    ByVal sUser As String)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ProvisionVoluntaria_Modificar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha
        oComm.Parameters.Add("@nProvision", SqlDbType.Money).Value = nProvision
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 20).Value = sUser

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

    Public Sub ProvisionVoluntaria_Eliminar(ByVal nCodCred As Integer, _
                    ByVal nCodAge As Integer, _
                    ByVal dFecha As Date)

        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Cred_ProvisionVoluntaria_Eliminar", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = nCodCred
        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = nCodAge
        oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = dFecha

        oComm.ExecuteNonQuery()
        objCon = Nothing
        oComm = Nothing
    End Sub

End Class
