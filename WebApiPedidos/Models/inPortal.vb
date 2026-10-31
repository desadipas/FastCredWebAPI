Imports System.Data.SqlClient

Public Class inPortal
    Public Property nSolicitud As Integer
    Public Property oZona As clsZona
    Public Property nIDZona As Integer
    Public Property cCodUsu As String
    Public Property oUbicacion As clsUbicacion
    Public Property oUsuario As clsUsuario
    Public Property cClave As String
    Public Property nCodRol As Integer
    Public Property nIDCom As Integer
    Public Property nNumTran As Integer
    Public Property nCantidadNueva As Integer
    Public Property nAnular As Integer

    Public Function RegistraCambioEntrega() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            If nAnular = 1 Then

                oComm = New SqlClient.SqlCommand("PA_AnulaEntrega", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nNumTran", SqlDbType.Int).Value = nNumTran

                oComm.ExecuteNonQuery()

            Else

                oComm = New SqlClient.SqlCommand("PA_RegistrarCambioEntrega", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nNumTran", SqlDbType.Int).Value = nNumTran
                oComm.Parameters.Add("@psUser", SqlDbType.VarChar, 50).Value = cCodUsu
                oComm.Parameters.Add("@nCantidadNueva", SqlDbType.Int).Value = nCantidadNueva


                oComm.ExecuteNonQuery()

            End If



            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function



    Public Function RecuperaEntregaACambiar() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer
        Dim oPagoAnulacion As clsAnulacionPago

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaEntregaACambiar", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nNroTrans", SqlDbType.Int).Value = nNumTran


            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            oPagoAnulacion = New clsAnulacionPago
            While dr.Read

                oPagoAnulacion.cCodUsu = dr.Item("cUser")
                oPagoAnulacion.cNomComercio = dr.Item("cRazonSocial")
                oPagoAnulacion.dFecha = dr.Item("dFecha")
                oPagoAnulacion.nCantidad = dr.Item("nMermaEntregada")
                oPagoAnulacion.nIDCom = dr.Item("nIDCom")
                oPagoAnulacion.nIdOrden = dr.Item("nIDOrden")
                oPagoAnulacion.nMonto = dr.Item("nMontoCobrado")
                oPagoAnulacion.nNumTran = dr.Item("nNroTrans")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oPagoAnulacion = oPagoAnulacion

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaEntregas() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaEntregas", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIDCom", SqlDbType.Int).Value = nIDCom
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oCodEntregas(i)
                R.oCodEntregas(i) = dr.Item("nNroTrans")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RegistraAnulaPago() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_AnulaPago", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nNumTran", SqlDbType.Int).Value = nNumTran

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RecuperaPagoParaAnulacion() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i As Integer
        Dim oPagoAnulacion As clsAnulacionPago

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaPagoParaAnular", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIdCom", SqlDbType.Int).Value = nIDCom
            oComm.Parameters.Add("@dFecha", SqlDbType.Date).Value = Now()

            oComm.ExecuteNonQuery()

            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            oPagoAnulacion = New clsAnulacionPago
            While dr.Read

                oPagoAnulacion.cCodUsu = dr.Item("cUser")
                oPagoAnulacion.cNomComercio = dr.Item("cRazonSocial")
                oPagoAnulacion.dFecha = dr.Item("dFecha")
                oPagoAnulacion.nCantidad = dr.Item("nMermaEntregada")
                oPagoAnulacion.nIDCom = dr.Item("nIDCom")
                oPagoAnulacion.nIdOrden = dr.Item("nIDOrden")
                oPagoAnulacion.nMonto = dr.Item("nMontoCobrado")
                oPagoAnulacion.nNumTran = dr.Item("nNroTrans")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            R.oPagoAnulacion = oPagoAnulacion

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function RecuperaComercios() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaComercios", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oComercios(i)
                R.oComercios(i) = New clsComercio
                R.oComercios(i).nIDCom = dr.Item("nIDCOM")
                R.oComercios(i).cRazonSocial = dr.Item("cRazonSocial")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function ValidaRol() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ValidaRol", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("psCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
            oComm.Parameters.Add("@nCodRol", SqlDbType.Int).Value = nCodRol
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nCorrecto = oComm.Parameters("@pnResp").Value

            oConex.Close()

            If nCorrecto = 1 Then
                D(0) = "1"
                D(1) = "Acceso Correcto"
            Else
                D(0) = "0"
                D(1) = "Clave o Usuario Invalidos"
            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RegistraUsuario() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraUsuario_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = oUsuario.cCodUsu
            oComm.Parameters.Add("@cNomApe", SqlDbType.VarChar, 150).Value = oUsuario.cNomApe
            oComm.Parameters.Add("@cClave", SqlDbType.VarChar, 50).Value = oUsuario.cClave
            oComm.Parameters.Add("@pnCodRol", SqlDbType.Int).Value = nCodRol
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = oUsuario.nCodAge


            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function




    Public Function ValidaClave() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ValidaClave", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
            oComm.Parameters.Add("@cClave", SqlDbType.VarChar, 50).Value = cClave
            oComm.Parameters.Add("@nCorrecto", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nCorrecto = oComm.Parameters("@nCorrecto").Value

            oConex.Close()

            If nCorrecto = 1 Then
                D(0) = "1"
                D(1) = "Acceso Correcto"
            Else
                D(0) = "0"
                D(1) = "Clave o Usuario Invalidos"
            End If


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function DesAsignaZona() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_DesAsignaZona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
            oComm.Parameters.Add("@nIDZona", SqlDbType.Int).Value = nIDZona

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

    Public Function AsignaZona() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_AsignaZona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cCodUsu
            oComm.Parameters.Add("@nIDZona", SqlDbType.Int).Value = nIDZona

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function



    Public Function RecuperaUsuariosDeZona() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaUsuariosDeZona", oConex)
            oComm.Parameters.Add("@nIDZona", SqlDbType.Int).Value = nIDZona
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oUsuarios(i)
                R.oUsuarios(i) = dr.Item("cCodUsu")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function RecuperaAgencias() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaAgencias", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oAgencias(i)
                R.oAgencias(i) = New clsAgencia
                R.oAgencias(i).cNomAge = dr.Item("cNomAge")
                R.oAgencias(i).nCodAge = dr.Item("nCodAge")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function

    Public Function RecuperaTodosLosUsuarios() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaTodosLosUsuarios", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oUsuarios(i)
                R.oUsuarios(i) = dr.Item("cCodUsu")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function RecuperaUbicacionesDeUsuario() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaUbicacionesDeUsuario", oConex)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = oUbicacion.cCodUsu
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oUbicacionUsu(i)
                R.oUbicacionUsu(i) = New clsUbicacion
                R.oUbicacionUsu(i).dFecha = dr.Item("dFecha")
                R.oUbicacionUsu(i).cCodUsu = dr.Item("cCodUsu")
                R.oUbicacionUsu(i).nID = dr.Item("nID")
                R.oUbicacionUsu(i).nLat = dr.Item("nLat")
                R.oUbicacionUsu(i).nLong = dr.Item("nLong")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function RecuperaUbicaciones() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaUbicacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0


            While dr.Read
                ReDim Preserve R.oUbicacionUsu(i)
                R.oUbicacionUsu(i) = New clsUbicacion
                R.oUbicacionUsu(i).dFecha = dr.Item("dFecha")
                R.oUbicacionUsu(i).cCodUsu = dr.Item("cCodUsu")
                R.oUbicacionUsu(i).nID = dr.Item("nID")
                R.oUbicacionUsu(i).nLat = dr.Item("nLat")
                R.oUbicacionUsu(i).nLong = dr.Item("nLong")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R

    End Function


    Public Function RegistraUbicacion() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraUbicacion", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = oUbicacion.cCodUsu
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = oUbicacion.nLat
            oComm.Parameters.Add("@nlong", SqlDbType.Float).Value = oUbicacion.nLong

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function EliminaZona() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_EliminaZona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIDZona", SqlDbType.Int).Value = nIDZona

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Eliminacón Correcta"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


    Public Function RecuperaZonas() As outPortal
        Dim R As New outPortal
        Dim i As Integer
        Dim D(2) As String

        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Try


            oComm = New SqlClient.SqlCommand("PA_RecuperaZonas", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0
            While dr.Read

                ReDim Preserve R.oZonas(i)
                R.oZonas(i) = New clsZona
                R.oZonas(i).nIdZona = dr.Item("nIDZona")
                R.oZonas(i).cNomZona = UCase(dr.Item("cNomZona"))
                R.oZonas(i).nLat1 = dr.Item("nLat1")
                R.oZonas(i).nLong1 = dr.Item("nLong1")
                R.oZonas(i).nLat2 = dr.Item("nLat2")
                R.oZonas(i).nLong2 = dr.Item("nLong2")
                R.oZonas(i).nLat3 = dr.Item("nLat3")
                R.oZonas(i).nLong3 = dr.Item("nLong3")
                R.oZonas(i).nLat4 = dr.Item("nLat4")
                R.oZonas(i).nLong4 = dr.Item("nLong4")

                i = i + 1
            End While
            dr.Close()
            dr = Nothing

            oConex.Close()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message

        End Try

        R.sRespuesta = D(1)
        R.nResultado = CInt(D(0))

        Return R




    End Function

    Public Function RegistraZona() As outPortal
        Dim R As New outPortal
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim i, j As Integer
        Dim nCorrecto As Integer
        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraZona", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@cNomZona", SqlDbType.VarChar, 50).Value = oZona.cNomZona
            oComm.Parameters.Add("@nLat1", SqlDbType.Float).Value = oZona.nLat1
            oComm.Parameters.Add("@nLong1", SqlDbType.Float).Value = oZona.nLong1
            oComm.Parameters.Add("@nLat2", SqlDbType.Float).Value = oZona.nLat2
            oComm.Parameters.Add("@nLong2", SqlDbType.Float).Value = oZona.nLong2
            oComm.Parameters.Add("@nLat3", SqlDbType.Float).Value = oZona.nLat3
            oComm.Parameters.Add("@nLong3", SqlDbType.Float).Value = oZona.nLong3
            oComm.Parameters.Add("@nLat4", SqlDbType.Float).Value = oZona.nLat4
            oComm.Parameters.Add("@nLong4", SqlDbType.Float).Value = oZona.nLong4

            oComm.ExecuteNonQuery()


            oConex.Close()


            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try


        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function

End Class
