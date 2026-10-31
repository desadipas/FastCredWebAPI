Imports System.Data
Imports System.Data.SqlClient
Public Class ClsExamen

    Public Function Ex_RegistraTema(ByVal pcNomTema As String, ByVal pcIndicaciones As String, ByVal pcCodUsu As String, _
                                    ByVal pbTiempo As Boolean, ByVal pnTipoEx As Integer, ByVal pnCodTema As Integer) As Integer
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Ex_GrabaTemas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cNomTema", SqlDbType.VarChar, 250).Value = pcNomTema
        oComm.Parameters.Add("@cIndicaciones", SqlDbType.VarChar, 500).Value = pcIndicaciones
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@bTiempo", SqlDbType.Bit).Value = pbTiempo
        oComm.Parameters.Add("@nCodT", SqlDbType.Int).Value = pnCodTema
        oComm.Parameters.Add("@nTipoEx", SqlDbType.Int).Value = pnTipoEx
        oComm.Parameters.Add("@nCodTem", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        Ex_RegistraTema = oComm.Parameters("@nCodTem").Value

        oComm = Nothing
        objCon = Nothing
    End Function

    Public Sub Ex_RegistraPreguntas(ByVal pnCodTema As Integer, ByVal pcPregunta As String, _
                                    ByVal pnTiempo As Integer, ByVal pnPuntaje As Integer, ByVal pcCodUsu As String, _
                                    ByVal oAlternativas As ComTipos.clsDataset, ByVal pnCodPregunta As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Ex_GrabaPreguntas", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodTema", SqlDbType.Int).Value = pnCodTema
        oComm.Parameters.Add("@cPregunta", SqlDbType.VarChar, 500).Value = pcPregunta
        oComm.Parameters.Add("@nTiempo", SqlDbType.Int).Value = pnTiempo
        oComm.Parameters.Add("@nPuntaje", SqlDbType.Int).Value = pnPuntaje
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@oAlternativas", SqlDbType.Xml).Value = oAlternativas.DevuelveXML
        oComm.Parameters.Add("@nCodPregunta", SqlDbType.Int).Value = pnCodPregunta

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function Ex_DevuelvePreguntas(ByVal pnCodTema As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_DevuelveListaPreguntas_del_Tema " & pnCodTema)
        Ex_DevuelvePreguntas = oData
        oData = Nothing
    End Function

    Public Function Ex_DevuelveAlternativas(ByVal pnCodTema As Integer, ByVal pnCodPreg As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_DevuelveAlternativas_Pregunta " & pnCodTema & ", " & pnCodPreg)
        Ex_DevuelveAlternativas = oData
        oData = Nothing
    End Function

    Public Function Ex_DevuelveTemas(ByVal pcTema As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_BuscaTemas '" & pcTema & "'")
        Ex_DevuelveTemas = oData
        oData = Nothing
    End Function

    Public Function Ex_DevuelveListaTemas() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_ListaTemas ")
        Ex_DevuelveListaTemas = oData
        oData = Nothing
    End Function

    Public Sub Ex_RegistraExamen(ByVal pnCodTema As String, ByVal pdFecIni As Date, _
    ByVal pdFecFin As DateTime, ByVal pnTiempo As Integer, ByVal pcIndicaciones As String, _
    ByVal pcCodUsu As String, ByVal pnCodExa As Integer)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("Ex_GrabaExamen", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@nCodTema", SqlDbType.Int).Value = pnCodTema
        oComm.Parameters.Add("@dFecIni", SqlDbType.DateTime).Value = pdFecIni
        'oComm.Parameters.Add("@dHoraIni", SqlDbType.DateTime).Value = pdHoraIni
        oComm.Parameters.Add("@dFecFin", SqlDbType.DateTime).Value = pdFecFin
        'oComm.Parameters.Add("@dHoraFin", SqlDbType.DateTime).Value = pdHoraFin
        oComm.Parameters.Add("@nTiempo", SqlDbType.Int).Value = pnTiempo
        oComm.Parameters.Add("@cIndicaciones", SqlDbType.VarChar, 500).Value = pcIndicaciones
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@nCodExa", SqlDbType.Int).Value = pnCodExa

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function Ex_DevuelveExamenes(ByVal pcTema As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_BuscaExamen '" & pcTema & "'")
        Ex_DevuelveExamenes = oData
        oData = Nothing
    End Function

    Public Function Ex_DevuelveExamenesActivos() As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_ExamenesActivos ")
        Ex_DevuelveExamenesActivos = oData
        oData = Nothing
    End Function

    Public Function Ex_PreguntasExamen(ByVal pnCodExa As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_PreguntasExamen " & pnCodExa)
        Ex_PreguntasExamen = oData
        oData = Nothing
    End Function

    Public Function Ex_AlternativasExamen(ByVal pnCodExa As Integer, ByVal pnCodPreg As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_AlternativasExamen " & pnCodExa & ", " & pnCodPreg)
        Ex_AlternativasExamen = oData
        oData = Nothing
    End Function

    Public Function Ex_TotalPreguntasExamen(ByVal pnCodExa As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_TotalPreguntasExamen " & pnCodExa)
        Ex_TotalPreguntasExamen = oData
        oData = Nothing
    End Function

    Public Sub Ex_RegistraResultados(ByVal pcCodUsu As String, ByVal pnCodExa As Integer, ByVal oResultados As ComTipos.clsDataset)
        Dim objCon As ComConecta.clsConecta
        Dim oComm As SqlCommand

        objCon = New ComConecta.clsConecta
        oComm = New SqlCommand("EX_GrabaResultados", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure
        oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 20).Value = pcCodUsu
        oComm.Parameters.Add("@nCodExa", SqlDbType.Int).Value = pnCodExa
        oComm.Parameters.Add("@oResultados", SqlDbType.Xml).Value = oResultados.DevuelveXML

        oComm.ExecuteNonQuery()

        oComm = Nothing
        objCon = Nothing
    End Sub

    Public Function Ex_ValidaExamenUsuario(ByVal pnCodExa As Integer, ByVal pcCodUsu As String) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Ex_ValidaExamenUsuario " & pnCodExa & ", '" & pcCodUsu & "'")
        Ex_ValidaExamenUsuario = oData
        oData = Nothing
    End Function
End Class
