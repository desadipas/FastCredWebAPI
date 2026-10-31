
Imports System.Data
Imports System.Data.SqlClient

Public Class clsConstante

    Public Function RecuperaConstantes(ByVal nCodigo As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla(, , "Cred_Constante_Recupera " & nCodigo)
        Return oData
    End Function


    Public Function RecuperaConstantes2(ByVal nCodigo As Integer) As ComTipos.clsDataset
        Dim oData As New ComTipos.clsDataset
        oData.CargaTabla("CatalogoCodigos", "nCodigo=" & nCodigo & " AND nValor<>" & nCodigo)
        Return oData
    End Function

    'Public Sub RecuperaConstantesCombo(ByVal nCodigo As Integer, _
    '                                ByRef oCombo() As ComTipos.clsTipoCombo)
    '    Dim oComm As SqlCommand
    '    Dim oCon As ComConecta.clsConecta
    '    Dim oData As SqlDataReader
    '    Dim i As Integer

    '    oCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand("Cred_Constante_Recupera", oCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@nCodigo", SqlDbType.Int)
    '    oComm.Parameters("@nCodigo").Value = nCodigo
    '    oData = oComm.ExecuteReader(CommandBehavior.SingleResult)

    '    ReDim oCombo(0)
    '    i = 0
    '    While oData.Read
    '        i = i + 1
    '        oCombo(i) = New ComTipos.clsTipoCombo
    '        ReDim Preserve oCombo(i)
    '        oCombo(i).sCodigo = oData.GetString(0)
    '        oCombo(i).sDescripcion = oData.GetString(1)
    '    End While
    '    oCon.CierraConexion()
    '    oComm = Nothing

    'End Sub

    'Public Function RecuperaConstantes(ByVal nCodigo As Integer) As SqlDataReader
    '    Dim oComm As SqlCommand
    '    Dim oCon As ComConecta.clsConecta
    '    Dim oData As SqlDataReader

    '    oCon = New ComConecta.clsConecta
    '    oComm = New SqlCommand
    '    oComm = New SqlCommand("Cred_Constante_Recupera", oCon.DameConexion)
    '    oComm.CommandType = CommandType.StoredProcedure
    '    oComm.Parameters.Add("@nCodigo", SqlDbType.Int).Value = nCodigo
    '    oData = oComm.ExecuteReader(CommandBehavior.SingleResult)
    '    'oData.Close()
    '    'oCon.CierraConexion()
    '    oComm = Nothing

    '    Return oData

    'End Function

    'Public Function RecuperaConstantes(ByVal nCodigo As Integer) As DataTable

    '    Dim dt As New DataTable
    '    Dim oCon As ComConecta.clsConecta
    '    oCon = New ComConecta.clsConecta

    '    Dim da As New SqlDataAdapter("Cred_Constante_Recupera", oCon.DameConexion)
    '    da.SelectCommand.CommandType = CommandType.StoredProcedure
    '    da.SelectCommand.Parameters.Add("@nCodigo", SqlDbType.Int)
    '    da.SelectCommand.Parameters("@nCodigo").Value = nCodigo

    '    da.Fill(dt)
    '    oCon.CierraConexion()
    '    Return dt

    'End Function
End Class
