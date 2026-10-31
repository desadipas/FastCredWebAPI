
Public Class ClsDataSetPage
    Private Datos As clsDataset
    Private DatosTotales As clsDataset
    Public nNumPaginas As Integer
    Public nCurPag As Integer
    Private nNumRegIni As Integer
    Private nNumRegfin As Integer


    Public Sub CargaTabla(Optional ByVal psNomTab As String = "", Optional ByVal psWhere As String = "", _
                            Optional ByVal psCadStore As String = "", _
                            Optional ByVal pnConexion As Integer = 1, _
                            Optional ByVal pbSinTrim As Boolean = False)

        DatosTotales = New clsDataset
        DatosTotales.CargaTabla(psNomTab, psWhere, psCadStore, pnConexion, pbSinTrim)

        If DatosTotales.nNumFilas > 0 Then
            DatosTotales.GeneraColumnaItem()
            nNumPaginas = (DatosTotales.nNumFilas / 39) + 1
            nNumRegIni = 1
            nNumRegfin = 0
            nCurPag = 0
        Else
            nNumPaginas = 0
            nNumRegIni = 0
            nNumRegfin = 0
            nCurPag = 0
        End If

        Datos = Nothing

    End Sub

    Public Function RecuperaPagina(ByVal pnNumPag As Integer) As clsDataset
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Datos = Nothing
        Datos = New clsDataset
        Datos.CrearVacio(DatosTotales.Columna)

        If pnNumPag = 1 Then
            i = 1
        Else
            i = ((pnNumPag - 1) * 30) + 1
        End If
        nNumRegIni = i

        j = 1
        DatosTotales.Seek(i)
        Do While Not DatosTotales.EOF And i <= nNumRegIni + 29
            Datos.AddRegistro()
            Datos.Seek(j)
            For k = 1 To Datos.nNumColumn
                Datos.AsignarDato(DatosTotales.ObtenerDato(Datos.Columna(k)), Datos.Columna(k))
            Next
            j = j + 1
            DatosTotales.Siguiente()
            i = i + 1
        Loop

        nNumRegfin = i - 1
        nCurPag = pnNumPag

        RecuperaPagina = Datos


    End Function




End Class
