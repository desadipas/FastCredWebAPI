<Serializable()> Public Class ClsCatalogoCodigos
    Public oDatos() As clsDataset
    Public nFilas As Integer

    Public Sub CargaCombos(ByVal aCodigos As String(), Optional ByVal bCargaAgencias As Boolean = False)
        Dim i As Integer
        Dim j As Integer
        nFilas = UBound(aCodigos)
        ReDim oDatos(nFilas)

        For i = 0 To nFilas - 1
            oDatos(i) = New clsDataset
            oDatos(i).CargaTabla(, , "CredW_RecuperaCatalogoCodigos " & aCodigos(i))
        Next

        'CARGAR OTROS COMO AGENCIAS
        If bCargaAgencias Then
            j = nFilas
            j = j + 1
            ReDim Preserve oDatos(j)
            oDatos(j - 1) = New clsDataset
            oDatos(j - 1).CargaTabla(, , "CredW_Recupera_Agencias ")
        End If





    End Sub



End Class
