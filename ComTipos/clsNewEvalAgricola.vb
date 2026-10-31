Option Explicit On
<Serializable()> Public Class clsNewEvalAgricola
    Public oDatosCredAgri As clsDataset
    Public oTecnologia As clsDataset
    Public oVariedad As clsDataset
    Public oUniMedida As clsDataset
    Public oLaboresCultivo As clsDataset
    Public oMaquinaria As clsDataset
    Public oSemilla As clsDataset
    Public oFertilizantes As clsDataset
    Public oHerbicidas As clsDataset
    Public oAgua As clsDataset
    Public oCostosIndirectos As clsDataset


    Public Sub New()
        oDatosCredAgri = New clsDataset
        oTecnologia = New clsDataset
        oVariedad = New clsDataset
        oLaboresCultivo = New clsDataset
        oMaquinaria = New clsDataset
        oSemilla = New clsDataset
        oFertilizantes = New clsDataset
        oHerbicidas = New clsDataset
        oAgua = New clsDataset
        oCostosIndirectos = New clsDataset
        oUniMedida = New clsDataset
    End Sub
End Class


