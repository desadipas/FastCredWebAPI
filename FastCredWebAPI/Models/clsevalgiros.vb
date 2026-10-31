Public Class clsevalgiros

    Public Property nIDGiro As Long
    Public Property nIDEVal As Long
    Public Property nCodGiro As Integer
    Public Property cDescripcion As String
    Public Property nVentas As Double
    Public Property nCostoVentas As Double
    Public Property nMargen As Double
    Public Property nPorcPart As Double
    Public Property nUtilidad As Double
    Public Property nGasOper As Double
    Public Property nPorcCred As Double
    Public Property nVentaCred As Double
    Public Property nProdTermin As Double
    Public Property nProdProceso As Double
    Public Property nMateriales As Double
    Public Property oEvalCosteoProd As clsEvalCosteoProd()
    Public Property oEvalinventario As clsEvalinventario()
    Public Property oEvalgasoper As clsevalgasoper()
    Public Property oEvalcostoventa As clsevalcostoventa()
    Public Property oEvalventas As clsevalventas()

End Class
