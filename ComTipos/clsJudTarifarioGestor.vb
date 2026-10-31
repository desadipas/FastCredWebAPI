Option Explicit On
<Serializable()> Public Class clsJudTarifarioGestor
    Public starif As String
    Public nestado As Integer

    Public ogrid As clsDataset  '
    Public ocabecera As clsDataset
    Public odetalle As clsDataset

    Public olista As clsDataset
    Public olistanueva As clsDataset
    Public onumerotarifarios As clsDataset
    Public Sub New()
        ogrid = New clsDataset
        ocabecera = New clsDataset
        odetalle = New clsDataset
        ''''''''''''''''''''''''''''''''       
        ''''''''''''''''''''''''''''''''
        olista = New clsDataset
        olistanueva = New clsDataset
        onumerotarifarios = New clsDataset
    End Sub
End Class
