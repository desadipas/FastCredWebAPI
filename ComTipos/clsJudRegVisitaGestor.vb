
'***********************************************************************
'CREADO POR: 
'USUARIO CREACION : 
'FECHA CREACION:
'USUARIO MODIFICACION:CCORDOVA
'FECHA DE MODIFICACION:28/05/2012
'COMENTARIO MODIFICACION:
'SE AGREGARON LAS FUNCIONES PARA AGREGAR LAS CAUSAS DE MOROSIDAD EN EL REGISTRO DE VISITA
'***********************************************************************
Option Explicit On
<Serializable()> Public Class clsJudRegVisitaGestor
    Public odatos As clsDataset
    Public ogestion As clsDataset
    Public oresultgestion As clsDataset
    Public oubiccred As clsDataset
    Public ocalifcred As clsDataset
    Public odirecc As clsDataset

    Public oCausaMorosidad As clsDataset 'ARCV 28-05-2012

    Public Sub New()
        odatos = New clsDataset
        ogestion = New clsDataset
        oresultgestion = New clsDataset
        oubiccred = New clsDataset
        ocalifcred = New clsDataset
        odirecc = New clsDataset

        oCausaMorosidad = New clsDataset 'ARCV 28-05-2012
    End Sub
End Class
