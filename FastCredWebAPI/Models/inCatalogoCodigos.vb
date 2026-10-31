Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Public Class inCatalogoCodigos

    Public Property sToken As String
    'nSolicitud = -1 -> Toda la tabla
    'nSolicitud = 1 -> filtrado por ncodigo
    'nSolicitud = 2 -> Registrar Nuevo Codigo
    'nSolicitud = 3 -> Actualizar Nuevo Codigo
    Public Property nSolicitud As Integer
    Public Property nCodCatalogo As Integer
    Public Property oCatalogo As CatalogoCodigos

    Private db As New DBNEGOCIOEntities

    Public Function RecuperaSolicitud() As outCatalogoCodigos
        Dim R As outCatalogoCodigos
        Dim S As iSesion

        S = New iSesion()

        If Not S.ValidaToken(sToken) Then
            R.nResultado = 0
            R.sRespuesta = "Token Invalido"
        End If

        R = New outCatalogoCodigos()

        If nSolicitud = -1 Then

            R.ListaCatalogoCodigos = db.CatalogoCodigos
            R.ListaOperaciones = db.Operaciones
            R.ListaZonas = db.Zonas
            R.ListaMenus = db.Menu
            R.ListaSubMenus = db.MenuDetalle
            R.ListaAgencias = db.Agencias
        Else

            R.ListaCatalogoCodigos = db.CatalogoCodigos.Find(nCodCatalogo)
        End If

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"


        Return R

    End Function



End Class
