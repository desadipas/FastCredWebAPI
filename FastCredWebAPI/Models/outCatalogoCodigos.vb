Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Public Class outCatalogoCodigos

    Public Property nResultado As Integer
    Public Property sRespuesta As String

    Public Property ListaCatalogoCodigos As IQueryable(Of CatalogoCodigos)
    Public Property ListaOperaciones As IQueryable(Of operaciones)
    Public Property ListaZonas As IQueryable(Of Zonas)
    Public Property ListaProductos As IQueryable(Of Productos)
    Public Property ListaMenus As IQueryable(Of Menu)
    Public Property ListaSubMenus As IQueryable(Of MenuDetalle)
    Public Property ListaAgencias As IQueryable(Of Agencias)

End Class
