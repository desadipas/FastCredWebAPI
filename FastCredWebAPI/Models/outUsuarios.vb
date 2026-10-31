Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Public Class outUsuarios

    Public Property nResultado As Integer
    Public Property sRespuesta As String

    Public Property ListaUsuarios As IQueryable(Of Usuarios)

End Class
