
Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports System.Data.Linq.SqlClient
Imports System.Data.Entity.Core.Objects
Imports FastCredWebAPI.FastCredWebAPI
Public Class inCredGastosOperacion

    Public Property sToken As String
    'nSolicitud -> 1 busqueda por producto
    'nSolicitud -> 2 Agregar por producto
    'nSolicitud -> 3 Actualizar por producto
    Public Property nSolicitud As Integer
    Public Property nProd As Integer

    'Public Property oCredGastosOperacion As CredGastosOperacion()

    Private db As New DBNEGOCIOEntities

    Public Function RecuperaSolicitud() As outCredGastosOperacion
        Dim R As outCredGastosOperacion
        Dim S As iSesion

        'Try

        '    S = New iSesion()

        '    If Not S.ValidaToken(sToken) Then
        '        R.nResultado = 0
        '        R.sRespuesta = "Token Invalido"
        '    End If

        '    R = New outCredGastosOperacion()


        '    R.oCredGastosOperacion = From D In db.CredGastosOperacion Where D.nProd = nProd Select D

        '    R.nResultado = 1
        '    R.sRespuesta = "Proceso Realizado de manera Correcta"


        '    Return R

        'Catch ex As Exception
        '    R.nResultado = 1
        '    R.sRespuesta = ex.Message
        '    Return R
        'End Try

    End Function

    Public Function RegistraNuevoGasto() As outCredGastosOperacion
        Dim R As outCredGastosOperacion
        Dim S As iSesion
        Dim i As Integer

        'Dim oCredGastoOperItem As CredGastosOperacion
        'Try
        '    S = New iSesion()

        '    If Not S.ValidaToken(sToken) Then
        '        R.nResultado = 0
        '        R.sRespuesta = "Token Invalido"
        '    End If

        '    R = New outCredGastosOperacion()

        '    For i = 0 To UBound(oCredGastosOperacion) - 1
        '        oCredGastoOperItem = oCredGastosOperacion(i)
        '        If nSolicitud = 2 Then
        '            oCredGastoOperItem = db.CredGastosOperacion.Add(oCredGastoOperItem)
        '            db.SaveChanges()

        '        ElseIf nSolicitud = 3 Then

        '            db.Entry(oCredGastoOperItem).State = EntityState.Modified
        '            Try
        '                db.SaveChanges()
        '            Catch ex As DbUpdateConcurrencyException
        '                R.nResultado = 0
        '                R.sRespuesta = ex.Message
        '                Return R
        '            End Try

        '        Else
        '            R.nResultado = 0
        '            R.sRespuesta = "Solicitud Invalida"
        '            Return R
        '        End If
        '    Next

        '    R.nResultado = 1
        '    R.sRespuesta = "Proceso Realizado de manera Correcta"

        '    Return R
        'Catch ex As Exception
        '    R.nResultado = 1
        '    R.sRespuesta = ex.Message
        '    Return R
        'End Try

    End Function


End Class
