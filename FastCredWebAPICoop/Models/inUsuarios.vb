
Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports System.Data.Linq.SqlClient
Imports System.Data.Entity.Core.Objects
Imports FastCredWebAPI.FastCredWebAPI
Public Class inUsuarios

    Public Property sToken As String
    'nSolicitud -> 1 busqueda por nombres y apellidos
    'nSolicitud -> 2 Cambiar Clave
    'nSolicitud -> 3 forzar cambio de clave siguiente sesion
    Public Property nSolicitud As Integer

    Public Property cCodUsu As String
    Public Property cNombresApe As String
    Public Property cNuevaClave As String
    Public Property oUsuarios As Usuarios

    Private db As New DBNEGOCIOEntities

    Public Function RecuperaSolicitud() As outUsuarios

        Dim R As outUsuarios
        Dim S As iSesion

        S = New iSesion()
        R = New outUsuarios()

        If Not S.ValidaToken(sToken) Then
            R.nResultado = 0
            R.sRespuesta = "Token Invalido"
            Return R
        End If


        If nSolicitud = -1 Then

            R.ListaUsuarios = db.Usuarios
        Else

            Dim dUsuarios = From U In db.Usuarios
                            Where U.cNombres.Contains(cNombresApe)
                            Select U

            R.ListaUsuarios = dUsuarios

        End If

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"


        Return R

    End Function

    Public Function CambiarClave() As outUsuarios

        Dim R As outUsuarios
        Dim S As iSesion
        Dim U As Usuarios = New Usuarios()
        R = New outUsuarios()

        S = New iSesion()

        If Not S.ValidaToken(sToken) Then
            R.nResultado = 0
            R.sRespuesta = "Token Invalido"
            Return R
        End If


        db.Configuration.AutoDetectChangesEnabled = False

        U.cCodUsu = cCodUsu
        U = db.Usuarios.First(Function(p) p.cCodUsu = cCodUsu)

        db.Usuarios.Attach(U)
        U.sContrasenia = cNuevaClave
        U.bCambioClave = 0
        db.Entry(U).Property("sContrasenia").IsModified = True
        db.Entry(U).Property("bCambioClave").IsModified = True
        db.SaveChanges()

        db.Configuration.AutoDetectChangesEnabled = True

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"

        Return R

    End Function

    Public Function ForzarCambioClave() As outUsuarios
        Dim R As outUsuarios
        Dim S As iSesion
        Dim U As Usuarios = New Usuarios()
        R = New outUsuarios()

        S = New iSesion()

        If Not S.ValidaToken(sToken) Then
            R.nResultado = 0
            R.sRespuesta = "Token Invalido"
            Return R
        End If


        db.Configuration.AutoDetectChangesEnabled = False

        U.cCodUsu = cCodUsu
        U = db.Usuarios.First(Function(p) p.cCodUsu = cCodUsu)

        db.Usuarios.Attach(U)
        U.bCambioClave = 1
        db.Entry(U).Property("bCambioClave").IsModified = True
        db.SaveChanges()

        db.Configuration.AutoDetectChangesEnabled = True

        R.nResultado = 1
        R.sRespuesta = "Proceso Realizado de manera Correcta"

        Return R

    End Function

End Class
