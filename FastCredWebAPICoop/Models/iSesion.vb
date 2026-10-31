Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI

Public Class iSesion

    Public Property cCodUsu As String
    Public Property cClave As String
    'App Web -> nub32019azure2019
    'App movil -> mov1l2019@seguraperu
    Public Property cClaveApp As String
    Public Property cImei As String
    '1-> Web
    '2->Movil
    Public Property cTipoApl As String

    Public Property cClaveNueva As String

    '1-> Login
    '2-> Cambio de Clave
    Public Property nTipoSolicitud As String

    Private db As New DBNEGOCIOEntities

    Public Function ValidaToken(ByVal sToken As String) As Boolean


        Dim T = From TK In db.Sesion Where TK.cToken = sToken And DbFunctions.DiffHours(TK.dFecVenc, Now()) < 0 Select TK.cCodUsu, TK.dFecVenc

        If T.Count() > 0 Then
            ValidaToken = True
        Else
            ValidaToken = False
        End If

    End Function
    Public Function ValidaLogin() As SesionResp
        Dim sesion As New Sesion
        Dim sesionResp As New SesionResp

        Try


            Dim U = From S In db.Usuarios Where S.cCodUsu = cCodUsu Select S.cNombres, S.cCodUsu, S.sContrasenia, S.bValidaImei, S.cIMEI, S.bCambioClave, S.nCodAge
            Dim R = From S In db.UsuarioRoles Where S.cCodUsu = cCodUsu Select S.cCodUsu, S.cCodRol Order By cCodRol Descending



            'Si no ecnuentra a Usuario
            If U.Count() = 0 Then
                sesionResp.dFecVenc = Now()
                sesionResp.sToken = ""
                sesionResp.sRespuesta = "Usuario No Existe"
                sesionResp.nResultado = 0
                sesionResp.sCodRol = ""
                Return sesionResp
            End If

            'Valida si es una aplicacion de la empresa
            If cTipoApl = "1" And cClaveApp <> "nub32019azure2019" Then
                sesionResp.dFecVenc = Now()
                sesionResp.sToken = ""
                sesionResp.sRespuesta = "Aplicacion no Valida"
                sesionResp.nResultado = 0
                sesionResp.sCodRol = ""
                Return sesionResp
            End If

            If cTipoApl = "2" And cClaveApp <> "mov1l2019@seguraperu" Then
                sesionResp.dFecVenc = Now()
                sesionResp.sToken = ""
                sesionResp.sRespuesta = "Aplicacion no Valida"
                sesionResp.nResultado = 0
                sesionResp.sCodRol = ""
                Return sesionResp
            End If

            'Valida IMEI
            If cTipoApl = "2" And (U.First.bValidaImei = 1 Or U.First.bValidaImei = True) Then
                If cImei <> U.First.cIMEI Then
                    sesionResp.dFecVenc = Now()
                    sesionResp.sToken = ""
                    sesionResp.sRespuesta = "IMEI Ivalido"
                    sesionResp.nResultado = 0
                    sesionResp.sCodRol = ""
                    Return sesionResp
                End If
            End If

            'Valida Si clave es correcta
            Dim sClaveBD As String
            sClaveBD = U.First.sContrasenia.Trim
            If sClaveBD = cClave Then

                sesion.cCodUsu = U.First.cCodUsu
                sesion.cToken = CrearPassword(64)
                sesion.dFecReg = Now()
                sesion.dFecVenc = DateAdd(DateInterval.Hour, 12, Now())
                db.Sesion.Add(sesion)
                db.SaveChanges()

                sesionResp.dFecVenc = sesion.dFecVenc
                sesionResp.sToken = sesion.cToken
                sesionResp.sRespuesta = U.First.cNombres
                sesionResp.nResultado = 1
                sesionResp.sCodRol = R.First.cCodRol
                If U.First.bCambioClave = True Or U.First.bCambioClave = 1 Then
                    sesionResp.nCambioClave = 1
                Else
                    sesionResp.nCambioClave = 0
                End If
                sesionResp.nCodAge = U.First.nCodAge

            Else
                sesionResp.dFecVenc = Now()
                sesionResp.sToken = ""
                sesionResp.sRespuesta = "Usuario Clave o Clave Invalida"
                sesionResp.nResultado = 0
                sesionResp.sCodRol = ""

            End If

            Return sesionResp
        Catch ex As Exception
            sesionResp.dFecVenc = DateAdd(DateInterval.Hour, 8, Now())
            sesionResp.sToken = ""
            sesionResp.sRespuesta = ex.Message
            sesionResp.nResultado = 1
            Return sesionResp
        End Try

    End Function

    Public Function CrearPassword(ByVal longitud As Integer) As String
        Dim caracteres As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
        Dim Res As StringBuilder = New StringBuilder()
        Dim Rnd As Random = New Random()
        Do While (0 < longitud)
            Res.Append(caracteres(Rnd.Next(caracteres.Length)))
            longitud = longitud - 1
        Loop


        Return Res.ToString()

    End Function


    Public Function CambiarClave() As SesionResp

        Dim S As SesionResp = New SesionResp()
        Dim U As Usuarios = New Usuarios()

        S = ValidaLogin()
        U.cCodUsu = cCodUsu

        'Valida si Credenciales Correctas procede a hacer cambio de clave
        If S.nResultado = 1 Then
            db.Configuration.AutoDetectChangesEnabled = False

            U = db.Usuarios.Find(U.cCodUsu)
            db.Usuarios.Attach(U)
            U.sContrasenia = cClaveNueva
            U.bCambioClave = 0
            db.Entry(U).Property("sContrasenia").IsModified = True
            db.Entry(U).Property("bCambioClave").IsModified = True
            db.SaveChanges()

            db.Configuration.AutoDetectChangesEnabled = True

            S.sRespuesta = "Datos Cambiados Correctamente"
            S.nResultado = 1
        End If

        Return S

    End Function

End Class
