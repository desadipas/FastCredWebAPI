Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Public Class iProspecto

    'Public Property oProspecto As ProspectosCred
    Public Property oPersonaRela As PersonaRelac()
    Public Property oCreditos As Creditos
    Public Property oCredCalendCond As CredCalendCond
    Public Property oEvaluacion As Evaluacion
    Public Property oEvaluacionFotos As EvaluacionFotos()
    Public Property oArticuloTasacion As ArticuloTasacionDet()
    Public Sub New()
        'oProspecto = New ProspectosCred()
        oCreditos = New Creditos()
        oCredCalendCond = New CredCalendCond()
        oEvaluacion = New Evaluacion

        'ReDim oEvaluacionFotos(2)
        'oEvaluacionFotos(0) = New EvaluacionFotos()
        'oEvaluacionFotos(0).nCodEval = 0
        'oEvaluacionFotos(1) = New EvaluacionFotos()
        'oEvaluacionFotos(1).nCodEval = 0


        ReDim oArticuloTasacion(2)
        oArticuloTasacion(0) = New ArticuloTasacionDet()
        oArticuloTasacion(0).nCodTasacion = 0
        oArticuloTasacion(1) = New ArticuloTasacionDet()
        oArticuloTasacion(1).nCodTasacion = 0

    End Sub
    Public Function RegistraProspecto() As iProspectoResp
        Dim oPersonaRelaTemp As Persona
        Dim db As New DBNEGOCIOEntities
        Dim i As Integer
        Dim R As New iProspectoResp
        Dim T As DbContextTransaction
        Dim oCredPersonas As CredPersonas
        Dim nCodPersTitular As Long
        Dim vnCodMarca As Integer
        Dim vnCodModelo As Integer

        'Try

        '    T = db.Database.BeginTransaction()
        '**********************************************************************
        'Registra Prospecto
        '**********************************************************************
        '    Dim DP = From P In db.ProspectosCred Where P.nCodPers = oProspecto.nCodPers And P.nEstado <> 4 Select P.nID, P.nCodPers

        '    If DP.Count() > 0 Then
        '        'Actualiza el prospecto
        '        oProspecto.nID = DP.First.nID
        '        db.Entry(oProspecto).State = EntityState.Modified
        '        Try
        '            db.SaveChanges()
        '        Catch ex As DbUpdateConcurrencyException
        '            T.Rollback()
        '            R.nResultado = 0
        '            R.sRespuesta = ex.Message
        '            Return R
        '        End Try
        '    Else
        '        'Registra el prospecto
        '        db.ProspectosCred.Add(oProspecto)
        '        db.SaveChanges()
        '    End If

        '    '**********************************************************************
        '    'Registra personas que van a estar relacionadas al credito
        '    '**********************************************************************
        '    Dim vnRela As Integer = 0

        '    For i = 0 To UBound(oPersonaRela) - 1
        '        oPersonaRelaTemp = oPersonaRela(i).oPersonaRela
        '        Dim Rel = From P In db.Persona Where P.cDoc = oPersonaRelaTemp.cDoc Select P.nCodPers, P.cDoc

        '        If Rel.Count() > 0 Then
        '            'Actualiza la persona
        '            oPersonaRela(i).oPersonaRela.nCodPers = Rel.First.nCodPers
        '            db.Entry(oPersonaRela(i).oPersonaRela).State = EntityState.Modified
        '            Try
        '                db.SaveChanges()
        '            Catch ex As DbUpdateConcurrencyException
        '                T.Rollback()
        '                R.nResultado = 0
        '                R.sRespuesta = ex.Message
        '                Return R
        '            End Try
        '        Else
        '            'Registra la persona
        '            oPersonaRela(i).oPersonaRela = db.Persona.Add(oPersonaRela(i).oPersonaRela)
        '            db.SaveChanges()

        '        End If

        '        vnRela = oPersonaRela(i).nCodRela

        '        If vnRela = 10 Then
        '            nCodPersTitular = oPersonaRela(i).oPersonaRela.nCodPers
        '        End If

        '    Next



        '    '**********************************************************************
        '    'Registra Evaluacion de Creditos
        '    '**********************************************************************
        '    Dim vEvaluacion = From P In db.Evaluacion Where P.nCodEval = oEvaluacion.nCodEval And P.nCodPers = nCodPersTitular Select P.nCodPers, P.nCodEval

        '    If vEvaluacion.Count() > 0 Then

        '        db.Entry(oEvaluacion).State = EntityState.Modified
        '        Try
        '            db.SaveChanges()
        '        Catch ex As DbUpdateConcurrencyException
        '            T.Rollback()
        '            R.nResultado = 0
        '            R.sRespuesta = ex.Message
        '            Return R
        '        End Try
        '    Else
        '        oEvaluacion.nCodPers = nCodPersTitular
        '        oEvaluacion = db.Evaluacion.Add(oEvaluacion)
        '        db.SaveChanges()
        '    End If

        '    '**********************************************************************
        '    'Registra Creditos
        '    '**********************************************************************

        '    Dim vCred = From C In db.Creditos Where C.nCodAge = oCreditos.nCodAge And C.nCodCred = oCreditos.nCodCred Select C.nCodAge, C.nCodCred
        '    If vCred.Count() > 0 Then
        '        db.Entry(oCreditos).State = EntityState.Modified
        '        Try
        '            db.SaveChanges()
        '        Catch ex As DbUpdateConcurrencyException
        '            T.Rollback()
        '            R.nResultado = 0
        '            R.sRespuesta = ex.Message
        '            Return R
        '        End Try
        '    Else
        '        oCreditos.nCodEval = oEvaluacion.nCodEval
        '        oCreditos = db.Creditos.Add(oCreditos)
        '        db.SaveChanges()
        '    End If

        '    '**********************************************************************
        '    'Registra Credpersonas
        '    '**********************************************************************
        '    For i = 0 To UBound(oPersonaRela) - 1

        '        Dim vCredPers = From P In db.CredPersonas Where P.nCodAge = oCreditos.nCodAge And P.nCodCred = oCreditos.nCodCred And P.nRelacion = vnRela Select P.nCodAge, P.nCodCred, P.nRelacion

        '        If vCredPers.Count() > 0 Then
        '            oCredPersonas = New CredPersonas
        '            oCredPersonas.nCodCred = vCredPers.First.nCodCred
        '            oCredPersonas.nCodAge = vCredPers.First.nCodAge
        '            oCredPersonas.nCodPers = oPersonaRela(i).oPersonaRela.nCodPers
        '            oCredPersonas.nRelacion = vCredPers.First.nRelacion

        '            db.Entry(oCredPersonas).State = EntityState.Modified
        '            Try
        '                db.SaveChanges()
        '            Catch ex As DbUpdateConcurrencyException
        '                T.Rollback()
        '                R.nResultado = 0
        '                R.sRespuesta = ex.Message
        '                Return R
        '            End Try

        '        Else
        '            oCredPersonas = New CredPersonas
        '            oCredPersonas.nCodCred = oCreditos.nCodCred
        '            oCredPersonas.nCodAge = oCreditos.nCodAge
        '            oCredPersonas.nCodPers = oPersonaRela(i).oPersonaRela.nCodPers
        '            oCredPersonas.nRelacion = oPersonaRela(i).nCodRela
        '            db.CredPersonas.Add(oCredPersonas)
        '            db.SaveChanges()

        '        End If

        '    Next


        '    '**********************************************************************
        '    'Registra CredCalendCond
        '    '**********************************************************************
        '    Dim vCredCalendCond = From P In db.CredCalendCond Where P.nCodAge = oCreditos.nCodAge And P.nCodCred = oCreditos.nCodCred Select P.nCodAge, P.nCodCred

        '    oCredCalendCond.nCodAge = oCreditos.nCodAge
        '    oCredCalendCond.nCodCred = oCreditos.nCodCred
        '    oCredCalendCond.nNroCalen = 0

        '    If vCredCalendCond.Count() > 0 Then

        '        db.Entry(oCredCalendCond).State = EntityState.Modified
        '        Try
        '            db.SaveChanges()
        '        Catch ex As DbUpdateConcurrencyException
        '            T.Rollback()
        '            R.nResultado = 0
        '            R.sRespuesta = ex.Message
        '            Return R
        '        End Try
        '    Else
        '        oCredCalendCond = db.CredCalendCond.Add(oCredCalendCond)
        '        db.SaveChanges()
        '    End If

        '    '**********************************************************************
        '    'Registra oEvaluacionFotos
        '    '**********************************************************************
        '    For i = 0 To UBound(oEvaluacionFotos) - 1

        '        oEvaluacionFotos(i).nCodEval = oEvaluacion.nCodEval
        '        oEvaluacionFotos(i).nCodFoto = i

        '        Dim vFotos = From P In db.EvaluacionFotos Where P.nCodEval = oEvaluacion.nCodEval And P.nCodFoto = i Select P.nCodEval, P.nCodFoto

        '        If vFotos.Count() > 0 Then
        '            db.Entry(oEvaluacionFotos).State = EntityState.Modified
        '            Try
        '                db.SaveChanges()
        '            Catch ex As DbUpdateConcurrencyException
        '                T.Rollback()
        '                R.nResultado = 0
        '                R.sRespuesta = ex.Message
        '                Return R
        '            End Try

        '        Else
        '            oEvaluacionFotos(i) = db.EvaluacionFotos.Add(oEvaluacionFotos(i))
        '            db.SaveChanges()

        '        End If
        '    Next

        '    '**********************************************************************
        '    'Registra oArticuloTasacion Padre
        '    '**********************************************************************
        '    Dim vnCodTasacion As Integer
        '    Dim oArticuloTasacionMaestro As ArticuloTasacion
        '    Dim AC = From P In db.ArticuloTasacionCred Where P.nCodAge = oCreditos.nCodAge And P.nCodCred = oCreditos.nCodCred Select P.nCodTasacion, P.nCodAge, P.nCodCred
        '    Dim vSumaArticulos As Double
        '    Dim oArticuloTasacionCred As ArticuloTasacionCred

        '    vSumaArticulos = 0
        '    For i = 0 To UBound(oArticuloTasacion) - 1
        '        vSumaArticulos = vSumaArticulos + oArticuloTasacion(i).nPrestamoFinal
        '    Next

        '    If vSumaArticulos > 0 Then


        '        If AC.Count() > 0 Then
        '            vnCodTasacion = AC.First.nCodTasacion
        '        Else
        '            vnCodTasacion = 0
        '        End If

        '        If vnCodTasacion = 0 Then
        '            'Registra Nueva Tasacion
        '            oArticuloTasacionMaestro = New ArticuloTasacion()
        '            oArticuloTasacionMaestro.nCodTasacion = 0
        '            oArticuloTasacionMaestro.nPrestamoFinal = vSumaArticulos
        '            oArticuloTasacionMaestro.dFecha = Now()
        '            oArticuloTasacionMaestro = db.ArticuloTasacion.Add(oArticuloTasacionMaestro)
        '            db.SaveChanges()
        '            vnCodTasacion = oArticuloTasacionMaestro.nCodTasacion

        '            oArticuloTasacionCred = New ArticuloTasacionCred()
        '            oArticuloTasacionCred.nCodAge = oCreditos.nCodAge
        '            oArticuloTasacionCred.nCodCred = oCreditos.nCodCred
        '            oArticuloTasacionCred.nCodTasacion = vnCodTasacion

        '            db.ArticuloTasacionCred.Add(oArticuloTasacionCred)
        '            db.SaveChanges()

        '        Else
        '            'Actualiza Tasacion
        '            Dim ATas = From P In db.ArticuloTasacion Where P.nCodTasacion = vnCodTasacion Select P.nCodTasacion, P.nPrestamoFinal
        '            oArticuloTasacionMaestro = New ArticuloTasacion()
        '            oArticuloTasacionMaestro.dFecha = Now()
        '            oArticuloTasacionMaestro.nCodTasacion = vnCodTasacion
        '            oArticuloTasacionMaestro.nPrestamoFinal = vSumaArticulos

        '            db.Entry(oArticuloTasacionMaestro).State = EntityState.Modified
        '            Try
        '                db.SaveChanges()
        '            Catch ex As DbUpdateConcurrencyException
        '                T.Rollback()
        '                R.nResultado = 0
        '                R.sRespuesta = ex.Message
        '                Return R
        '            End Try
        '        End If


        '        '**********************************************************************
        '        'Registra oArticuloTasacion Hijo
        '        '**********************************************************************
        '        For i = 0 To UBound(oArticuloTasacion) - 1
        '            vnCodMarca = oArticuloTasacion(i).nCodMarca
        '            vnCodModelo = oArticuloTasacion(i).nCodModelo

        '            Dim ATasDet = From P In db.ArticuloTasacionDet Where P.nCodTasacion = vnCodTasacion And P.nCodMarca = vnCodMarca And P.nCodModelo = vnCodModelo Select P.nCodTasacion, P.nCodMarca, P.nCodModelo
        '            oArticuloTasacion(i).nCodTasacion = vnCodTasacion

        '            If ATasDet.Count() = 0 Then
        '                'Registrar Nuevo detalle de Tasacion
        '                oArticuloTasacion(i) = db.ArticuloTasacionDet.Add(oArticuloTasacion(i))
        '                db.SaveChanges()

        '            Else
        '                'Actualizar detalle de Tasacion
        '                db.Entry(oArticuloTasacion).State = EntityState.Modified
        '                Try
        '                    db.SaveChanges()
        '                Catch ex As DbUpdateConcurrencyException
        '                    T.Rollback()
        '                    R.nResultado = 0
        '                    R.sRespuesta = ex.Message
        '                    Return R
        '                End Try
        '            End If
        '        Next

        '    End If

        '    T.Commit()
        '    db.Database.Connection.Close()

        '    R.nResultado = 1
        '    R.sRespuesta = "Registro Correcto"
        '    Return R

        'Catch ex As Exception

        '    T.Rollback()
        '    db.Database.Connection.Close()
        '    R.nResultado = 0
        '    R.sRespuesta = ex.Message
        '    Return R

        'End Try


    End Function


End Class
