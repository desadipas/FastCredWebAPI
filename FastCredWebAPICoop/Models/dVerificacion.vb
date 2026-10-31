Imports FastCredWebAPI.FastCredWebAPI

Public Class dVerificacion

    Private db As New DBNEGOCIOEntities

    Public Property cDoc As String

    Public Property nCodCred As Integer

    Public Property nLatDom As Double

    Public Property nLonDom As Double

    Public Property cReferencia1 As String

    Public Property cReferencia2 As String

    Public Property cReferencia3 As String

    Public Property cComentario As String

    Public Property dFechaReg As Date
    Public Property cUser As String

    Public Property oImagenes As dVerificacionimagenes()


    Public Function RegistraVerificacion() As String()
        Dim D(2) As String
        Dim i As Integer

        Dim oVerifi As New VerificacionDomic

        oVerifi.nCodCred = nCodCred
        oVerifi.cRefer1 = cReferencia1
        oVerifi.cRefer2 = cReferencia2
        oVerifi.cRefer3 = cReferencia3
        oVerifi.cUser = cUser
        oVerifi.dFechaVerif = dFechaReg

        db.VerificacionDomic.Add(oVerifi)
        Try
            db.SaveChanges()

            D(0) = "1"
            D(1) = "Registro Correcto"

        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
            Return D
        End Try


        Dim oVerifiImg As VerificacionDomicImagenes
        For i = 0 To oImagenes.Count - 1
            oVerifiImg = New VerificacionDomicImagenes
            oVerifiImg.nCodCred = nCodCred
            oVerifiImg.nImagen = oImagenes(i).nImagen

            db.VerificacionDomicImagenes.Add(oVerifiImg)
            Try
                db.SaveChanges()

                D(0) = "1"
                D(1) = "Registro Correcto"

            Catch ex As Exception
                D(0) = "0"
                D(1) = ex.Message
                Return D
            End Try
        Next

        Dim P As New Persona

        Dim vP = From U In db.Persona
                 Where U.cDoc = cDoc
                 Select U

        db.Configuration.AutoDetectChangesEnabled = False

        P = db.Persona.First
        db.Persona.Attach(P)
        P.nLatitud = nLatDom
        P.nLongitud = nLonDom
        P.dFechaVerif = Now.Date

        db.Entry(P).Property("nLatitud").IsModified = True
        db.Entry(P).Property("nLongitud").IsModified = True
        db.Entry(P).Property("dFechaVerif").IsModified = True

        Try

            db.SaveChanges()
            D(0) = "1"
            D(1) = "Registro Correcto"


        Catch ex As Exception
            D(0) = "0"
            D(1) = ex.Message
            Return D
        End Try

        db.Configuration.AutoDetectChangesEnabled = True

        Return D

    End Function

End Class
