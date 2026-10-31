Imports System
Imports System.Collections.Generic
Imports FastCredWebAPI.FastCredWebAPI
Public Class PersonaRelac

    Public Property nCodRela As Integer
    Public Property oPersonaRela As Persona

    Public Sub New()
        nCodRela = 0
        oPersonaRela = New Persona()
    End Sub
End Class
