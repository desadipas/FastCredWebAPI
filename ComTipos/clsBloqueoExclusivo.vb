<Serializable()> Public Class clsBloqueoExclusivo
    Public Enum BloqueoExclusivoConsulta
        Consulta = 3
        ConsultaUsuario = 4
    End Enum

    Public Enum BloqueoExclusivoModifica
        Consulta = 1
        ConsultaUsuario = 2
        Actualiza = 5
    End Enum
End Class
