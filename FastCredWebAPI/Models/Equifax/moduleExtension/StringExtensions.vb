Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function NullToEmpty(value As String) As String
        If value Is Nothing Then
            Return String.Empty
        End If
        Return value
    End Function
    <Extension()>
    Public Function NullToZero(value As Integer?) As Integer
        Return If(value.HasValue, value.Value, 0)
    End Function
End Module
