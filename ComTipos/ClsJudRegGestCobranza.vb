Option Explicit On
<Serializable()> Public Class ClsJudRegGestCobranza
    Public oagencias As clsDataset
    Public ogestor As clsDataset
    Public orecuperacartera As clsDataset
    ''
  


    Public Sub New()
        oagencias = New clsDataset
        ogestor = New clsDataset
        orecuperacartera = New clsDataset

   
    End Sub

End Class
