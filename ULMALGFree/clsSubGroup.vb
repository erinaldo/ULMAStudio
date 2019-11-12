Public Class clsSubGroup
    Public active As Boolean = True
    Public subgroupCode As String = ""
    Public defaultDescription As String = ""
    Public DicDescritions As Dictionary(Of String, String)         '' Key=es, Value=descripción en el idioma de Key.
    Public visualizationOrder As String
    Public LstblockCodes As List(Of String)
    Public LstarticleCodes As List(Of String)
End Class