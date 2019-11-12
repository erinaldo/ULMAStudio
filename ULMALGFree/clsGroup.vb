Imports uf = ULMALGFree.clsBase
Public Class clsGroup
    Public active As Boolean = True
    Public groupCode As String = ""
    Public DefaultDescription As String = ""
    Public DicDescritions As Dictionary(Of String, String)         '' Key=es, Value=descripción en el idioma de Key.
    Public productType As String = ""
    Public shortName As String = ""
    Public LstFilenameOnly As List(Of String)
    Public LstarticleCodes As List(Of String)
    ''
    'Public Sub fiPon(FullFile As String)
    '    If LstFilenameOnly Is Nothing Then LstFilenameOnly = New List(Of String)
    '    If LstarticleCodes Is Nothing Then LstarticleCodes = New List(Of String)
    '    '
    '    Dim nombresSin As String = IO.Path.GetFileNameWithoutExtension(FullFile)
    '    Dim partes() As String = nombresSin.Split("_"c)
    '    Dim article As String = partes(UBound(partes))

    '    If LstFilenameOnly.Contains(nombresSin) = False Then
    '        LstFilenameOnly.Add(nombresSin)
    '    End If
    '    If LstarticleCodes.Contains(article) = False Then
    '        LstarticleCodes.Add(article)
    '    End If
    'End Sub

    Public Sub RellenaFilenameOnly()
        If LstFilenameOnly Is Nothing Then LstFilenameOnly = New List(Of String)
        '
        For Each fiRFA In IO.Directory.GetFiles(uf.path_families_base, "*.rfa", IO.SearchOption.TopDirectoryOnly).AsParallel
            Dim nombresin As String = IO.Path.GetFileNameWithoutExtension(fiRFA)
            Dim partes() As String = nombresin.Split("_"c)
            Dim article As String = partes(UBound(partes))
            If Me.LstarticleCodes.Contains(article) Then
                LstFilenameOnly.Add(IO.Path.GetFileNameWithoutExtension(fiRFA))
            End If
        Next
    End Sub
End Class
