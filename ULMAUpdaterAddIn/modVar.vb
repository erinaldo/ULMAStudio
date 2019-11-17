Module modVar
    ' ***** CLASES ******************
    'Public cIni As clsINI
    'Public cLcsv As ULMALGFree.clsBase
    '
    Public Function INICargar() As String()
        If cIni Is Nothing Then cIni = New clsINI
        Dim mensaje(1) As String
        ' Mensaje(0) contendrá los errores.
        ' Mensaje(1) contendrá los valores leidos del .INI
        mensaje(0) = "" : mensaje(1) = ""
        '
        Return mensaje
    End Function
End Module
