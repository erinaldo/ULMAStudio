Public Class clsVersions
    Public Shared colVersions As New Dictionary(Of String, clsVersion)    ' Key=2018, Value=clsVersion

    Public Sub New(oAppC As Autodesk.Revit.ApplicationServices.ControlledApplication, oP As Process)
        If colVersions.ContainsKey(oAppC.VersionNumber) = False Then
            colVersions.Add(oAppC.VersionNumber, New clsVersion(oAppC, oP))
        End If
    End Sub

    Public Shared Function Version_Dame(Folder As String) As clsVersion
        Dim year As String = Folder
        Dim partes() As String = Folder.Split("\"c) ' por si viene al final "\2018"
        year = partes(UBound(partes))
        partes = year.Split(" ")                    ' Por si viene al final "Revit 2018"
        year = partes(UBound(partes))
        '
        Dim resultado As clsVersion = Nothing
        If colVersions.ContainsKey(year) Then
            resultado = colVersions(year)
        End If
        Return resultado
    End Function
End Class

Public Class clsVersion
    Public RevitSubVersionNumber As String = "2018.3"
    Public RevitVersionBuild As String = "20190510_1515(X64)"
    Public RevitVersionName As String = "Autodesk Revit 2018"
    Public RevitVersionNumber As String = "2018"
    Public RevitVersionNumberR As String = "R2018"
    Public RevitVersionLogText As String = "Revit 2018.3 (20190510_1515(X64))"
    Public RevitFullPath As String = "G:\Program files\Autodesk\Revit 2018\revit.exe"

    Public Sub New(oAppC As Autodesk.Revit.ApplicationServices.ControlledApplication, oP As Process)
        RevitSubVersionNumber = oAppC.SubVersionNumber
        RevitVersionBuild = oAppC.VersionBuild
        RevitVersionName = oAppC.VersionName
        RevitVersionNumber = oAppC.VersionNumber
        RevitVersionNumberR = "R" & oAppC.VersionNumber
        RevitVersionLogText = "Revit " & oAppC.SubVersionNumber & " (" & oAppC.VersionBuild & ")"
        RevitFullPath = oP.MainModule.FileName
    End Sub
End Class
