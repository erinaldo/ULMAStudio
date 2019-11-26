Public Class clsVersions
    Public Shared colVersions As New Dictionary(Of String, clsVersion)    ' Key=2018, Value=clsVersion

    Public Shared Function Version_Dame(oAppC As Autodesk.Revit.ApplicationServices.ControlledApplication, oP As Process) As clsVersion
        Dim resultado As clsVersion = Nothing
        If colVersions.ContainsKey(oAppC.VersionNumber) = False Then
            Try
                colVersions.Add(oAppC.VersionNumber, New clsVersion(oAppC, oP))
            Catch ex As Exception
                colVersions.Add(oAppC.VersionNumber, New clsVersion(oP))
            End Try
        End If
        resultado = colVersions(oAppC.VersionNumber)

        Return resultado
    End Function
    Public Shared Function Version_Dame(oP As Process) As clsVersion
        Dim resultado As clsVersion = New clsVersion(oP)
        If colVersions.ContainsKey(resultado.RevitVersionNumber) = False Then
            colVersions.Add(resultado.RevitVersionNumber, resultado)
        End If
        Return resultado
    End Function
End Class

Public Class clsVersion
    Public RevitSubVersionNumber As String = ""  ' "2018.3"
    Public RevitVersionBuild As String = ""  ' "20190510_1515(X64)"
    Public RevitVersionName As String = ""   ' "Autodesk Revit 2018"
    Public RevitVersionNumber As String = "" ' "2018"
    Public RevitVersionNumberR As String = ""    ' "R2018"
    Public RevitVersionLogText As String = ""    ' "Revit 2018.3 (20190510_1515(X64))"
    Public RevitFullPath As String = ""  ' "G:\Program files\Autodesk\Revit 2018\revit.exe"

    Public Sub New(oAppC As Autodesk.Revit.ApplicationServices.ControlledApplication, oP As Process)
        If RevitSubVersionNumber = "" Then RevitSubVersionNumber = oAppC.SubVersionNumber
        If RevitVersionBuild = "" Then RevitVersionBuild = oAppC.VersionBuild
        If RevitVersionName = "" Then RevitVersionName = oAppC.VersionName
        If RevitVersionNumber = "" Then RevitVersionNumber = oAppC.VersionNumber
        If RevitVersionNumberR = "" Then RevitVersionNumberR = "R" & oAppC.VersionNumber
        If RevitVersionLogText = "" Then RevitVersionLogText = "Revit " & oAppC.SubVersionNumber & " (" & oAppC.VersionBuild & ")"
        If RevitFullPath = "" Then RevitFullPath = oP.MainModule.FileName
    End Sub
    Public Sub New(oP As Process)
        Dim oI As FileVersionInfo = oP.MainModule.FileVersionInfo
        Dim FMa As Integer = oI.FileMajorPart
        Dim FMi As Integer = oI.FileMinorPart
        Dim FBu As Integer = oI.FileBuildPart
        Dim FPr As Integer = oI.FilePrivatePart

        If RevitSubVersionNumber = "" Then RevitSubVersionNumber = "20" & oI.FileVersion
        If RevitVersionNumber = "" Then RevitVersionNumber = "20" & FMa.ToString
        If RevitVersionNumberR = "" Then RevitVersionNumberR = "R" & RevitVersionNumber
        If RevitVersionBuild = "" Then RevitVersionBuild = oI.ProductVersion
        If RevitVersionName = "" Then RevitVersionName = oI.FileDescription & " " & RevitVersionNumber
        If RevitVersionLogText = "" Then RevitVersionLogText = "Revit " & RevitSubVersionNumber & " (" & RevitVersionBuild & ")"
        If RevitFullPath = "" Then RevitFullPath = oP.MainModule.FileName
    End Sub
End Class
