Module Module1
    Public years() As String = {"2018", "2019"}
    Public texto As String = ""
    Public _appPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    Public _appFolder As String = IO.Path.GetDirectoryName(_appPath)
    Public _appDLL As String = "G:\ULMA\INSTALL\ULMAStudio\ULMAStudio.dll"
    Public _appINIPre As String = "ULMAUpdaterAddin"   'G:\ULMA\ULMAUpdaterAddin2018.ini
    Public _appINIFolder As String = "G:\ULMA"
    Sub Main()
        ' Le pasamos como parámetro la carpeta inicial G:\ULMA
        If Environment.GetCommandLineArgs.Count < 2 Then Exit Sub
        _appINIFolder = Environment.GetCommandLineArgs(1)
        If IO.Directory.Exists(_appINIFolder) = False Then Exit Sub
        '
        'MsgBox(_appINIFolder)
        texto &= ";© Jose Alberto Torres (2aCAD Graitec Group)" & vbCrLf
        texto &= ";Last updates." & vbCrLf
        texto &= "[LAST]" & vbCrLf
        texto &= "ULMAStudio = ULMAStudio_#YEAR#.zip" & vbCrLf
        texto &= "XML = XML_PUBLIC_20191028.zip" & vbCrLf
        texto &= ";Updates availables." & vbCrLf
        texto &= "[UPDATES]"
        _appDLL = IO.Path.Combine(_appINIFolder, "ULMAStudio\ULMAStudio.dll")
        'MsgBox(_appDLL)
        If IO.File.Exists(_appDLL) = False Then Exit Sub
        WriteIni()
    End Sub

    Public Sub WriteIni()
        Dim assembly As System.Reflection.Assembly
        assembly = System.Reflection.Assembly.LoadFile(_appDLL)
        '
        If assembly Is Nothing Then
            MsgBox("Error cargando DLL " & _appDLL)
            Exit Sub
        End If
        'MsgBox(assembly.Location)
        Dim oAinf As New Microsoft.VisualBasic.ApplicationServices.AssemblyInfo(assembly)
        'Console.WriteLine(oAinf.AssemblyName & "_v" & oAinf.Version.ToString)
        'Console.WriteLine(oAinf.AssemblyName)
        'Console.WriteLine(oAinf.Version.ToString)
        Dim sufijo As String = oAinf.Version.ToString.Substring(4)
        For Each year As String In years
            Dim t As String = texto.Replace("#YEAR#", year & sufijo)
            Dim fullfi As String = IO.Path.GetFullPath(IO.Path.Combine(_appINIFolder & "..\", _appINIPre & year & ".ini"))
            'MsgBox(fullfi)
            IO.File.WriteAllText(fullfi, t, System.Text.Encoding.UTF8)
        Next
        'Console.ReadLine()
    End Sub
End Module
