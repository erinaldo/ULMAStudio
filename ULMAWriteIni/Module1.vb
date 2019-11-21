Module Module1
    Public _appDLLFull As String = "C:\ULMA\INSTALL\ULMAStudio\ULMAStudio.dll"
    Public _appININame As String = "ULMAUpdaterAddin.ini"
    Public _appINIFull As String = IO.Path.Combine(IO.Path.GetDirectoryName(_appDLLFull), _appININame)
    Sub Main()
        ' Le pasamos como parámetro la carpeta inicial C:\ULMA
        If Environment.GetCommandLineArgs.Count < 2 Then Exit Sub
        If _appDLLFull <> Environment.GetCommandLineArgs(1) Then
            _appDLLFull = Environment.GetCommandLineArgs(1)
            _appINIFull = IO.Path.Combine(IO.Path.GetDirectoryName(_appDLLFull), _appININame)
        End If
        If IO.File.Exists(_appDLLFull) = False Then Exit Sub
        '
        WriteIni()
        'texto &= ";Last updates." & vbCrLf
        'texto &= "[LAST]" & vbCrLf
        'texto &= "ULMAStudio = ULMAStudio_0.0.0.23.zip" & vbCrLf
        'texto &= "XML = XML_PUBLIC_20191028.zip" & vbCrLf
        'texto &= ";Updates availables." & vbCrLf
        'texto &= "[UPDATES]"
    End Sub

    Public Sub WriteIni()
        Dim assembly As System.Reflection.Assembly
        assembly = System.Reflection.Assembly.LoadFile(_appDLLFull)
        '
        If assembly Is Nothing Then
            MsgBox("Error cargando DLL " & _appDLLFull)
            Exit Sub
        End If
        'MsgBox(assembly.Location)
        Dim oAinf As New Microsoft.VisualBasic.ApplicationServices.AssemblyInfo(assembly)
        'Console.WriteLine(oAinf.AssemblyName & "_v" & oAinf.Version.ToString)
        'Console.WriteLine(oAinf.AssemblyName)
        'Console.WriteLine(oAinf.Version.ToString)
        'Console.ReadLine()
        '
        Dim namesolo As String = IO.Path.GetFileNameWithoutExtension(_appDLLFull)
        Dim lineas() As String = IO.File.ReadAllLines(_appINIFull)
        Dim cambiado As Boolean = False
        Dim sufijo As String = oAinf.Version.ToString.Substring(4)
        For x As Integer = 0 To lineas.Count - 1
            Dim linea As String = lineas(x)
            If linea.StartsWith(namesolo) Then
                lineas(x) = namesolo & "=" & namesolo & "_" & oAinf.Version.ToString & ".zip"
                cambiado = True
                Exit For
            End If
        Next
        '
        If cambiado = True Then
            Dim textofinal As String = String.Join(vbCrLf, lineas)
            IO.File.WriteAllText(_appINIFull, textofinal, System.Text.Encoding.UTF8)
            'MsgBox(textofinal)
        End If
    End Sub
End Module
