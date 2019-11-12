Module modUtiles
    'Public Sub Process_Close(name As String)   'Optional queProceso As String = "EXCEL"
    '    Try
    '        Dim proc As System.Diagnostics.Process
    '        ''
    '        For Each proc In System.Diagnostics.Process.GetProcessesByName(name)
    '            Try
    '                proc.Kill()
    '            Catch ex As Exception
    '                ' No hacemos nada
    '            End Try
    '        Next
    '    Catch ex As Exception
    '        '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
    '    End Try
    'End Sub
    'Private Async Sub Process_Run(quePath As String, Optional argumentos As String = Nothing)
    '    Dim p As Process = IIf(argumentos Is Nothing, Process.Start(quePath), Process.Start(quePath, argumentos))
    '    Await Task.Run(Sub() p.WaitForExit(15000))
    'End Sub

End Module
