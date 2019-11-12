Module InicioMain
    Public fInf As frmInforme
    Public Sub Main(ByVal cmdArgs() As String)
        Application.EnableVisualStyles()
        If Process.GetProcessesByName(My.Application.Info.AssemblyName).Count > 1 Then Exit Sub
        fInf = New frmInforme
        Call fInf.ShowDialog()
    End Sub
End Module
