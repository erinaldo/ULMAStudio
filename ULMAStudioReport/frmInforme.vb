Imports Microsoft.Reporting.WinForms
Public Class frmInforme
    'Public informe As ReportViewer
    Public filas As New List(Of fila)
    Public filasNoUlma As New List(Of fila)
    Public datos As DataTable
    Public datosNoUlma As DataTable
    Public view As String
    Public unidadPeso As String

    Private Sub FrmInforme_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.filas = New List(Of fila)
        Me.filasNoUlma = New List(Of fila)
        Me.datos = New DataTable("datos")
        Me.datosNoUlma = New DataTable("datosNoUlma")
        Application.EnableVisualStyles()
        RellenaDatosInforme()
        Me.Text = view
        Dim ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer
        ReportViewer1.Dock = DockStyle.Fill
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.ProcessingMode = ProcessingMode.Local
        ReportViewer1.LocalReport.EnableExternalImages = True
        ReportViewer1.SetDisplayMode(DisplayMode.PrintLayout)
        ' Crear y asignar los parámetros
        Dim parametros As New List(Of ReportParameter)
        parametros.Add(New ReportParameter("nView", view, True))
        parametros.Add(New ReportParameter("unPeso", unidadPeso, True))
        parametros.Add(New ReportParameter())
        '
        ' Asignar DataSource
        ReportViewer1.LocalReport.DataSources.Clear()
        ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("filas", filas))
        ReportViewer1.LocalReport.ReportPath = IO.Path.Combine(My.Application.Info.DirectoryPath, "UlmaReport1.rdlc")
        ReportViewer1.LocalReport.SetParameters(parametros)      '
        ReportViewer1.RefreshReport()
        Me.Controls.Add(ReportViewer1)
    End Sub

    Private Sub RellenaDatosInforme()
        Dim lineas() As String = IO.File.ReadAllLines(ULMALGFree.clsBase._ULMAStudioReport, System.Text.Encoding.UTF8)
        ' Borrar el fichero con los datos, una vez leido.
        Try
            IO.File.Delete(ULMALGFree.clsBase._ULMAStudioReport)
        Catch ex As Exception

        End Try
        '
        Me.view = lineas(0)
        Me.unidadPeso = lineas(1)

        For x As Integer = 2 To lineas.Count - 1
            Dim partes() As String = lineas(x).Split(";"c)
            Dim path As String = partes(0)
            Dim img As Image = Image.FromFile(path)
            Dim name As String = partes(1)
            Dim code As String = partes(2)
            Dim weight As String = partes(3)
            Dim quantity As String = partes(4)
            Dim esulma As Boolean = True
            Try
                esulma = CBool(partes(5))
            Catch ex As Exception
                esulma = False
            End Try
            If esulma = True Then
                Me.filas.Add(New fila(x, System.Drawing.Image.FromFile(path), name, code, weight, quantity, esulma))
            Else
                Me.filas.Add(New fila(x, System.Drawing.Image.FromFile(path), name, code, weight, quantity, esulma))
                Me.filasNoUlma.Add(New fila(x, System.Drawing.Image.FromFile(path), name, code, weight, quantity, esulma))
            End If
        Next
    End Sub

    Private Sub frmInforme_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        System.Windows.Forms.Application.Exit()
    End Sub
End Class