Public Class frmGeneratingReport
    Public Property titulo As String = "** Generating Report **"

    Private Sub FrmGeneratingReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblTitulo.Text = titulo
        System.Windows.Forms.Application.EnableVisualStyles()
        pb1.Style = Windows.Forms.ProgressBarStyle.Marquee
        pb1.MarqueeAnimationSpeed = 50
    End Sub

    Public Sub pb1_Pon()
        'If pb1.Value = pb1.Maximum Then
        '    pb1.Value = 1
        'End If
        'If pb1.Value < pb1.Maximum Then
        '    pb1.Value += 1
        '    pb1.Refresh()
        'End If
    End Sub
End Class