<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateAddin
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdateAddin))
        Me.Pb1 = New System.Windows.Forms.ProgressBar()
        Me.LblAccion = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Pb1
        '
        Me.Pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Pb1.Location = New System.Drawing.Point(187, 305)
        Me.Pb1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Pb1.Name = "Pb1"
        Me.Pb1.Size = New System.Drawing.Size(526, 27)
        Me.Pb1.TabIndex = 5
        '
        'LblAccion
        '
        Me.LblAccion.Location = New System.Drawing.Point(183, 244)
        Me.LblAccion.Name = "LblAccion"
        Me.LblAccion.Size = New System.Drawing.Size(530, 45)
        Me.LblAccion.TabIndex = 6
        Me.LblAccion.Text = "Action :"
        '
        'frmUpdateAddin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.ULMAUpdaterAddIn.My.Resources.Resources.ActivationCode
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(908, 483)
        Me.ControlBox = False
        Me.Controls.Add(Me.LblAccion)
        Me.Controls.Add(Me.Pb1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Neo Sans Pro", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdateAddin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DOWNLOAD - UPDATE"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Pb1 As Windows.Forms.ProgressBar
    Friend WithEvents LblAccion As Windows.Forms.Label
End Class
