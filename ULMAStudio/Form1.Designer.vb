<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdater
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
        Me.fpSuperGroups = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnRellena = New System.Windows.Forms.Button()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.pb1 = New System.Windows.Forms.PictureBox()
        Me.fpGroups = New System.Windows.Forms.FlowLayoutPanel()
        Me.PanelSep1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lblInf = New System.Windows.Forms.Label()
        CType(Me.pb1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fpSuperGroups
        '
        Me.fpSuperGroups.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.fpSuperGroups.Location = New System.Drawing.Point(22, 64)
        Me.fpSuperGroups.Name = "fpSuperGroups"
        Me.fpSuperGroups.Size = New System.Drawing.Size(100, 425)
        Me.fpSuperGroups.TabIndex = 0
        '
        'btnRellena
        '
        Me.btnRellena.Location = New System.Drawing.Point(12, 12)
        Me.btnRellena.Name = "btnRellena"
        Me.btnRellena.Size = New System.Drawing.Size(194, 41)
        Me.btnRellena.TabIndex = 1
        Me.btnRellena.Text = "Fill SuperGroup"
        Me.btnRellena.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCerrar.Location = New System.Drawing.Point(662, 12)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(126, 32)
        Me.btnCerrar.TabIndex = 2
        Me.btnCerrar.Text = "Close"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'pb1
        '
        Me.pb1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.Location = New System.Drawing.Point(290, 64)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(498, 425)
        Me.pb1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pb1.TabIndex = 3
        Me.pb1.TabStop = False
        '
        'fpGroups
        '
        Me.fpGroups.AutoScroll = True
        Me.fpGroups.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.fpGroups.Location = New System.Drawing.Point(138, 64)
        Me.fpGroups.Margin = New System.Windows.Forms.Padding(0)
        Me.fpGroups.Name = "fpGroups"
        Me.fpGroups.Size = New System.Drawing.Size(143, 425)
        Me.fpGroups.TabIndex = 4
        Me.fpGroups.Visible = False
        Me.fpGroups.WrapContents = False
        '
        'PanelSep1
        '
        Me.PanelSep1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PanelSep1.AutoScroll = True
        Me.PanelSep1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.PanelSep1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.PanelSep1.Location = New System.Drawing.Point(127, 64)
        Me.PanelSep1.Margin = New System.Windows.Forms.Padding(0)
        Me.PanelSep1.Name = "PanelSep1"
        Me.PanelSep1.Size = New System.Drawing.Size(5, 425)
        Me.PanelSep1.TabIndex = 5
        Me.PanelSep1.Visible = False
        Me.PanelSep1.WrapContents = False
        '
        'lblInf
        '
        Me.lblInf.AutoSize = True
        Me.lblInf.Location = New System.Drawing.Point(261, 20)
        Me.lblInf.Name = "lblInf"
        Me.lblInf.Size = New System.Drawing.Size(240, 17)
        Me.lblInf.TabIndex = 6
        Me.lblInf.Text = "* Right button on buttons, for options"
        Me.lblInf.Visible = False
        '
        'frmUpdater
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnCerrar
        Me.ClientSize = New System.Drawing.Size(800, 515)
        Me.Controls.Add(Me.lblInf)
        Me.Controls.Add(Me.PanelSep1)
        Me.Controls.Add(Me.fpGroups)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.btnRellena)
        Me.Controls.Add(Me.fpSuperGroups)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdater"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        CType(Me.pb1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents fpSuperGroups As FlowLayoutPanel
    Friend WithEvents btnRellena As Button
    Friend WithEvents btnCerrar As Button
    Friend WithEvents pb1 As PictureBox
    Friend WithEvents fpGroups As FlowLayoutPanel
    Friend WithEvents PanelSep1 As FlowLayoutPanel
    Friend WithEvents lblInf As Label
End Class
