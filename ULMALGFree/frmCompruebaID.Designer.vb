<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCompruebaID
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
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtActivacionCode = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.LinkLabel()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.OK_Button.Location = New System.Drawing.Point(679, 230)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(96, 28)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "Activate"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.BackColor = System.Drawing.Color.Transparent
        Me.Cancel_Button.BackgroundImage = Global.ULMALGFree.My.Resources.Resources.Cross
        Me.Cancel_Button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.FlatAppearance.BorderSize = 0
        Me.Cancel_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Cancel_Button.Location = New System.Drawing.Point(959, -1)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(57, 28)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label1.Font = New System.Drawing.Font("Calibri", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(195, 311)
        Me.Label1.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(382, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "If you have any issues entering the activation code, please contact:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(225, 235)
        Me.Label2.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(182, 19)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Enter your activation code:"
        '
        'TxtActivacionCode
        '
        Me.TxtActivacionCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtActivacionCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtActivacionCode.Location = New System.Drawing.Point(439, 231)
        Me.TxtActivacionCode.Margin = New System.Windows.Forms.Padding(5, 5, 5, 5)
        Me.TxtActivacionCode.Name = "TxtActivacionCode"
        Me.TxtActivacionCode.Size = New System.Drawing.Size(231, 25)
        Me.TxtActivacionCode.TabIndex = 4
        Me.TxtActivacionCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmail.Location = New System.Drawing.Point(617, 310)
        Me.lblEmail.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(181, 17)
        Me.lblEmail.TabIndex = 5
        Me.lblEmail.TabStop = True
        Me.lblEmail.Text = "bim@ulmaconstruction.com"
        '
        'frmCompruebaID
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackgroundImage = Global.ULMALGFree.My.Resources.Resources.ActivationCode
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(1013, 452)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.TxtActivacionCode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(1647, 862)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(667, 246)
        Me.Name = "frmCompruebaID"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmCompruebaID"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents TxtActivacionCode As Windows.Forms.TextBox
    Friend WithEvents lblEmail As Windows.Forms.LinkLabel
End Class
