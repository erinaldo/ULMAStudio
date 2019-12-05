<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportOptions
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RbView = New System.Windows.Forms.RadioButton()
        Me.RbProject = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RbAll = New System.Windows.Forms.RadioButton()
        Me.RbUlma = New System.Windows.Forms.RadioButton()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(171, 122)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "Aceptar"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancelar"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RbView)
        Me.GroupBox1.Controls.Add(Me.RbProject)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 10)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(116, 93)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Visible families of"
        '
        'RbView
        '
        Me.RbView.AutoSize = True
        Me.RbView.Checked = True
        Me.RbView.Location = New System.Drawing.Point(11, 26)
        Me.RbView.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RbView.Name = "RbView"
        Me.RbView.Size = New System.Drawing.Size(81, 17)
        Me.RbView.TabIndex = 1
        Me.RbView.TabStop = True
        Me.RbView.Text = "Active View"
        Me.RbView.UseVisualStyleBackColor = True
        '
        'RbProject
        '
        Me.RbProject.AutoSize = True
        Me.RbProject.Location = New System.Drawing.Point(11, 58)
        Me.RbProject.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RbProject.Name = "RbProject"
        Me.RbProject.Size = New System.Drawing.Size(72, 17)
        Me.RbProject.TabIndex = 0
        Me.RbProject.Text = "All Project"
        Me.RbProject.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RbAll)
        Me.GroupBox2.Controls.Add(Me.RbUlma)
        Me.GroupBox2.Location = New System.Drawing.Point(175, 10)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(142, 93)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Families"
        '
        'RbAll
        '
        Me.RbAll.AutoSize = True
        Me.RbAll.Location = New System.Drawing.Point(18, 58)
        Me.RbAll.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RbAll.Name = "RbAll"
        Me.RbAll.Size = New System.Drawing.Size(110, 17)
        Me.RbAll.TabIndex = 1
        Me.RbAll.TabStop = True
        Me.RbAll.Text = "All custom families"
        Me.RbAll.UseVisualStyleBackColor = True
        '
        'RbUlma
        '
        Me.RbUlma.AutoSize = True
        Me.RbUlma.Checked = True
        Me.RbUlma.Location = New System.Drawing.Point(18, 26)
        Me.RbUlma.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.RbUlma.Name = "RbUlma"
        Me.RbUlma.Size = New System.Drawing.Size(92, 17)
        Me.RbUlma.TabIndex = 0
        Me.RbUlma.TabStop = True
        Me.RbUlma.Text = "ULMA families"
        Me.RbUlma.UseVisualStyleBackColor = True
        '
        'frmReportOptions
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(329, 163)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReportOptions"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Report Options"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents RbView As Windows.Forms.RadioButton
    Friend WithEvents RbProject As Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents RbAll As Windows.Forms.RadioButton
    Friend WithEvents RbUlma As Windows.Forms.RadioButton
End Class
