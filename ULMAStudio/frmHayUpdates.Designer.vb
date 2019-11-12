<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHayUpdates
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
        Me.PbAddins = New System.Windows.Forms.PictureBox()
        Me.PbFamiles = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnYes = New System.Windows.Forms.Button()
        Me.BtnNo = New System.Windows.Forms.Button()
        Me.LblAddin = New System.Windows.Forms.Label()
        Me.LblFamilies = New System.Windows.Forms.Label()
        Me.LblFam = New System.Windows.Forms.Label()
        Me.LblAdd = New System.Windows.Forms.Label()
        CType(Me.PbAddins, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PbFamiles, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PbAddins
        '
        Me.PbAddins.Image = Global.ULMAStudio.My.Resources.Resources.mnuAboutUU32
        Me.PbAddins.Location = New System.Drawing.Point(73, 271)
        Me.PbAddins.Name = "PbAddins"
        Me.PbAddins.Size = New System.Drawing.Size(48, 48)
        Me.PbAddins.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PbAddins.TabIndex = 1
        Me.PbAddins.TabStop = False
        '
        'PbFamiles
        '
        Me.PbFamiles.Image = Global.ULMAStudio.My.Resources.Resources.uf32_0
        Me.PbFamiles.Location = New System.Drawing.Point(73, 85)
        Me.PbFamiles.Name = "PbFamiles"
        Me.PbFamiles.Size = New System.Drawing.Size(48, 48)
        Me.PbFamiles.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PbFamiles.TabIndex = 0
        Me.PbFamiles.TabStop = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(31, 421)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(429, 32)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "¿Close REVIT and update?"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(23, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(244, 20)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "There are updates to download:"
        '
        'BtnYes
        '
        Me.BtnYes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnYes.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnYes.Location = New System.Drawing.Point(52, 474)
        Me.BtnYes.Name = "BtnYes"
        Me.BtnYes.Size = New System.Drawing.Size(75, 37)
        Me.BtnYes.TabIndex = 4
        Me.BtnYes.Text = "&Yes"
        Me.BtnYes.UseVisualStyleBackColor = True
        '
        'BtnNo
        '
        Me.BtnNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnNo.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNo.Location = New System.Drawing.Point(364, 474)
        Me.BtnNo.Name = "BtnNo"
        Me.BtnNo.Size = New System.Drawing.Size(75, 37)
        Me.BtnNo.TabIndex = 5
        Me.BtnNo.Text = "&No"
        Me.BtnNo.UseVisualStyleBackColor = True
        '
        'LblAddin
        '
        Me.LblAddin.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAddin.Location = New System.Drawing.Point(139, 271)
        Me.LblAddin.Name = "LblAddin"
        Me.LblAddin.Size = New System.Drawing.Size(331, 99)
        Me.LblAddin.TabIndex = 6
        Me.LblAddin.Text = "Updates AddIn/XML"
        '
        'LblFamilies
        '
        Me.LblFamilies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblFamilies.Location = New System.Drawing.Point(142, 85)
        Me.LblFamilies.Name = "LblFamilies"
        Me.LblFamilies.Size = New System.Drawing.Size(328, 100)
        Me.LblFamilies.TabIndex = 7
        Me.LblFamilies.Text = "Updates Families"
        '
        'LblFam
        '
        Me.LblFam.AutoSize = True
        Me.LblFam.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblFam.Location = New System.Drawing.Point(142, 63)
        Me.LblFam.Name = "LblFam"
        Me.LblFam.Size = New System.Drawing.Size(72, 17)
        Me.LblFam.TabIndex = 8
        Me.LblFam.Text = "Families:"
        '
        'LblAdd
        '
        Me.LblAdd.AutoSize = True
        Me.LblAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAdd.Location = New System.Drawing.Point(142, 250)
        Me.LblAdd.Name = "LblAdd"
        Me.LblAdd.Size = New System.Drawing.Size(98, 17)
        Me.LblAdd.TabIndex = 9
        Me.LblAdd.Text = "AddIns/XML:"
        '
        'frmHayUpdates
        '
        Me.AcceptButton = Me.BtnYes
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.BtnNo
        Me.ClientSize = New System.Drawing.Size(490, 541)
        Me.ControlBox = False
        Me.Controls.Add(Me.LblAdd)
        Me.Controls.Add(Me.LblFam)
        Me.Controls.Add(Me.LblFamilies)
        Me.Controls.Add(Me.LblAddin)
        Me.Controls.Add(Me.BtnNo)
        Me.Controls.Add(Me.BtnYes)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PbAddins)
        Me.Controls.Add(Me.PbFamiles)
        Me.MaximumSize = New System.Drawing.Size(508, 588)
        Me.MinimumSize = New System.Drawing.Size(508, 388)
        Me.Name = "frmHayUpdates"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CONTENT AVAILABLE"
        CType(Me.PbAddins, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PbFamiles, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PbFamiles As Windows.Forms.PictureBox
    Friend WithEvents PbAddins As Windows.Forms.PictureBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents BtnYes As Windows.Forms.Button
    Friend WithEvents BtnNo As Windows.Forms.Button
    Friend WithEvents LblAddin As Windows.Forms.Label
    Friend WithEvents LblFamilies As Windows.Forms.Label
    Friend WithEvents LblFam As Windows.Forms.Label
    Friend WithEvents LblAdd As Windows.Forms.Label
End Class
