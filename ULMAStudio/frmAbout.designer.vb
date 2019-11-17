<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAbout
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.LblEmail = New System.Windows.Forms.Label()
        Me.BtnUpdateAddIn = New System.Windows.Forms.Button()
        Me.Pbox_New = New System.Windows.Forms.PictureBox()
        Me.PBox_Web = New System.Windows.Forms.PictureBox()
        Me.Pbox_Latest = New System.Windows.Forms.PictureBox()
        Me.LblVersion = New System.Windows.Forms.Label()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.pbActualiza = New System.Windows.Forms.ProgressBar()
        CType(Me.Pbox_New, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PBox_Web, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Pbox_Latest, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblEmail
        '
        Me.LblEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblEmail.BackColor = System.Drawing.Color.Transparent
        Me.LblEmail.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LblEmail.Font = New System.Drawing.Font("Calibri", 10.0!)
        Me.LblEmail.Location = New System.Drawing.Point(444, 356)
        Me.LblEmail.Name = "LblEmail"
        Me.LblEmail.Size = New System.Drawing.Size(267, 27)
        Me.LblEmail.TabIndex = 3
        Me.LblEmail.Text = " ·    bim@ulmaconstruction.com"
        Me.LblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'BtnUpdateAddIn
        '
        Me.BtnUpdateAddIn.BackColor = System.Drawing.Color.Transparent
        Me.BtnUpdateAddIn.BackgroundImage = Global.ULMAStudio.My.Resources.Resources.toupdate2
        Me.BtnUpdateAddIn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.BtnUpdateAddIn.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BtnUpdateAddIn.FlatAppearance.BorderSize = 0
        Me.BtnUpdateAddIn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnUpdateAddIn.Location = New System.Drawing.Point(413, 236)
        Me.BtnUpdateAddIn.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnUpdateAddIn.Name = "BtnUpdateAddIn"
        Me.BtnUpdateAddIn.Size = New System.Drawing.Size(57, 53)
        Me.BtnUpdateAddIn.TabIndex = 6
        Me.BtnUpdateAddIn.UseVisualStyleBackColor = False
        '
        'Pbox_New
        '
        Me.Pbox_New.BackColor = System.Drawing.Color.Transparent
        Me.Pbox_New.Image = CType(resources.GetObject("Pbox_New.Image"), System.Drawing.Image)
        Me.Pbox_New.Location = New System.Drawing.Point(341, 289)
        Me.Pbox_New.Margin = New System.Windows.Forms.Padding(4)
        Me.Pbox_New.Name = "Pbox_New"
        Me.Pbox_New.Size = New System.Drawing.Size(200, 28)
        Me.Pbox_New.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.Pbox_New.TabIndex = 7
        Me.Pbox_New.TabStop = False
        '
        'PBox_Web
        '
        Me.PBox_Web.BackColor = System.Drawing.Color.Transparent
        Me.PBox_Web.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PBox_Web.Image = CType(resources.GetObject("PBox_Web.Image"), System.Drawing.Image)
        Me.PBox_Web.Location = New System.Drawing.Point(176, 325)
        Me.PBox_Web.Margin = New System.Windows.Forms.Padding(4)
        Me.PBox_Web.Name = "PBox_Web"
        Me.PBox_Web.Size = New System.Drawing.Size(533, 25)
        Me.PBox_Web.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PBox_Web.TabIndex = 8
        Me.PBox_Web.TabStop = False
        '
        'Pbox_Latest
        '
        Me.Pbox_Latest.BackColor = System.Drawing.Color.Transparent
        Me.Pbox_Latest.Image = CType(resources.GetObject("Pbox_Latest.Image"), System.Drawing.Image)
        Me.Pbox_Latest.Location = New System.Drawing.Point(307, 246)
        Me.Pbox_Latest.Margin = New System.Windows.Forms.Padding(4)
        Me.Pbox_Latest.Name = "Pbox_Latest"
        Me.Pbox_Latest.Size = New System.Drawing.Size(272, 28)
        Me.Pbox_Latest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.Pbox_Latest.TabIndex = 9
        Me.Pbox_Latest.TabStop = False
        '
        'LblVersion
        '
        Me.LblVersion.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblVersion.BackColor = System.Drawing.Color.Transparent
        Me.LblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblVersion.Font = New System.Drawing.Font("Calibri", 10.0!)
        Me.LblVersion.Location = New System.Drawing.Point(176, 356)
        Me.LblVersion.Name = "LblVersion"
        Me.LblVersion.Size = New System.Drawing.Size(261, 27)
        Me.LblVersion.TabIndex = 10
        Me.LblVersion.Text = "ULMA Studio - v.2019.0.0.XX"
        Me.LblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.BackColor = System.Drawing.Color.Transparent
        Me.Cancel_Button.BackgroundImage = Global.ULMAStudio.My.Resources.Resources.Cross
        Me.Cancel_Button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.FlatAppearance.BorderSize = 0
        Me.Cancel_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Cancel_Button.Location = New System.Drawing.Point(689, 129)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(7, 6, 7, 6)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(57, 28)
        Me.Cancel_Button.TabIndex = 11
        Me.Cancel_Button.UseVisualStyleBackColor = False
        '
        'pbActualiza
        '
        Me.pbActualiza.Location = New System.Drawing.Point(307, 295)
        Me.pbActualiza.Name = "pbActualiza"
        Me.pbActualiza.Size = New System.Drawing.Size(272, 23)
        Me.pbActualiza.TabIndex = 12
        Me.pbActualiza.Visible = False
        '
        'frmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BackgroundImage = Global.ULMAStudio.My.Resources.Resources.About_ai_bim_Mesa_de_trabajo_1_copia_9
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(885, 516)
        Me.ControlBox = False
        Me.Controls.Add(Me.pbActualiza)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.LblVersion)
        Me.Controls.Add(Me.Pbox_Latest)
        Me.Controls.Add(Me.PBox_Web)
        Me.Controls.Add(Me.Pbox_New)
        Me.Controls.Add(Me.BtnUpdateAddIn)
        Me.Controls.Add(Me.LblEmail)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.Opacity = 0.25R
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About"
        CType(Me.Pbox_New, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PBox_Web, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Pbox_Latest, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LblEmail As Windows.Forms.Label
    Friend WithEvents BtnUpdateAddIn As Windows.Forms.Button
    Friend WithEvents Pbox_New As Windows.Forms.PictureBox
    Friend WithEvents PBox_Web As Windows.Forms.PictureBox
    Friend WithEvents Pbox_Latest As Windows.Forms.PictureBox
    Friend WithEvents LblVersion As Windows.Forms.Label
    Friend WithEvents Cancel_Button As Windows.Forms.Button
    Friend WithEvents pbActualiza As Windows.Forms.ProgressBar
End Class
