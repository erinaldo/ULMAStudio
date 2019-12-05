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
        Me.lbl_Terms = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pbox_NoNetwork = New System.Windows.Forms.PictureBox()
        CType(Me.Pbox_New, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PBox_Web, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Pbox_Latest, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbox_NoNetwork, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblEmail
        '
        Me.LblEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblEmail.BackColor = System.Drawing.Color.Transparent
        Me.LblEmail.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LblEmail.Font = New System.Drawing.Font("Calibri", 8.0!)
        Me.LblEmail.Location = New System.Drawing.Point(252, 295)
        Me.LblEmail.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblEmail.Name = "LblEmail"
        Me.LblEmail.Size = New System.Drawing.Size(217, 22)
        Me.LblEmail.TabIndex = 3
        Me.LblEmail.Text = " ·  bim@ulmaconstruction.com  · "
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
        Me.BtnUpdateAddIn.Location = New System.Drawing.Point(310, 192)
        Me.BtnUpdateAddIn.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnUpdateAddIn.Name = "BtnUpdateAddIn"
        Me.BtnUpdateAddIn.Size = New System.Drawing.Size(43, 43)
        Me.BtnUpdateAddIn.TabIndex = 6
        Me.BtnUpdateAddIn.UseVisualStyleBackColor = False
        '
        'Pbox_New
        '
        Me.Pbox_New.BackColor = System.Drawing.Color.Transparent
        Me.Pbox_New.Image = CType(resources.GetObject("Pbox_New.Image"), System.Drawing.Image)
        Me.Pbox_New.Location = New System.Drawing.Point(256, 235)
        Me.Pbox_New.Name = "Pbox_New"
        Me.Pbox_New.Size = New System.Drawing.Size(150, 23)
        Me.Pbox_New.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.Pbox_New.TabIndex = 7
        Me.Pbox_New.TabStop = False
        '
        'PBox_Web
        '
        Me.PBox_Web.BackColor = System.Drawing.Color.Transparent
        Me.PBox_Web.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PBox_Web.Image = CType(resources.GetObject("PBox_Web.Image"), System.Drawing.Image)
        Me.PBox_Web.Location = New System.Drawing.Point(132, 264)
        Me.PBox_Web.Name = "PBox_Web"
        Me.PBox_Web.Size = New System.Drawing.Size(400, 20)
        Me.PBox_Web.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PBox_Web.TabIndex = 8
        Me.PBox_Web.TabStop = False
        '
        'Pbox_Latest
        '
        Me.Pbox_Latest.BackColor = System.Drawing.Color.Transparent
        Me.Pbox_Latest.Image = CType(resources.GetObject("Pbox_Latest.Image"), System.Drawing.Image)
        Me.Pbox_Latest.Location = New System.Drawing.Point(230, 200)
        Me.Pbox_Latest.Name = "Pbox_Latest"
        Me.Pbox_Latest.Size = New System.Drawing.Size(204, 23)
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
        Me.LblVersion.Font = New System.Drawing.Font("Calibri", 8.0!)
        Me.LblVersion.Location = New System.Drawing.Point(81, 295)
        Me.LblVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblVersion.Name = "LblVersion"
        Me.LblVersion.Size = New System.Drawing.Size(172, 22)
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
        Me.Cancel_Button.Location = New System.Drawing.Point(517, 105)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(43, 23)
        Me.Cancel_Button.TabIndex = 11
        Me.Cancel_Button.UseVisualStyleBackColor = False
        '
        'pbActualiza
        '
        Me.pbActualiza.Location = New System.Drawing.Point(230, 240)
        Me.pbActualiza.Margin = New System.Windows.Forms.Padding(2)
        Me.pbActualiza.Name = "pbActualiza"
        Me.pbActualiza.Size = New System.Drawing.Size(204, 19)
        Me.pbActualiza.TabIndex = 12
        Me.pbActualiza.Visible = False
        '
        'lbl_Terms
        '
        Me.lbl_Terms.AutoSize = True
        Me.lbl_Terms.BackColor = System.Drawing.Color.White
        Me.lbl_Terms.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lbl_Terms.Font = New System.Drawing.Font("Calibri", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Terms.Location = New System.Drawing.Point(407, 300)
        Me.lbl_Terms.Name = "lbl_Terms"
        Me.lbl_Terms.Size = New System.Drawing.Size(75, 13)
        Me.lbl_Terms.TabIndex = 13
        Me.lbl_Terms.Text = "Terms of Use   ·"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label1.Font = New System.Drawing.Font("Calibri", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(483, 300)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(57, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "User guide"
        '
        'pbox_NoNetwork
        '
        Me.pbox_NoNetwork.BackColor = System.Drawing.Color.Transparent
        Me.pbox_NoNetwork.Image = Global.ULMAStudio.My.Resources.Resources.NoNetwork
        Me.pbox_NoNetwork.Location = New System.Drawing.Point(217, 197)
        Me.pbox_NoNetwork.Name = "pbox_NoNetwork"
        Me.pbox_NoNetwork.Size = New System.Drawing.Size(231, 39)
        Me.pbox_NoNetwork.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pbox_NoNetwork.TabIndex = 15
        Me.pbox_NoNetwork.TabStop = False
        '
        'frmAbout
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BackgroundImage = Global.ULMAStudio.My.Resources.Resources.About_ai_bim_Mesa_de_trabajo_1_copia_9
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ClientSize = New System.Drawing.Size(664, 419)
        Me.ControlBox = False
        Me.Controls.Add(Me.pbox_NoNetwork)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbl_Terms)
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
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(664, 419)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(664, 419)
        Me.Name = "frmAbout"
        Me.Opacity = 0.25R
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About"
        CType(Me.Pbox_New, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PBox_Web, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Pbox_Latest, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbox_NoNetwork, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblEmail As Windows.Forms.Label
    Friend WithEvents BtnUpdateAddIn As Windows.Forms.Button
    Friend WithEvents Pbox_New As Windows.Forms.PictureBox
    Friend WithEvents PBox_Web As Windows.Forms.PictureBox
    Friend WithEvents Pbox_Latest As Windows.Forms.PictureBox
    Friend WithEvents LblVersion As Windows.Forms.Label
    Friend WithEvents Cancel_Button As Windows.Forms.Button
    Friend WithEvents pbActualiza As Windows.Forms.ProgressBar
    Friend WithEvents lbl_Terms As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents pbox_NoNetwork As Windows.Forms.PictureBox
End Class
