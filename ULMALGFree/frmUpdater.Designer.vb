Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdater))
        Me.fpSuperGroups = New System.Windows.Forms.FlowLayoutPanel()
        Me.fpGroups = New System.Windows.Forms.FlowLayoutPanel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.LblAction = New System.Windows.Forms.Label()
        Me.PbImagen = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.PbImagen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fpSuperGroups
        '
        Me.fpSuperGroups.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.fpSuperGroups.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.fpSuperGroups.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.fpSuperGroups.Location = New System.Drawing.Point(20, 38)
        Me.fpSuperGroups.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.fpSuperGroups.Name = "fpSuperGroups"
        Me.fpSuperGroups.Size = New System.Drawing.Size(40, 560)
        Me.fpSuperGroups.TabIndex = 0
        '
        'fpGroups
        '
        Me.fpGroups.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.fpGroups.AutoScroll = True
        Me.fpGroups.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.fpGroups.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.fpGroups.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.fpGroups.Location = New System.Drawing.Point(62, 38)
        Me.fpGroups.Margin = New System.Windows.Forms.Padding(0)
        Me.fpGroups.Name = "fpGroups"
        Me.fpGroups.Size = New System.Drawing.Size(132, 560)
        Me.fpGroups.TabIndex = 4
        Me.fpGroups.Visible = False
        Me.fpGroups.WrapContents = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ProgressBar1.Location = New System.Drawing.Point(519, 605)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(320, 25)
        Me.ProgressBar1.TabIndex = 7
        Me.ProgressBar1.Visible = False
        '
        'LblAction
        '
        Me.LblAction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAction.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblAction.Location = New System.Drawing.Point(23, 605)
        Me.LblAction.Name = "LblAction"
        Me.LblAction.Size = New System.Drawing.Size(359, 27)
        Me.LblAction.TabIndex = 8
        Me.LblAction.Text = "Action:"
        Me.LblAction.Visible = False
        '
        'PbImagen
        '
        Me.PbImagen.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PbImagen.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.PbImagen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PbImagen.Location = New System.Drawing.Point(196, 38)
        Me.PbImagen.Margin = New System.Windows.Forms.Padding(0)
        Me.PbImagen.MaximumSize = New System.Drawing.Size(644, 561)
        Me.PbImagen.MinimumSize = New System.Drawing.Size(644, 561)
        Me.PbImagen.Name = "PbImagen"
        Me.PbImagen.Size = New System.Drawing.Size(644, 561)
        Me.PbImagen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PbImagen.TabIndex = 3
        Me.PbImagen.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri Light", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(21, 11)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(437, 15)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Use the menu to download the systems you want to have available in the Browser:"
        '
        'frmUpdater
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(859, 635)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblAction)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.fpGroups)
        Me.Controls.Add(Me.PbImagen)
        Me.Controls.Add(Me.fpSuperGroups)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdater"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Download"
        CType(Me.PbImagen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public Sub cambiaFondosGpanel(ByRef panelSel As Panel)
        For Each obj In fpGroups.Controls
            Dim panel As Panel
            panel = TryCast(obj, Panel)
            If panel.Controls(0).Text <> panelSel.Controls(0).Text Then
                If panel.BackColor <> System.Drawing.Color.White Then

                    panel.BackColor = System.Drawing.Color.White
                End If
            End If
        Next
    End Sub

    Public Sub cambiaFondosSeleccionadoGpanel(ByRef panelSel As Panel)
        Dim colorLightGray As System.Drawing.Color = System.Drawing.Color.FromArgb(245, 245, 245)
        Dim colorDarkGray As System.Drawing.Color = System.Drawing.Color.FromArgb(184, 181, 173)
        Dim colorLightBlue As System.Drawing.Color = System.Drawing.Color.FromArgb(229, 244, 251)
        Dim colorDarkBlue As System.Drawing.Color = System.Drawing.Color.FromArgb(203, 232, 246)

        For Each obj In fpGroups.Controls
            Dim panel As Panel
            panel = TryCast(obj, Panel)
            If panel.Controls(0).Text <> panelSel.Controls(0).Text Then
                If panel.BackColor = colorLightBlue Then
                    panel.BackColor = System.Drawing.Color.Transparent
                End If
            Else
                If panel.BackColor <> colorDarkBlue Then
                    panel.BackColor = colorLightBlue
                End If
            End If
        Next
    End Sub


    Public Sub cambiaFondosSGpanel(ByRef panelSel As Panel)
        For Each obj In fpSuperGroups.Controls
            Dim panel As Panel
            panel = TryCast(obj, Panel)
            If panel.ForeColor <> panelSel.ForeColor Then
                If panel.BackColor <> System.Drawing.Color.White Then

                    panel.BackColor = System.Drawing.Color.White
                    panel.ForeColor = System.Drawing.Color.White
                End If
            End If
        Next
    End Sub

    Public Sub cambiaFondosSGpanelABlanco()
        For Each obj In fpSuperGroups.Controls
            Dim panel As Panel
            panel = TryCast(obj, Panel)
            panel.ForeColor = System.Drawing.Color.White
        Next
    End Sub

    Public WithEvents fpSuperGroups As FlowLayoutPanel
    Public WithEvents PbImagen As PictureBox
    Public WithEvents fpGroups As FlowLayoutPanel
    Public WithEvents ProgressBar1 As ProgressBar
    Public WithEvents LblAction As Label
    Friend WithEvents Label1 As Label
End Class
