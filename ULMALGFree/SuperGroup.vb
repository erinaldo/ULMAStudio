Imports System.Windows.Forms
Imports System.Drawing
Imports uf = ULMALGFree.clsBase
Imports ufd = ULMALGFree.Datos

Public Class SuperGroup
    Public sgCode As String = ""        ' <superGroupCode>EV</superGroupCode>
    Public sgDescription As String      ' <superGroupDescription language="en">Vertical formwork</superGroupDescription>
    Public sgShortName As String = ""   ' <superGroupShortName language="en">VERT.</superGroupShortName>
    Public sgImage As Image = Nothing            ' <superGroupIcon>wall_column_formwork_icon.jpg</superGroupIcon>
    'Public sgImageSelected As Image = Nothing    ' <superGroupIconSelected>wall_column_formwork_icon_selected.jpg</superGroupIconSelected>
    ' <groups>
    Public groups As List(Of Group)
    '
    Public SgPanel As Panel = Nothing
    Public pBox As PictureBox = Nothing
    Public WithEvents SgButton As Button = Nothing
    Public SgLabel As Label = Nothing
    Public Sgtt As ToolTip = Nothing
    Public menu As ContextMenuStrip = Nothing
    Public WithEvents UpdateSuperGroup As ToolStripMenuItem
    Public WithEvents RemoveSuperGroup As ToolStripMenuItem
    Private _nActualizaciones As Integer = 0

    Public Property nActualizaciones As Integer
        Get
            Return _nActualizaciones
        End Get
        Set(value As Integer)
            _nActualizaciones = value
            If value = 0 Then
                Me.pBox.Image = Nothing
                Me.pBox.BackColor = Color.Transparent
            Else
                Me.pBox.Image = uf.lNumeros.Images(_nActualizaciones)
            End If
        End Set
    End Property
    '
    Public Sub New(fS As FlowLayoutPanel, fG As FlowLayoutPanel, pI As PictureBox, gCode As String, gDes As String, gName As String, gImg As Image)
        If uf.frmUFam.fpSuperGroups Is Nothing Then uf.frmUFam.fpSuperGroups = fS
        If uf.frmUFam.fpGroups Is Nothing Then uf.frmUFam.fpGroups = fG
        If uf.frmUFam.PbImagen Is Nothing Then uf.frmUFam.PbImagen = pI
        ' ***** Cargar las imágenes la primera vez sólo
        If uf.lNumeros Is Nothing Then
            uf.lNumeros = New ImageList()
            uf.lNumeros.ImageSize = New Size(24, 24)
            For x As Integer = 0 To 10
                uf.lNumeros.Images.Add(x, Image.FromFile(IO.Path.Combine(uf._imgFolder, x & ".png"))) : uf.lNumeros.Images(x).Tag = x
            Next
        End If
        '
        If uf.lupdate Is Nothing Then
            uf.lupdate = New ImageList()
            uf.lupdate.ImageSize = New Size(24, 24)
            uf.lupdate.Images.Add("Updated", Image.FromFile(IO.Path.Combine(uf._imgFolder, "Updated3.png"))) : uf.lupdate.Images(0).Tag = "Updated"    ' 0 Esta correcto
            uf.lupdate.Images.Add("Download", Image.FromFile(IO.Path.Combine(uf._imgFolder, "Download3.png"))) : uf.lupdate.Images(1).Tag = "Download"   ' 1 Hay que descargarlo. No estaba
            uf.lupdate.Images.Add("DownloadUpdate", Image.FromFile(IO.Path.Combine(uf._imgFolder, "DownloadUpdate3.png"))) : uf.lupdate.Images(2).Tag = "DownloadUpdate" ' 2 Hay que descargarlo y actualizarlo. Ya estaba
        End If
        ' **********************************************
        'menu = New ContextMenuStrip()
        Me.sgCode = gCode.Trim
        Me.sgDescription = gDes.Trim
        Me.sgShortName = gName.Trim
        Me.sgImage = gImg
        Sgtt = New ToolTip
        '
        groups = New List(Of [group])
        CreaPanel()
        'menu_rellena()
    End Sub

    Public Sub CreaPanel()
        SgPanel = New System.Windows.Forms.Panel
        SgPanel.Width = 32
        SgPanel.Height = 32
        SgPanel.Margin = New Padding(0, 0, 0, 0)
        SgPanel.Padding = New Padding(0, 0, 0, 0)
        'SgPanel.BackColor = Color.White
        '
        SgButton = New Button
        SgButton.Name = "Button"
        SgButton.Width = SgPanel.Width - 1
        SgButton.Height = SgPanel.Height - 1
        'SgButton.BackColor = Color.LightGray
        SgButton.FlatStyle = FlatStyle.Flat
        SgButton.FlatAppearance.BorderSize = 0
        SgButton.ForeColor = System.Drawing.Color.Black
        'SgButton.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        'SgButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.LightGray
        'tButton.BackgroundImage = New Bitmap(Me.sgImage, 64, 64)
        SgButton.Image = New Bitmap(Me.sgImage, 31, 31)
        SgButton.ImageAlign = ContentAlignment.MiddleCenter
        'SgButton.Text = Me.sgShortName
        'SgButton.TextAlign = ContentAlignment.BottomCenter
        'SgButton.TextImageRelation = TextImageRelation.ImageAboveText
        'SgButton.Font = New Font(SgButton.Font.Name, 6)
        'If Me.sgShortName.Length > 6 Then
        '    SgButton.Font = New Font(SgButton.Font.Name, SgButton.Font.Size - 2)
        'End If
        SgButton.UseVisualStyleBackColor = True
        'SgButton.FlatStyle = FlatStyle.Flat
        'SgButton.FlatAppearance.BorderSize = 0
        'SgButton.ContextMenuStrip = menu
        Sgtt.SetToolTip(SgButton, Me.sgDescription)
        '
        Dim img As Image = Nothing  ' SuperGroup.lNumeros.Images(Aleatorio_Dame(0, 10))
        pBox = New PictureBox
        pBox.Name = "PictureBox"
        pBox.Size = New Size(14, 14)
        pBox.Image = img
        pBox.BackColor = Color.Transparent
        pBox.SizeMode = PictureBoxSizeMode.Zoom
        pBox.BackColor = Color.Transparent
        '
        SgButton.Controls.Add(pBox)
        SgButton.Controls.Item(0).Location = New Point(SgButton.Width - 13, 0)
        SgPanel.Controls.Add(SgButton)
        SgPanel.Controls.Item(0).Location = New Point(0, 0)
        'SgPanel.Controls.Item(0).Dock = DockStyle.Fill
        SgPanel.Tag = Me
        'menu_rellena()
    End Sub

    Private Sub tButton_Click(sender As Object, e As EventArgs) Handles SgButton.Click
        PonFondoBlancoSuperGroups()
        'SgButton.BackColor = Color.LightGray
        If uf.frmUFam.fpGroups.Visible = False Then uf.frmUFam.fpGroups.Visible = True
        uf.frmUFam.PbImagen.Image = uf._imgBase : uf.frmUFam.PbImagen.Refresh()
        'For Each oC As Control In SgButton.Parent.Controls
        '    oC.BackColor = SgButton.BackColor
        'Next
        uf.frmUFam.fpGroups.Visible = True
        Me.PonBotonesGrupos()
        uf.UltimoSuperGrupo = Me
    End Sub
    '
    Public Sub PonBotonesGrupos()
        If groups Is Nothing OrElse groups.Count = 0 Then Exit Sub
        uf.frmUFam.fpGroups.Controls.Clear()
        uf.frmUFam.fpGroups.Controls.Clear()
        Dim primeraVuelta As Boolean = True
        For Each g As group In groups
            g.CreaPanel()
            If primeraVuelta Then
                '2019/10/28 Xabier Calvo: Altura LayoutPanel ajustado a los grupos que se muestran
                uf.frmUFam.fpGroups.Height = groups.Count * groups(0).gPanel.Height + 2
                uf.frmUFam.fpGroups.Width = groups(0).gPanel.Width + 3
                If uf.frmUFam.fpGroups.Height > 500 Then
                    uf.frmUFam.fpGroups.Height = 500
                    uf.frmUFam.fpGroups.AutoScroll = True
                End If
                primeraVuelta = False
            End If
            uf.frmUFam.fpGroups.Controls.Add(g.gPanel)
            uf.frmUFam.fpGroups.Controls.Add(g.gPanel)
        Next
    End Sub
    Public Sub PonFondoBlancoSuperGroups()
        'If uf.frmUFam.fpSuperGroups Is Nothing OrElse uf.frmUFam.fpSuperGroups.Controls.Count = 0 Then Exit Sub
        'For Each oC As Control In uf.frmUFam.fpSuperGroups.Controls
        '    Dim oP As Panel = oC
        '    oP.BackColor = Color.White
        '    oP.Update()
        '    For Each oC1 As Control In oC.Controls
        '        oC1.BackColor = Color.White
        '        oC1.Update()
        '    Next
        'Next
        'uf.frmUFam.fpSuperGroups.Update()
    End Sub
    Public Sub menu_rellena()
        UpdateSuperGroup = New ToolStripMenuItem("Download all " & sgShortName)
        menu.Items.Add(UpdateSuperGroup)
        RemoveSuperGroup = New ToolStripMenuItem("Remove all " & sgShortName)
        menu.Items.Add(RemoveSuperGroup)
    End Sub

    Private Sub ActualizaTodo_Click(sender As Object, e As EventArgs) Handles UpdateSuperGroup.Click
        MsgBox("Download all groups in " & sgDescription)
    End Sub

    Private Sub QuitarActualizacion_Click(sender As Object, e As EventArgs) Handles RemoveSuperGroup.Click
        MsgBox("Remove all groups Downloads in " & sgDescription)
    End Sub

    Private Sub SgButton_MouseClick(sender As Object, e As MouseEventArgs) Handles SgButton.MouseClick
        uf.UltimoSuperGrupo = Me
    End Sub

    Public Sub simulaClickBoton()
        uf.UltimoSuperGrupo = Me
    End Sub
End Class


