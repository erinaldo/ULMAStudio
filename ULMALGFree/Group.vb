Imports System.Windows.Forms
Imports System.Drawing
Imports uf = ULMALGFree.clsBase
Imports ULMALGFree

Public Class Group
    Public gCode As String = ""             ' <groupCode>E1900</groupCode>
    Public gDescription As String = ""      ' <groupDescription language = "en" >ORMA FORMWORK</groupDescription>
    Public gShortName As String = ""        ' <groupShortName language = "en" >ORMA</groupShortName>
    Public gImage As Image = Nothing        ' <groupImage>ORMA_render.jpg</groupImage>
    Public gStatus As status
    '
    Public gPanel As Panel = Nothing
    Public WithEvents gButton As Button = Nothing
    Public WithEvents gButton2 As Button = Nothing
    Public pBox As PictureBox = Nothing
    Public gLabel As Label = Nothing
    Public gTT As ToolTip = Nothing
    Public gTT2 As ToolTip = Nothing
    Private _action As queAction = queAction.updated
    '
    Public menu As ContextMenuStrip = Nothing
    Public WithEvents DownloadGroup As ToolStripMenuItem
    Public WithEvents UpdateGroup As ToolStripMenuItem
    Public WithEvents RemoveGroup As ToolStripMenuItem

    'Colores
    Public colorSnowFlake As Color
    Public colorLightGray As Color
    Public colorDarkGray As Color
    Public lightBlue As Color
    Public darkBlue As Color
    Public cLcsv As ULMALGFree.clsBase

    Public Property Action As queAction
        Get
            Return _action
        End Get
        Set(value As queAction)
            _action = value
            If pBox Is Nothing Then
                pBox = New PictureBox
                pBox.Size = New Size(32, 32)
                pBox.BackColor = Color.Transparent
                pBox.SizeMode = PictureBoxSizeMode.Zoom
            End If
            If value = queAction.updated Then
                Me.pBox.Image = uf.lupdate.Images.Item(0)
            ElseIf value = queAction.toupdate Then
                Me.pBox.Image = uf.lupdate.Images.Item(1)
            ElseIf value = queAction.notupdated Then
                Me.pBox.Image = uf.lupdate.Images.Item(2)
            End If
        End Set
    End Property

    Public Sub New(gC As String, gD As String, gS As String, gI As Image)
        If cLcsv Is Nothing Then cLcsv = New ULMALGFree.clsBase(System.Reflection.Assembly.GetExecutingAssembly)
        Me.gCode = gC.Trim
        Me.gDescription = gD.Trim
        Me.gShortName = gS.Trim
        Me.gImage = gI
        gTT = New ToolTip
        gTT2 = New ToolTip
        menu = New ContextMenuStrip()
        colorSnowFlake = Color.FromArgb(240, 240, 240)
        colorLightGray = Color.FromArgb(245, 245, 245)
        colorDarkGray = Color.FromArgb(184, 181, 173)
        lightBlue = Color.FromArgb(229, 244, 251)
        darkBlue = Color.FromArgb(203, 232, 246)
        'CreaPanel()
        'menu_rellena()
    End Sub

    Public Sub CreaPanel()
        gPanel = New System.Windows.Forms.Panel
        'gPanel.Width = 120
        gPanel.Width = 130
        'gPanel.Height = 120
        gPanel.Height = 32
        gPanel.BackColor = Color.White
        gPanel.BorderStyle = BorderStyle.None
        gPanel.BackgroundImage = My.Resources.FondoPanel
        gPanel.BackgroundImageLayout = ImageLayout.Center
        gPanel.Margin = New Padding(0, 0, 0, 0)
        gPanel.Padding = New Padding(0, 0, 0, 0)
        'gPanel.LayoutEngin
        '
        gButton = New Button
        gButton.Width = gPanel.Width - 35
        gButton.Height = gPanel.Height - 8
        gButton.BackColor = Color.Transparent
        gButton.FlatStyle = FlatStyle.Flat
        gButton.FlatAppearance.BorderSize = 0
        gButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        gButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        'gButton.BackColor = Color.Transparent
        'gButton.BackgroundImage = New Bitmap(Me.sgImage, 64, 64)
        'gButton.Image = New Bitmap(Me.gImage, 64, 64)
        'gButton.ImageAlign = ContentAlignment.MiddleCenter
        gButton.Text = Me.gShortName.ToUpper
        'gButton.TextAlign = ContentAlignment.BottomCenter
        gButton.TextAlign = ContentAlignment.BottomLeft
        gButton.TextImageRelation = TextImageRelation.ImageAboveText
        gButton.Font = New Font("Calibri", 8)
        'If Me.gShortName.Length > 8 Then
        '    gButton.Font = New Font(gButton.Font.Name, gButton.Font.Size - 2)
        'End If
        'gButton.UseVisualStyleBackColor = True
        'gButton.FlatStyle = FlatStyle.Standard
        'gButton.FlatAppearance.BorderSize = 0
        'gButton.ContextMenuStrip = menu
        gTT.SetToolTip(gButton, Me.gDescription)
        '
        'Dim img As Image = Grupo_PonImageAction()
        'If pBox Is Nothing Then
        '    pBox = New PictureBox
        '    pBox.Image = img
        '    pBox.Size = New Size(24, 24)
        '    pBox.BackColor = Color.Transparent
        '    pBox.SizeMode = PictureBoxSizeMode.Zoom
        '    AddHandler pBox.Click, AddressOf pbox_Click
        '    AddHandler pBox.MouseDown, AddressOf pbox_MouseDown
        'End If
        '
        gButton2 = New Button
        gButton2.Width = gPanel.Height - 6
        gButton2.Height = gPanel.Height - 8
        gButton2.BackColor = Color.Transparent
        gButton2.FlatStyle = Windows.Forms.FlatStyle.Flat
        gButton2.FlatAppearance.BorderSize = 0
        gButton2.FlatAppearance.MouseDownBackColor = Color.Transparent
        gButton2.FlatAppearance.MouseOverBackColor = Color.Transparent
        'gButton2.BackColor = Color.Transparent
        gButton2.Image = Grupo_PonImageAction()
        'AddHandler gButton2.Click, AddressOf gButton2_Click
        AddHandler gButton2.MouseDown, AddressOf gButton2_MouseDown
        gButton2.ContextMenuStrip = menu
        ''''
        'gButton.Controls.Add(pBox)
        'gButton.Controls.Item(0).Location = New Point(gButton.Width - 26, 2)
        'gButton.Controls.Item(0).Location = New Point(gButton.Width - 26, gButton.Height / 2 - 12)
        gButton.Tag = Me
        gButton2.Tag = Me
        gPanel.Controls.Add(gButton)
        gPanel.Controls.Item(0).Location = New Point(3, 3)
        'gPanel.Controls.Item(0).Dock = DockStyle.Fill
        gPanel.Controls.Add(gButton2)
        gPanel.Controls.Item(1).Location = New Point(gPanel.Width - 32, 3)
        gPanel.Tag = Me
    End Sub

    Public Function Grupo_PonImageAction() As Image
        Call uf.INIUpdates_LeeTODO()
        Dim img As Image = Nothing  ' SuperGroup.lupdate.Images(Aleatorio_Dame(0, 2))
        'lupdate.Images.Add("updated", Image.FromFile(IO.Path.Combine(_imgFolder, "updated.png"))) : lupdate.Images(0).Tag = "updated"    ' 0
        'lupdate.Images.Add("toupdate", Image.FromFile(IO.Path.Combine(_imgFolder, "toupdate.png"))) : lupdate.Images(1).Tag = "toupdate"
        'lupdate.Images.Add("notupdated", Image.FromFile(IO.Path.Combine(_imgFolder, "notupdated.png"))) : lupdate.Images(2).Tag = "notupdated"
        Dim nAction As Integer = Grupo_DameActionNumero(Me.gCode, Me.gShortName)
        If nAction = 0 Then
            ' No tiene actualizaciones
            img = uf.lupdate.Images.Item(0) '"Updated"
            Me.Action = queAction.updated   ' "Updated"
            gTT2.SetToolTip(gButton2, "Product Updated")
        ElseIf nAction = 1 Then
            img = uf.lupdate.Images.Item(1) '"Download"
            Me.Action = queAction.toupdate  ' "Download"
            gTT2.SetToolTip(gButton2, "Download Product")
        ElseIf nAction = 2 Then
            img = uf.lupdate.Images.Item(2) '"DownloadUpdate"
            Me.Action = queAction.notupdated  '"DownloadUpdate"
            gTT2.SetToolTip(gButton2, "Download Update")
        End If
        Return img
    End Function

    Public Function Grupo_DameAction() As queAction
        Dim resultado As queAction = queAction.updated
        Dim nAction As Integer = Grupo_DameActionNumero(Me.gCode, Me.gShortName)
        '
        If nAction = 0 Then
            ' Esta ya descargado y actualizado
            resultado = queAction.updated   ' "updated"
        ElseIf nAction = 1 Then
            ' Hay que descargarlo. No esta en families
            resultado = queAction.toupdate  ' "toupdate"
        ElseIf nAction = 2 Then
            ' Hay que descargarlo. Si esta en families, pero tiene actualizacion.
            resultado = queAction.notupdated    ' "notupdated"
        End If
        Return resultado
    End Function
    Private Sub gButton_Click(sender As Object, e As EventArgs) Handles gButton.Click
        'PonFondoBlancoGroups()
        'gButton.BackColor = Color.LightGray
        'For Each oC As Control In gButton.Parent.Controls
        '    oC.BackColor = gButton.BackColor
        'Next
        'gButton.Parent.BackColor = gButton.BackColor
        'uf.frmUFam.PbImagen.Size = New Size(857, 689)
        'uf.frmUFam.PbImagen.MaximumSize = New Size(857, 689)
        'uf.frmUFam.PbImagen.SizeMode = PictureBoxSizeMode.Zoom
        uf.frmUFam.PbImagen.Image = uf._imgBase : uf.frmUFam.PbImagen.Refresh()
        uf.frmUFam.PbImagen.Image = gImage
        'gButton.ContextMenuStrip.Show()
        uf.UltimoGrupo = Me
        gPanel.BackColor = darkBlue
        uf.frmUFam.cambiaFondos(gPanel)

        'menu_rellena()
        'gButton.ContextMenuStrip.Show() 'New Point(gButton.Left + gButton.Width, gButton.Top))
        'gButton.PerformClick()
        'gButton.ContextMenu.Show()
    End Sub


    Private Sub gButton2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'uf.frmUFam.PbImagen.Image = gImage
        'uf.UltimoGrupo = Me
        'gPanel.BackColor = darkBlue
        'uf.frmUFam.cambiaFondos(gPanel)
        gButton_Click(Nothing, Nothing)
        If e.Button = MouseButtons.Right Then
            menu_rellena()
        ElseIf e.Button = MouseButtons.Left Then
            If Me.Action = queAction.toupdate Then          ' No se ha descargado nunca. Descargarlo
                DownloadGroupFromPictureBox()
            ElseIf Me.Action = queAction.notupdated Then    ' Si se ha descargado alguna vez. Actualizarlo
                UpdateGroupFromPictureBox()
            End If
        End If
        ' etc...
    End Sub


    Private Sub gButton_MouseEnter(sender As Object, e As EventArgs) Handles gButton.MouseEnter
        'gButton.FlatAppearance.MouseOverBackColor = colorLightGray
        'gButton2.BackColor = colorLightGray
        uf.frmUFam.cambiaFondosSeleccionado(gPanel)
    End Sub

    Private Sub gButton_MouseLeave(sender As Object, e As EventArgs) Handles gButton.MouseLeave
        'gButton.FlatAppearance.MouseOverBackColor = colorSnowFlake
        'gButton2.BackColor = colorSnowFlake
    End Sub

    Private Sub gButton2_MouseLeave(sender As Object, e As EventArgs) Handles gButton2.MouseLeave
        'gButton.BackColor = colorSnowFlake
        'gButton2.FlatAppearance.MouseOverBackColor = colorSnowFlake
    End Sub

    Private Sub gButton2_MouseEnter(sender As Object, e As EventArgs) Handles gButton2.MouseEnter
        'gButton.BackColor = colorLightGray
        'gButton2.FlatAppearance.MouseOverBackColor = colorLightGray
        uf.frmUFam.cambiaFondosSeleccionado(gPanel)
    End Sub



    'Public Sub PonFondoBlancoGroups()
    '    'If uf.frmUFam.fpGroups Is Nothing OrElse uf.frmUFam.fpGroups.Controls.Count = 0 Then Exit Sub
    '    'For Each oC As Control In uf.frmUFam.fpGroups.Controls
    '    '    Dim oP As Panel = oC
    '    '    oP.BackColor = Color.White
    '    '    oP.Update()
    '    '    For Each oC1 As Control In oC.Controls
    '    '        oC1.BackColor = Color.White
    '    '        oC1.Update()
    '    '        'If TypeOf oC1 Is Button Then
    '    '        '    CType(oC1, Button).BackgroundImage = New Bitmap(CType(oP.Tag, SuperGroup).sgImage, 64, 64)
    '    '        'End If
    '    '    Next
    '    'Next
    '    'uf.frmUFam.fpGroups.Update()
    'End Sub
    Public Sub menu_rellena()
        If menu IsNot Nothing AndAlso menu.Items.Count > 0 Then
            menu.Items.Clear()
        End If

        'DownloadGroup = New ToolStripMenuItem("Download " & gShortName)
        'UpdateGroup = New ToolStripMenuItem("Update " & gShortName)
        RemoveGroup = New ToolStripMenuItem("Remove " & gShortName)
        '
        'If Me.Action = queAction.toupdate Then          ' No se ha descargado nunca. Descargarlo
        '    menu.Items.Add(DownloadGroup)
        'ElseIf Me.Action = queAction.notupdated Then    ' Si se ha descargado alguna vez. Actualizarlo
        '    menu.Items.Add(UpdateGroup)
        '    menu.Items.Add(RemoveGroup)
        'ElseIf Me.Action = queAction.updated Then       ' Ya está descargado y correcto
        '    menu.Items.Add(RemoveGroup)
        'End If

        If Me.Action = queAction.notupdated Or Me.Action = queAction.updated Then
            menu.Items.Add(RemoveGroup)
        End If
    End Sub


    'Private Sub DownloadGroup_Click(sender As Object, e As EventArgs) Handles DownloadGroup.Click
    '    'No hace falta borrar. Nunca han estado estas familias
    '    For Each d As ULMALGFree.Datos In uf.LUp
    '        If d.Local_File.Contains(gCode) AndAlso d.Local_File.Contains(gShortName) Then
    '            uf.frmUFam.ProgressBar1.Visible = True
    '            uf.frmUFam.LblAction.Visible = True
    '            uf.FTP_DescargarYDescomprimir(Me, d, uf.frmUFam.LblAction, uf.frmUFam.ProgressBar1)
    '            uf.frmUFam.LblAction.Text = "Action:"
    '            uf.frmUFam.ProgressBar1.Value = 0
    '            Exit For
    '        End If
    '    Next
    '    'MsgBox("Product downloaded successfully", MsgBoxStyle.Information, "Product Download")
    '    Me.Action = queAction.updated
    '    uf.frmUFam.ProgressBar1.Visible = False
    '    uf.frmUFam.LblAction.Visible = False
    '    ' No hacer nada mas. Ya lo hace FTP_DescargarYDescomprimir 
    '    ' Actualizar número en SuperGrupo
    '    'uf.UltimoSuperGrupo.nActualizaciones -= 1
    '    'uf.UltimoSuperGrupo.SgButton.PerformClick()
    'End Sub

    'Private Sub UpdateGroup_Click(sender As Object, e As EventArgs) Handles UpdateGroup.Click
    '    'Borrar antes los que hubiera. Ya habría familias antes en families
    '    'Dim task As New System.Threading.Thread(AddressOf BorraFamiliasGroup) : task.Start()
    '    BorraFamiliasGroup()
    '    '
    '    For Each d As ULMALGFree.Datos In uf.LUp
    '        If d.Local_File.Contains(gCode) AndAlso d.Local_File.Contains(gShortName) Then
    '            uf.frmUFam.ProgressBar1.Visible = True
    '            uf.frmUFam.LblAction.Visible = True
    '            uf.FTP_DescargarYDescomprimir(Me, d, uf.frmUFam.LblAction, uf.frmUFam.ProgressBar1)
    '            uf.frmUFam.LblAction.Text = "Action:"
    '            uf.frmUFam.ProgressBar1.Value = 0
    '            Exit For
    '        End If
    '    Next
    '    'MsgBox("Product updated successfully", MsgBoxStyle.Information, "Product Update")
    '    Me.Action = queAction.toupdate
    '    uf.frmUFam.ProgressBar1.Visible = False
    '    uf.frmUFam.LblAction.Visible = False
    '    ' No hacer nada mas. Ya lo hace FTP_DescargarYDescomprimir 
    '    'uf.UltimoSuperGrupo.SgButton.PerformClick()
    'End Sub

    Private Sub RemoveGroup_Click(sender As Object, e As EventArgs) Handles RemoveGroup.Click
        Dim result As DialogResult = MsgBox("Do you want to remove " + gButton.Text + "?", MsgBoxStyle.OkCancel, "Product Remove Confirmation")
        If result = DialogResult.OK Then
            'Borrar antes los que hubiera. Ya habría familias antes en families
            uf.frmUFam.ProgressBar1.Visible = True
            uf.frmUFam.LblAction.Visible = True
            If ULMALGFree.clsBase.cIni Is Nothing Then ULMALGFree.clsBase.cIni = New ULMALGFree.clsINI
            ' Leer el valor actual de LAST
            Dim valor As String = ULMALGFree.clsBase.cIni.IniGet(ULMALGFree.clsBase._IniUpdaterFull, "LAST", Me.gCode)
            ' Escribirlo en UPDATES
            ULMALGFree.clsBase.cIni.IniWrite(ULMALGFree.clsBase._IniUpdaterFull, "UPDATES", Me.gCode, valor)
            ' Borrarlo de LAST
            ULMALGFree.clsBase.cIni.IniDeleteKey(ULMALGFree.clsBase._IniUpdaterFull, "LAST", Me.gCode)
            '2019/11/20 Xabier Calvo: Al mostrar solo boton rojo por actualizaciones (no descarga de producto), eliminar un porducto no significa tener una nueva actualizacion
            'Quitamos actualización si eliminamos producto a actualizar
            Dim oAction As ULMALGFree.queAction = Me.Grupo_DameAction()
            uf.INIUpdates_LeeTODO()
            Me.Action = queAction.toupdate
            ' Actualizar número en SuperGrupo
            If oAction = ULMALGFree.queAction.toupdate Then
                uf.UltimoSuperGrupo.nActualizaciones -= 1
            End If
            'uf.UltimoSuperGrupo.SgButton.PerformClick()
            Dim task As New System.Threading.Thread(AddressOf BorraFamiliasGroup_UnGrupo) : task.Start()
                'MsgBox("Product removed successfully", MsgBoxStyle.Information, "Product Remove")
                uf.frmUFam.ProgressBar1.Visible = False
                uf.frmUFam.LblAction.Visible = False
                gButton2.Image = Grupo_PonImageAction()
                uf._recargarBrowser = True
                uf.cambiosEnGrupos = True
                uf.yo.PonLog_ULMA("REMOVE_GROUP", UPDATE_GROUP:=gCode, NOTES:="Name=" & gButton.Text)
            End If
    End Sub

    'Private Sub gButton_MouseClick(sender As Object, e As MouseEventArgs) Handles gButton.MouseClick
    '    gButton.ContextMenuStrip.Show(gButton.PointToScreen(e.Location)) 'New Point(gButton.Left + gButton.Width, gButton.Top))
    'End Sub

    'Private Sub gButton2_MouseClick(sender As Object, e As MouseEventArgs) Handles gButton2.MouseClick
    '    gButton2.ContextMenuStrip.Show(gButton2.PointToScreen(e.Location)) 'New Point(gButton.Left + gButton.Width, gButton.Top))
    'End Sub

    Public Sub BorraFamiliasGroup_UnGrupo()
        'If uf._ActualizarXMLs = True Then 'nuf.cUp("XML").Count > 0
        '    Dim fiXmlFtp As String = uf.cIni.IniGet(uf._IniUpdaterFull, "UPDATES", "XML")
        '    Dim d As New ULMALGFree.Datos("XML", fiXmlFtp)
        '    uf.FTP_DescargarYDescomprimirXML(d)
        'End If

        If uf.colGroups Is Nothing OrElse uf.colGroups.Count = 0 Then Exit Sub
        If uf.colGroups.ContainsKey(gCode) = False Then Exit Sub
        '

        Dim oG As clsGroup = uf.colGroups(gCode)
        Dim LBorrarTemp As New List(Of String)
        Dim LBorrarEnd As New List(Of String)
        '
        For Each fiRFA In IO.Directory.GetFiles(uf.path_families_base, "*.rfa", IO.SearchOption.TopDirectoryOnly)
            Dim nombresin As String = IO.Path.GetFileNameWithoutExtension(fiRFA)
            Dim partes() As String = nombresin.Split("_"c)
            Dim article As String = partes(UBound(partes))
            If oG.LstarticleCodes.Contains(article) Then
                LBorrarTemp.Add(fiRFA)
            End If
        Next
        ' Recorrer los ficheros candidatos a borrar. Para comprobar si están en otros grupos
        For Each fiBorra In LBorrarTemp
            Dim estaenotros As Boolean = False
            For Each oGTemp As clsGroup In uf.colGroups.Values
                If oG.groupCode = oGTemp.groupCode Then
                    Continue For
                End If
                If oGTemp.LstFilenameOnly.Contains(IO.Path.GetFileNameWithoutExtension(fiBorra)) = True Then
                    estaenotros = True
                    Exit For
                End If
            Next
            ' Si no está en otros grupos, añadir a la lista definitiva de borrado
            If estaenotros = False Then
                LBorrarEnd.Add(fiBorra)
            End If
        Next
        '
        For Each fiBorra In LBorrarEnd
            ' Quitamos el que borre las imágenes PNG de las familias borradas (Para que puedan salir en informes)
            'Dim fiImg As String = IO.Path.Combine(uf.path_families_base_images, IO.Path.GetFileNameWithoutExtension(fiBorra) & ".png")
            Try
                If IO.File.Exists(fiBorra) Then IO.File.Delete(fiBorra)
                'If IO.File.Exists(fiImg) Then IO.File.Delete(fiImg)
            Catch ex As Exception

            End Try
        Next
        ' Log de la Familia borrada.
        If uf.yo Is Nothing Then uf.yo = New clsBase(Reflection.Assembly.GetExecutingAssembly)
        'uf.yo.PonLog_ULMA(ULMALGFree.ACTION.REMOVE_FAMILIES,,,,, gCode)
        '
        LBorrarTemp = Nothing
        LBorrarEnd = Nothing
        oG = Nothing
    End Sub

    'Private Sub pbox_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Dim pic As PictureBox = DirectCast(sender, PictureBox)
    '    If Me.Action = queAction.toupdate Then          ' No se ha descargado nunca. Descargarlo
    '        DownloadGroupFromPictureBox()
    '    ElseIf Me.Action = queAction.notupdated Then    ' Si se ha descargado alguna vez. Actualizarlo
    '        UpdateGroupFromPictureBox()
    '    End If
    '    ' etc...
    'End Sub

    Private Sub DownloadGroupFromPictureBox()
        Dim result As DialogResult = MsgBox("Do you want to download " + gButton.Text + "?", MsgBoxStyle.OkCancel, "Product Download Confirmation")
        If result = DialogResult.OK Then
            uf.INIUpdates_LeeTODO()
            ' Si hay actualizaciones de los XML, descargar antes.
            If uf.cUp("xmls").Count > 0 Then
                Dim fiXmlFtp As String = uf.cIni.IniGet(uf._IniUpdaterFull, "UPDATES", "XML")
                Dim d As New ULMALGFree.Datos("XML", fiXmlFtp)
                uf.FTP_DescargarYDescomprimirXML(d)
            End If
            '
            'No hace falta borrar. Nunca han estado estas familias
            For Each d As ULMALGFree.Datos In uf.LUp
                If d.Local_File.Contains(gCode) AndAlso d.Local_File.Contains(gShortName) Then
                    uf.frmUFam.ProgressBar1.Visible = True
                    uf.frmUFam.LblAction.Visible = True
                    uf.FTP_DescargarYDescomprimir(Me, d, uf.frmUFam.LblAction, uf.frmUFam.ProgressBar1)
                    uf.frmUFam.LblAction.Text = "Action:"
                    uf.frmUFam.ProgressBar1.Value = 0
                    ' Log de la Familia descargada.
                    If uf.yo Is Nothing Then uf.yo = New clsBase(Reflection.Assembly.GetExecutingAssembly)
                    uf.yo.PonLog_ULMA(ULMALGFree.ACTION.DOWNLOAD_GROUP, UPDATE_GROUP:=gCode, UPDATE_FILES:=d.Local_File)
                    '
                    Exit For
                End If
            Next
            Me.Action = queAction.updated
            uf.frmUFam.ProgressBar1.Visible = False
            uf.frmUFam.LblAction.Visible = False
            uf.frmUFam.Update()
            uf._recargarBrowser = True
            uf.cambiosEnGrupos = True
        End If
    End Sub

    Private Sub UpdateGroupFromPictureBox()
        Dim result As DialogResult = MsgBox("Do you want to update " + gButton.Text + "?", MsgBoxStyle.OkCancel, "Product Update Confirmation")
        If result = DialogResult.OK Then
            ' Si hay actualizaciones de los XML, descargar antes.
            If uf._ActualizarXMLs = True Then 'nuf.cUp("XML").Count > 0
                Dim fiXmlFtp As String = uf.cIni.IniGet(uf._IniUpdaterFull, "UPDATES", "XML")
                Dim d As New ULMALGFree.Datos("XML", fiXmlFtp)
                uf.FTP_DescargarYDescomprimirXML(d)
                '
                ' Hay que volver a leer los XML nuevos
                uf.LlenaDatosMercados()
                '
                ' Borrar los .RFA que no se esten utilizando en otros grupos.
                BorraFamiliasGroup_UnGrupo()
            End If
            '
            For Each d As ULMALGFree.Datos In uf.LUp
                If d.Local_File.Contains(gCode) AndAlso d.Local_File.Contains(gShortName) Then
                    uf.frmUFam.ProgressBar1.Visible = True
                    uf.frmUFam.LblAction.Visible = True
                    uf.FTP_DescargarYDescomprimir(Me, d, uf.frmUFam.LblAction, uf.frmUFam.ProgressBar1)
                    uf.frmUFam.LblAction.Text = "Action:"
                    uf.frmUFam.ProgressBar1.Value = 0
                    ' Log de la Familia descargada.
                    If uf.yo Is Nothing Then uf.yo = New clsBase(Reflection.Assembly.GetExecutingAssembly)
                    'uf.yo.PonLog_ULMA(ULMALGFree.ACTION.UPDATE_FAMILIES,,,,, gCode, d.Local_File)
                    '
                    Exit For
                End If
            Next
            Me.Action = queAction.toupdate
            uf.frmUFam.ProgressBar1.Visible = False
            uf.frmUFam.LblAction.Visible = False
            uf.frmUFam.Update()
            uf._recargarBrowser = True
            uf.cambiosEnGrupos = True
        End If
    End Sub
End Class
