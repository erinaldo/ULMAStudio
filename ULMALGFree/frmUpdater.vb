Imports System.Windows.Forms
Imports System.Drawing
Imports uf = ULMALGFree.clsBase


Public Class frmUpdater
    Public UltimoButtonSuperGrupo As Button = Nothing
    Public UltimoButtonGrupo As Button = Nothing
    Public SuperGroups As List(Of ULMALGFree.SuperGroup)

    Private Sub frmUpdater_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        uf.frmUFam = Nothing
        If (uf.cambiosEnGrupos) Then
            'MsgBox("Browser will be closed to apply changes.", MsgBoxStyle.Information, "Information")
            UCBrowser.WindowManager.CloseWindows()
            uf.cambiosEnGrupos = False
        End If
        'Cursor.Position = New System.Drawing.Point(Cursor.Position.X + 1, Cursor.Position.Y)
    End Sub

    Public Sub FrmUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'System.Windows.Forms.Application.VisualStyleState = VisualStyles.VisualStyleState.NoneEnabled
        'Me.ProgressBar1.ForeColor = Color.FromArgb(128, 128, 255)
        Me.ProgressBar1.ForeColor = Color.Blue
        Me.ProgressBar1.Update()
        Control.CheckForIllegalCrossThreadCalls = False
        'If uf.LUp Is Nothing OrElse uf.LUp.Count = 0 OrElse uf.CLast Is Nothing OrElse uf.CLast.Count = 0 Then
        Call ULMALGFree.clsBase.INIUpdates_LeeTODO()
        'End If
        'uf.frmUFam = Me
        'PbImagen.Image = Image.FromFile(IO.Path.Combine(ULMALGFree.clsBase._imgFolder, "render_ORMA.jpg"))
        SuperGroups = New List(Of ULMALGFree.SuperGroup)
        uf._imgBasePath = IO.Path.GetFullPath(uf._imgBasePath)
        If IO.File.Exists(uf._imgBasePath) = False Then
            uf._imgBasePath = IO.Path.ChangeExtension(uf._imgBasePath, "png")
        End If
        If IO.File.Exists(uf._imgBasePath) = True Then
            uf._imgBase = Image.FromFile(uf._imgBasePath)
        End If
        btnRellena_Click(Nothing, Nothing)
        ' Ver cuantas actualizaciones tiene el supergrupo
        For Each oSg As ULMALGFree.SuperGroup In SuperGroups
            Dim comparador As ComparadorGrupos = New ComparadorGrupos()
            oSg.groups.Sort(comparador)

            Dim contador As Integer = 0
            For Each oG As ULMALGFree.Group In oSg.groups
                Dim oAction As ULMALGFree.queAction = oG.Grupo_DameAction()
                '2019/11/20 Xabier Calvo: Mostrar botón rojo solo para actualizacaciones (no incluir productos no descargados)
                If oAction = ULMALGFree.queAction.notupdated Then
                    contador += 1
                End If
            Next
            If contador > 0 Then
                oSg.nActualizaciones = contador
            End If
        Next
        '2019/10/29 Xabier Calvo: Mostrar la ventana Download con el primer Supergrupo/grupo seleccionado
        TryCast(fpSuperGroups.Controls(0).Controls(0), Button).PerformClick()
        TryCast(fpGroups.Controls(0).Controls(0), Button).PerformClick()
    End Sub
    Private Sub btnRellena_Click(sender As Object, e As EventArgs)
        SuperGroups.Clear()
        fpGroups.Controls.Clear()
        fpSuperGroups.Controls.Clear()
        PbImagen.Image = Image.FromFile(uf._imgBasePath)
        PbImagen.SizeMode = PictureBoxSizeMode.AutoSize

        'PbImagen.Image.
        'If Me.fpGroups.Visible = True Then Me.fpGroups.Visible = False
        'Me.InitializeComponent()
        LeeXML_publictructure()
        If SuperGroups Is Nothing OrElse SuperGroups.Count = 0 Then Exit Sub
        '2019/10/28 Orden alfabetico (igual que el browser)
        Dim comparador As ComparadorSuperGrupos = New ComparadorSuperGrupos()
        SuperGroups.Sort(comparador)
        Me.fpSuperGroups.Controls.Clear()
        For Each sGroup As ULMALGFree.SuperGroup In SuperGroups
            Me.fpSuperGroups.Controls.Add(sGroup.SgPanel)
            'sGroup.PonBotonesGrupos()
        Next
        'Dim cSp As New SuperGroup("EV", "Vertical formwork", "VERT.", My.Resources.wall_column_formwork_icon)
        'Me.fpSuperGroup.Controls.Add(cSp.spPanel)
        ''
        'cSp = New SuperGroup("AND", "Scaffolding", "SCAFF.", My.Resources.scaffolding_icon)
        'Me.fpSuperGroup.Controls.Add(cSp.spPanel)

    End Sub

    Private Sub BtnCerrar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Public Sub LeeXML_publictructure()
        Dim xdoc As System.Xml.Linq.XDocument = Nothing
        Dim publicStructure As XElement = Nothing
        Try
            xdoc = XDocument.Load(uf._xmlFull)
        Catch ex As Exception
            MsgBox("Not exist xml document " & uf._xmlFull)
            Exit Sub
        End Try
        '
        Try
            publicStructure = xdoc.Element("publicStructure")
            '
            For Each Sgs In publicStructure.Elements("superGroup")
                'Dim rfc As String = v.att.Attribute("XXX").Value
                Dim SuperGroupCode As String = Sgs.Element("superGroupCode").Value
                'Dim superGroupDescription As String = Sgs.Element("superGroupDescription").Value
                Dim superGroupDescription As String = Sgs.Descendants("superGroupDescription").Where(Function(x) x.Attribute("language").Value = "en").First.Value
                'Dim superGroupShortName As String = Sgs.Element("superGroupShortName").Value
                Dim superGroupShortName As String = Sgs.Descendants("superGroupShortName").Where(Function(x) x.Attribute("language").Value = "en").First.Value
                'Dim superGroupIcon As String = Sgs.Element("superGroupIcon").Value
                Dim superGroupIcon As String = Sgs.Element("superGroupImage").Value
                'Dim superGroupIconSelected As String = v.Element("superGroupIconSelected").Value
                Dim imgPath As String = IO.Path.Combine(uf._imgFolder, superGroupIcon)
                imgPath = IO.Path.GetFullPath(imgPath)
                Dim imgPath1 As String = IO.Path.ChangeExtension(imgPath, IIf(imgPath.EndsWith("jpg"), "png", "jpg"))
                '
                Dim img As Image = Nothing
                If IO.File.Exists(imgPath) Then
                    img = System.Drawing.Image.FromFile(imgPath)
                    Dim bm As New Bitmap(img)
                    bm.MakeTransparent()
                    img = bm
                ElseIf IO.File.Exists(imgPath1) Then
                    img = System.Drawing.Image.FromFile(imgPath1)
                End If
                Dim cSp As New ULMALGFree.SuperGroup(Me.fpSuperGroups, Me.fpGroups, Me.PbImagen, SuperGroupCode, superGroupDescription, superGroupShortName, img)
                '
                ' Read groups
                For Each Gs In Sgs.Elements("groups")
                    For Each G In Gs.Elements("group")
                        Dim groupCode As String = G.Element("groupCode").Value
                        Dim groupDescription As String = G.Element("groupDescription").Value
                        Dim groupShortName As String = G.Element("groupShortName").Value
                        Dim groupImage As String = G.Element("groupImage").Value
                        Dim imgP As String = IO.Path.Combine(uf._imgFolder, groupImage)
                        imgP = IO.Path.GetFullPath(imgP)
                        Dim imgP1 As String = IO.Path.ChangeExtension(imgP, IIf(imgP.EndsWith("jpg"), "png", "jpg"))
                        '
                        Dim imgG As Image = Nothing
                        If IO.File.Exists(imgP) Then
                            imgG = System.Drawing.Image.FromFile(imgP)
                            Dim bm As New Bitmap(imgG)
                            bm.MakeTransparent()
                            imgG = bm
                        ElseIf IO.File.Exists(imgP1) Then
                            imgG = System.Drawing.Image.FromFile(imgP1)
                        End If
                        Dim cG As New ULMALGFree.Group(groupCode, groupDescription, groupShortName, imgG)
                        cSp.groups.Add(cG)
                        cG = Nothing
                    Next
                Next
                '
                SuperGroups.Add(cSp)
                cSp = Nothing
            Next
        Catch ex As Exception
            MsgBox("Format of XML Documento is not correct...")
            Exit Sub
        End Try
    End Sub

    Public Sub cambiaFondos(ByRef panel As Panel)
        cambiaFondosGpanel(panel)
    End Sub

    Public Sub cambiaFondosSeleccionado(ByRef panel As Panel)
        cambiaFondosSeleccionadoGpanel(panel)
    End Sub

    Public Sub cambiaFondosSuperGrupo(ByRef panel As Panel)
        cambiaFondosSGpanel(panel)
    End Sub

    Public Sub cambiaFondosSuperGrupoABlanco()
        cambiaFondosSGpanelABlanco()
    End Sub
End Class
