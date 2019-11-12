Public Class frmUpdater
    Public SuperGroups As List(Of SuperGroup)

    Private Sub FrmUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        frmIni = Me
        pb1.Image = My.Resources.ResourceManager.GetObject("ORMA_render")
        SuperGroups = New List(Of SuperGroup)
        _imgBase = IO.Path.GetFullPath(_imgBase)
        If IO.File.Exists(_imgBase) = False Then
            _imgBase = IO.Path.ChangeExtension(_imgBase, "png")
        End If
    End Sub
    Private Sub btnRellena_Click(sender As Object, e As EventArgs) Handles btnRellena.Click
        SuperGroups.Clear()
        fpGroups.Controls.Clear()
        fpSuperGroups.Controls.Clear()
        pb1.Image = Image.FromFile(_imgBase)
        If Me.PanelSep1.Visible = True Then Me.PanelSep1.Visible = False
        If Me.fpGroups.Visible = True Then Me.fpGroups.Visible = False
        lblInf.Visible = True
        'Me.InitializeComponent()
        LeeXML_publictructure()
        If SuperGroups Is Nothing OrElse SuperGroups.Count = 0 Then Exit Sub
        Me.fpSuperGroups.Controls.Clear()
        For Each sGroup As SuperGroup In SuperGroups
            Me.fpSuperGroups.Controls.Add(sGroup.SgPanel)
            'sGroup.PonBotonesGrupos()
        Next
        'Dim cSp As New SuperGroup("EV", "Vertical formwork", "VERT.", My.Resources.wall_column_formwork_icon)
        'Me.fpSuperGroup.Controls.Add(cSp.spPanel)
        ''
        'cSp = New SuperGroup("AND", "Scaffolding", "SCAFF.", My.Resources.scaffolding_icon)
        'Me.fpSuperGroup.Controls.Add(cSp.spPanel)
    End Sub

    Private Sub BtnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        Me.Close()
    End Sub

    Public Sub LeeXML_publictructure()
        Dim xdoc As System.Xml.Linq.XDocument = Nothing
        Dim publicStructure As XElement = Nothing
        Try
            xdoc = XDocument.Load(_xmlPath)
        Catch ex As Exception
            MsgBox("Not exist xml document " & _xmlPath)
            Exit Sub
        End Try
        '
        Try
            publicStructure = xdoc.Element("publicStructure")
            '
            For Each Sgs In publicStructure.Elements("superGroup")
                'Dim rfc As String = v.att.Attribute("XXX").Value
                Dim SuperGroupCode As String = Sgs.Element("superGroupCode").Value
                Dim superGroupDescription As String = Sgs.Element("superGroupDescription").Value
                Dim superGroupShortName As String = Sgs.Element("superGroupShortName").Value
                Dim superGroupIcon As String = Sgs.Element("superGroupIcon").Value
                'Dim superGroupIconSelected As String = v.Element("superGroupIconSelected").Value
                Dim imgPath As String = IO.Path.Combine(_imgPath, superGroupIcon)
                imgPath = IO.Path.GetFullPath(imgPath)
                Dim imgPath1 As String = IO.Path.ChangeExtension(imgPath, IIf(imgPath.EndsWith("jpg"), "png", "jpg"))
                '
                Dim img As Image = Nothing
                If IO.File.Exists(imgPath) Then
                    img = System.Drawing.Image.FromFile(imgPath)
                ElseIf IO.File.Exists(imgPath1) Then
                    img = System.Drawing.Image.FromFile(imgPath1)
                End If
                Dim cSp As New SuperGroup(Me.fpSuperGroups, Me.fpGroups, Me.pb1, SuperGroupCode, superGroupDescription, superGroupShortName, img)
                '
                ' Read groups
                For Each Gs In Sgs.Elements("groups")
                    For Each G In Gs.Elements("group")
                        Dim groupCode As String = G.Element("groupCode").Value
                        Dim groupDescription As String = G.Element("groupDescription").Value
                        Dim groupShortName As String = G.Element("groupShortName").Value
                        Dim groupImage As String = G.Element("groupImage").Value
                        Dim imgP As String = IO.Path.Combine(_imgPath, groupImage)
                        imgP = IO.Path.GetFullPath(imgP)
                        Dim imgP1 As String = IO.Path.ChangeExtension(imgP, IIf(imgP.EndsWith("jpg"), "png", "jpg"))
                        '
                        Dim imgG As Image = Nothing
                        If IO.File.Exists(imgP) Then
                            imgG = System.Drawing.Image.FromFile(imgP)
                        ElseIf IO.File.Exists(imgP1) Then
                            imgG = System.Drawing.Image.FromFile(imgP1)
                        End If
                        Dim cG As New group(groupCode, groupDescription, groupShortName, imgG)
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
End Class
