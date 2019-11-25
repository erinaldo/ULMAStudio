Imports System.Windows.Forms
Imports uf = ULMALGFree.clsBase

Public Class frmCompruebaID
    Dim cLcsv As ULMALGFree.clsBase
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        cLcsv = New ULMALGFree.clsBase(System.Reflection.Assembly.GetExecutingAssembly)

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        'Dim res  ULMALGFree.validacion = uf.ID_Comprueba_OffLine
        uf.RespID = New UCClientWebService.Models.ResponseID
        uf.RespID.id = Me.TxtActivacionCode.Text
        Dim networkinternet As String = uf.EstadoRed_String
        If networkinternet <> "" Then
            uf.RespID.message = networkinternet
            uf.RespID.valid = False
            MsgBox(uf.RespID.message, MsgBoxStyle.Critical, "Registration")
            cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="Form Code Error: " & uf.RespID.message)
            Me.TxtActivacionCode.Text = ""
        Else
            Dim srvId As New UCClientWebService.Services.AddInService
            Dim rId As New UCClientWebService.Models.RequestID With {.id = uf.RespID.id}
            uf.RespID = srvId.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", rId)

            If uf.RespID.valid = True Then
                MsgBox("ULMA Studio successfully activated.", MsgBoxStyle.Information, "Registration")
                cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="Form Code OK: " & uf.RespID.message)
                uf.keyfile_escribe(uf.RespID.id)
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                MsgBox(uf.RespID.message, MsgBoxStyle.Critical, "Registration")
                cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="Form Code Error: " & uf.RespID.message)
                'uf.RespID = Nothing
                Me.TxtActivacionCode.Text = ""
            End If
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        If uf.RespID Is Nothing Then uf.RespID = New UCClientWebService.Models.ResponseID
        uf.RespID.id = ""
        uf.RespID.message = "ULMA Studio not activated."
        uf.RespID.valid = False
        MsgBox(uf.RespID.message, MsgBoxStyle.Critical, "Registration")
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Dim target As String = "mailto:" & lblEmail.Text & "?subject=I request information about activation code of Ulma Studio" '& My.Application.Info.Version.ToString
        ' Navigate to a URL.
        System.Diagnostics.Process.Start(target)
    End Sub

    Private Sub LblEmail_Click(sender As Object, e As EventArgs)
        Dim target As String = "mailto:" & lblEmail.Text & "?subject=I request information about activation code of Ulma Studio" '& My.Application.Info.Version.ToString
        ' Navigate to a URL.
        System.Diagnostics.Process.Start(target)
    End Sub

    Private Sub LblEmail_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblEmail.LinkClicked
        Dim target As String = "mailto:" & lblEmail.Text & "?subject=I request information about activation code of ULMA Studio" '& My.Application.Info.Version.ToString
        ' Navigate to a URL.
        System.Diagnostics.Process.Start(target)
    End Sub
End Class
