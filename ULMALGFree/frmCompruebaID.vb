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
        ULMALGFree.clsBase.idform = Me.TxtActivacionCode.Text
        'Dim res  ULMALGFree.validacion = uf.ID_Comprueba_OffLine
        uf.resultado = New UCClientWebService.Models.ResponseID
        Dim networkinternet As String = uf.EstadoRed_String
        If networkinternet <> "" Then
            uf.resultado.id = Me.TxtActivacionCode.Text
            uf.resultado.message = networkinternet
            uf.resultado.valid = False
            MsgBox(uf.resultado.message, MsgBoxStyle.Critical, "Registration")
            cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.resultado.id, NOTES:="Form Code Error: " & uf.resultado.message)
            Me.TxtActivacionCode.Text = ""
        Else
            uf.resultado = uf.ID_Comprueba_OnLine
            If uf.resultado.valid = True Then
                MsgBox("Ulma Studio successfully activated.", MsgBoxStyle.Information, "Registration")
                cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.resultado.id, NOTES:="Form Code OK: " & uf.resultado.message)
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                MsgBox(uf.resultado.message, MsgBoxStyle.Critical, "Registration")
                cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.resultado.id, NOTES:="Form Code Error: " & uf.resultado.message)
                Me.TxtActivacionCode.Text = ""
                uf.resultado.id = ""
                uf.resultado.message = ""
                uf.resultado.valid = False
            End If
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        uf.resultado = New UCClientWebService.Models.ResponseID
        uf.resultado.id = ""
        uf.resultado.message = "ULMA Studio not activated."
        uf.resultado.valid = False
        ULMALGFree.clsBase.idform = ""
        MsgBox(uf.resultado.message, MsgBoxStyle.Critical, "Registration")
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
        Dim target As String = "mailto:" & lblEmail.Text & "?subject=I request information about activation code of Ulma Studio" '& My.Application.Info.Version.ToString
        ' Navigate to a URL.
        System.Diagnostics.Process.Start(target)
    End Sub
End Class
