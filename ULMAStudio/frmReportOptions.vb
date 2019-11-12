Imports System.Windows.Forms

Public Class frmReportOptions


    Private Sub frmReportOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Application.VisualStyleState = VisualStyles.VisualStyleState.ClientAndNonClientAreasEnabled
        ' View / Project
        If onlyactiveview Then
            RbView.Checked = True
        Else
            RbProject.Checked = True
        End If
        ' Ulma / All
        If onlyulma Then
            RbUlma.Checked = True
        Else
            RbAll.Checked = True
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If cIni Is Nothing Then cIni = New clsINI
        onlyactiveview = RbView.Checked
        cIni.IniWrite(ULMALGFree.clsBase._IniFull, "OPTIONS", "onlyactiveview", IIf(onlyactiveview = True, "1", "0").ToString)
        onlyulma = RbUlma.Checked
        cIni.IniWrite(ULMALGFree.clsBase._IniFull, "OPTIONS", "onlyulma", IIf(onlyulma = True, "1", "0").ToString)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class
