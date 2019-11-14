Imports System.Drawing.Text
Imports System.Windows
Imports System.Windows.Forms
Imports uf = ULMALGFree.clsBase


Public Class frmAbout
    Public WithEvents oT As Timers.Timer
    Public abre As Boolean = True
    ''
    Private Sub frmAbout_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UCR_ABOUT,,, arrM123, arrL123)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        frmA = Nothing
    End Sub
    ''
    Private Sub frmAbout_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Para quitar el aviso de subprocesos al llamar al timer
        CheckForIllegalCrossThreadCalls = False
        Dim version As String = My.Application.Info.Version.ToString
        Dim partes() As String = version.Split("."c)
        partes(0) = uf.AppRevitVersionYear
        Me.Text = "About v" & Join(partes, "."c)    'RevitVersion & " - v." & My.Application.Info.Version.ToString       
        Me.LblVersion.Text = "ULMA Studio - v." & My.Application.Info.Version.ToString
        ''
        Me.oT = New Timers.Timer
        Me.abre = True
        ''
        Me.oT.Interval = 100
        Me.oT.Start()
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UCR_ABOUT, , , arrM123, arrL123)
        If ULMALGFree.clsBase._ActualizarAddIns = False Then
            Me.BtnUpdateAddIn.BackgroundImage = Nothing
            'Me.BtnUpdateAddIn.Enabled = False
            Me.BtnUpdateAddIn.Visible = False
            Me.Pbox_New.Visible = False
            Me.Pbox_Latest.Visible = True

        Else
            Me.BtnUpdateAddIn.BackgroundImage = ULMAStudio.My.Resources.Resources.toupdate2
            Me.BtnUpdateAddIn.Visible = True
            Me.Pbox_Latest.Visible = False
            Me.Pbox_New.Visible = True
        End If
    End Sub

    Private Sub frmAbout_MouseClick(sender As Object, e As MouseEventArgs) Handles Cancel_Button.MouseClick ', Me.MouseClick, Pbox_Latest.MouseClick
        'Dim img As Object = UCRevitFree.My.Resources.ResourceManager.GetObject("mnuReport")
        abre = False
        Me.oT.Start()
    End Sub
    ''
    Private Sub BtnUpdateAddIn_Click(sender As Object, e As EventArgs) Handles BtnUpdateAddIn.Click
        'Dim img As Object = UCRevitFree.My.Resources.ResourceManager.GetObject("mnuReport")
        'If BtnUpdateAddIn.BackgroundImage.Equals(UCRevitFree.My.Resources.Resources.mnuDownloadGris) Then
        If ULMALGFree.clsBase._ActualizarAddIns = False Then
            abre = False
            Me.oT.Start()
        Else
            Dim msg As String = "Do you really want to update the addin? (Revit will be closed and reopened)"

            If MsgBox(msg, MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Update AddIn") = MsgBoxResult.Yes Then
                Try
                    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UPDATE_ADDIN,,,,,, uf.cUp("addins").First.ToString.Split("="c)(1))
                Catch ex As Exception
                End Try
                ULMALGFree.clsBase.Bat_CreaEjecuta(cerrarRevit:=True)
            End If
        End If
    End Sub

    'Private Sub LblWeb_Click(sender As Object, e As EventArgs)
    '    Dim target As String = "http://" & LblWeb.Text
    '    ' Navigate to a URL.
    '    System.Diagnostics.Process.Start(target)
    'End Sub
    Private Sub LblEmail_Click(sender As Object, e As EventArgs) Handles LblEmail.Click
        Dim target As String = "mailto:" & contact & "?subject=I request information about --> ULMA Studio" '& My.Application.Info.Version.ToString
        ' Navigate to a URL.
        System.Diagnostics.Process.Start(target)
    End Sub

    Private Sub oT_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles oT.Elapsed
        If abre = True Then
            If Me.Opacity < 1.0 Then
                Me.Opacity += 0.15
            ElseIf Me.Opacity >= 1 Then
                Me.oT.Stop()
            End If
        ElseIf abre = False Then
            If Me.Opacity > 0.25 Then
                Me.Opacity -= 0.15
            ElseIf Me.Opacity <= 0.25 Then
                Me.Close()
            End If
        End If
    End Sub

    Private Sub PBox_Web_Click(sender As Object, e As EventArgs) Handles PBox_Web.Click
        Dim webAddress As String = "http://www.ulmaconstruction.com"
        Process.Start(webAddress)
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click

    End Sub
End Class