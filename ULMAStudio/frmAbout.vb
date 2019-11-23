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
        uf.Version_Put()
        Dim version As String = My.Application.Info.Version.ToString
        ' *****Quitamos esto, que el año de la versión sea el real. No makeado
        'Dim partes() As String = version.Split("."c)
        'partes(0) = uf.oVersion.RevitVersionNumber
        'Me.Text = "About v" & Join(partes, "."c)       
        'Me.LblVersion.Text = "ULMA Studio - v." & Join(partes, "."c)
        ' ********************************************************************
        Me.Text = "About v" & version
        Me.LblVersion.Text = "ULMA Studio - v." & version
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

    'Private Sub frmAbout_MouseClick(sender As Object, e As MouseEventArgs) Handles Cancel_Button.MouseClick ', Me.MouseClick, Pbox_Latest.MouseClick
    '    'Dim img As Object = UCRevitFree.My.Resources.ResourceManager.GetObject("mnuReport")
    '    abre = False
    '    Me.oT.Start()
    'End Sub
    ''
    Private Sub BtnUpdateAddIn_Click(sender As Object, e As EventArgs) Handles BtnUpdateAddIn.Click
        'Dim img As Object = UCRevitFree.My.Resources.ResourceManager.GetObject("mnuReport")
        'If BtnUpdateAddIn.BackgroundImage.Equals(UCRevitFree.My.Resources.Resources.mnuDownloadGris) Then
        If ULMALGFree.clsBase._ActualizarAddIns = False Then
            abre = False
            Me.oT.Start()
        Else
            Dim msg As String = "Do you want to update ULMA Studio? Revit will be closed and relaunched."

            If MsgBox(msg, MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Update ULMA Studio") = MsgBoxResult.Yes Then
                ' Comprobar si hay ficheros sin guardar.
                Dim haysinguardar As Boolean = False
                For Each oD As Document In evRevit.evApp.Documents
                    If oD.IsModified = True Then
                        haysinguardar = True
                        Exit For
                    End If
                Next
                If haysinguardar = True Then
                    If MsgBox("There are unsaved documents, do you want to save them?",
                              MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Update ULMA Studio") = MsgBoxResult.Yes Then
                        For Each oD As Document In evRevit.evApp.Documents
                            If oD.IsModified = True Then
                                Try
                                    oD.Save()
                                Catch ex As Exception
                                    Dim SaveFileDialog1 As SaveFileDialog = New SaveFileDialog()
                                    Dim pathFile As String
                                    If oD.IsFamilyDocument Then
                                        SaveFileDialog1.Filter = "RFA Files (*.rfa*)|*.rfa|RFT Files (*.rft*)|*.rft|RVT Files (*.rvt*)|*.rvt|RTE Files (*.rte*)|*.rte|All Files (*.*)|*.*"
                                    Else
                                        SaveFileDialog1.Filter = "RVT Files (*.rvt*)|*.rvt|RTE Files (*.rte*)|*.rte|RFA Files (*.rfa*)|*.rfa|RFT Files (*.rft*)|*.rft|All Files (*.*)|*.*"
                                    End If
                                    If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                                        pathFile = SaveFileDialog1.FileName
                                        oD.SaveAs(pathFile)
                                        Continue For
                                    Else
                                        MsgBox("It is not possible to update ULMA Studio. Please save documents before updating.", MsgBoxStyle.Exclamation, "ULMA Studio Update Aborted")
                                        Exit Sub
                                    End If
                                End Try
                            End If
                        Next
                    End If
                End If
                ' *******
                Pbox_New.Visible = False
                pbActualiza.Visible = True : pbActualiza.Value = 0
                'Try
                '    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UPDATE_ADDIN, NOTES:=uf.cUp("addins").First.ToString.Split("="c)(1))
                'Catch ex As Exception
                'End Try
                ' Crear el directorio destino, si no existe. Para que lo podamos descargar ahí.
                If IO.Directory.Exists(uf._updatesFolder) = False Then
                    Try
                        IO.Directory.CreateDirectory(uf._updatesFolder)
                    Catch ex As Exception
                        MsgBox("Error creating updates folder", MsgBoxStyle.Critical, "ATTENTION")
                        abre = False
                        Me.oT.Start()
                    End Try
                End If
                Dim FullPathZip As String = IO.Path.Combine(uf._updatesFolder, uf.cUp("addins").First)
                If IO.File.Exists(FullPathZip) = False Then
                    ' Descargar el fichero si no está en la carpeta Updates
                    ULMALGFree.clsBase.DescargaFicheroFTPUCRevitFree(ULMALGFree.FOLDERWEB.Addins, uf._updatesFolder, IO.Path.GetFileName(FullPathZip), Me.pbActualiza)
                End If
                If IO.File.Exists(FullPathZip) Then
                    Try
                        If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UPDATE_ADDIN, NOTES:=IO.Path.GetFileName(FullPathZip))
                    Catch ex As Exception
                    End Try
                End If
                ULMALGFree.clsBase.Bat_CreaEjecuta()
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
        abre = False
        Me.oT.Start()
    End Sub
End Class