Imports System.Windows.Forms

Public Class frmHayUpdates
    Public txtFamilies As String = ""
    Public txtAddins As String = ""
    Public hayFamilias As Boolean = False
    Public hayAddins As Boolean = False
    Private _Interactivo As Boolean = False

    Public Property Interactivo As Boolean
        Get
            Return _Interactivo
        End Get
        Set(value As Boolean)
            _Interactivo = value
            If value = True Then
                BtnNo_Click(Nothing, Nothing)
            End If
        End Set
    End Property

    '
    Private Sub FrmHayUpdates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmHayUpdates_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmU = Nothing
    End Sub

    Private Sub BtnYes_Click(sender As Object, e As EventArgs) Handles BtnYes.Click
        Interactivo = False
        ULMALGFree.clsBase.Bat_CreaEjecuta(cerrarRevit:=True)
        Me.DialogResult = DialogResult.Yes
        Me.Close()
    End Sub

    Private Sub BtnNo_Click(sender As Object, e As EventArgs) Handles BtnNo.Click
        If Interactivo = True Then
            ULMALGFree.clsBase.Bat_CreaEjecuta(cerrarRevit:=False)
        End If
        Me.DialogResult = DialogResult.No
        Me.Close()
    End Sub

    Public Sub ActualizarFormulario(hayFamilies As Boolean, hayAddins As Boolean)
        ' 185 / 482
        'families
        PbFamiles.Visible = hayFamilies
        LblFam.Visible = hayFamilies
        LblFamilies.Visible = hayFamilies
        'Addin"
        PbAddins.Visible = hayAddins
        LblAdd.Visible = hayAddins
        LblAddin.Visible = hayAddins
        ' Redimensionar si no hay updates
        If hayFamilies = True AndAlso hayAddins = True Then
            'Me.Height = 482
        ElseIf hayFamilies = True AndAlso hayAddins = False Then
            Me.Height = 482 - 100
        ElseIf hayFamilies = False AndAlso hayAddins = True Then
            PbAddins.Location = New Drawing.Point(PbAddins.Location.X, PbAddins.Location.Y - 100)
            LblAdd.Location = New Drawing.Point(LblAdd.Location.X, LblAdd.Location.Y - 100)
            LblAddin.Location = New Drawing.Point(LblAddin.Location.X, LblAddin.Location.Y - 100)
            Me.Height = 482 - 100
        End If
    End Sub
End Class