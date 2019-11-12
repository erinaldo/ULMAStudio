Public Class clsCompanies
    Public companyCode As String = ""
    Public description As String = ""
    Public languageCode As String = ""
    Public countryCode As String = ""
    Public idiomapais As String = ""
    ''
    Public Sub New(companyC As String, descrip As String, languageC As String, countryC As String)
        Me.companyCode = companyC
        Me.description = descrip
        Me.languageCode = languageC
        Me.countryCode = countryC
        If countryCode = "" Then
            Me.idiomapais = languageCode
        Else
            Me.idiomapais = languageCode & "-" & countryCode
        End If
    End Sub
End Class
