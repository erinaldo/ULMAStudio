Imports System.Xml


Public Class Compania

    Public codigo As Integer
    Public descripcion As String
    Public idioma As String
    Public pais As String


    Public Sub New()
        codigo = 0
        descripcion = ""
        idioma = ""
        pais = ""
    End Sub

    Public Overrides Function ToString() As String
        Return "[" & codigo & "] " & descripcion & " [" & idioma & "-" & pais & "]"
    End Function

    Public Const zzXML_NombreArchivo As String = "companies"
    Public Const zzXML_EtiquetaRaiz As String = "companies"
    Public Const zzXML_EtiquetaNodo As String = "company"

    Public Const zzEtiquetaCodigo As String = "companyCode"
    Public Const zzEtiquetaCodigoIdioma As String = "languageCode"
    Public Const zzEtiquetaCodigoPais As String = "countryCode"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode

        Dim nCompania As XmlNode
        nCompania = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

            Dim nCodigo As XmlNode
            nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigo)
            nCodigo.InnerText = codigo.ToString
            nCompania.AppendChild(nCodigo)

            Dim nDescripcion As XmlNode
            nDescripcion = documentoDondeSeVaAInsertar.CreateElement(ConstantesComunesParaXML.zzEtiquetaDescripcion)
            nDescripcion.InnerText = descripcion
            nCompania.AppendChild(nDescripcion)

            Dim nIdioma As XmlNode
            nIdioma = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoIdioma)
            nIdioma.InnerText = idioma
            nCompania.AppendChild(nIdioma)

            Dim nPais As XmlNode
            nPais = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoPais)
            nPais.InnerText = pais
            nCompania.AppendChild(nPais)

        Return nCompania

    End Function

    Sub fromXML(ByVal LeercompaniaDesde As XmlNode)
        If LeercompaniaDesde.Name = zzXML_EtiquetaNodo Then
            Dim nodo As XmlNode
            For Each nodo In LeercompaniaDesde.ChildNodes
                Select Case nodo.Name

                  Case zzEtiquetaCodigo
                      codigo = Integer.Parse(nodo.InnerText)

                  Case ConstantesComunesParaXML.zzEtiquetaDescripcion
                      descripcion = nodo.InnerText

                  Case zzEtiquetaCodigoIdioma
                      idioma = nodo.InnerText

                  Case zzEtiquetaCodigoPais
                      pais = nodo.InnerText

                End Select
            Next
        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeercompaniaDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If
    End Sub



End Class
