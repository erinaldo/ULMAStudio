Imports System.Xml


Public Class Norma

    Public Codigo As String
    Public NombreLocal As String
    Public NombreOficialIngles As String


    Public Sub New()
        Codigo = ""
        NombreLocal = ""
        NombreOficialIngles = ""
    End Sub


    Public Const zzXML_NombreArchivo As String = "standards"
    Public Const zzXML_EtiquetaRaiz As String = "standards"
    Public Const zzXML_EtiquetaNodo As String = "standard"

    Public Const zzXML_EtiquetaCodigo As String = "standardCode"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode

        Dim nNorma As XmlNode
        nNorma = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

            Dim nCodigo As XmlNode
            nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaCodigo)
            nCodigo.InnerText = Codigo.ToString
            nNorma.AppendChild(nCodigo)

            Dim nDescripcionLocal As XmlNode
            nDescripcionLocal = documentoDondeSeVaAInsertar.CreateElement(ConstantesComunesParaXML.zzEtiquetaDescripcion)
            nDescripcionLocal.InnerText = NombreLocal
            Dim aIdiomaLocal As XmlAttribute
            aIdiomaLocal = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdiomaLocal.InnerText = "local"
            nDescripcionLocal.Attributes.Append(aIdiomaLocal)
            nNorma.AppendChild(nDescripcionLocal)

            Dim nDescripcionEN As XmlNode
            nDescripcionEN = documentoDondeSeVaAInsertar.CreateElement(ConstantesComunesParaXML.zzEtiquetaDescripcion)
            nDescripcionEN.InnerText = NombreOficialIngles
            Dim aIdiomaEN As XmlAttribute
            aIdiomaEN = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdiomaEN.InnerText = "en"
            nDescripcionEN.Attributes.Append(aIdiomaEN)
            nNorma.AppendChild(nDescripcionEN)

        Return nNorma

    End Function


    Sub fromXML(ByVal LeerNormaDesde As XmlNode)
        If LeerNormaDesde.Name = zzXML_EtiquetaNodo Then

            Dim nCodigo As XmlNode
            nCodigo = LeerNormaDesde.SelectSingleNode(zzXML_EtiquetaCodigo)
            If Not IsNothing(nCodigo) Then
                Codigo = nCodigo.InnerText
            End If


            Dim nNombreLocal As XmlNode
            nNombreLocal = LeerNormaDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & _
                                                    "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                    "='local']")
            If Not IsNothing(nNombreLocal) Then
                NombreLocal = nNombreLocal.InnerText
            End If


            Dim nNombreOficialIngles As XmlNode
            nNombreOficialIngles = LeerNormaDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & _
                                                            "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                            "='en']")
            If Not IsNothing(nNombreOficialIngles) Then
                NombreOficialIngles = nNombreOficialIngles.InnerText
            End If

        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerNormaDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If
    End Sub


End Class
