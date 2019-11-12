Imports System.Xml


Public Class Tarifa

    Public CodTarifa As String
    Public NombreTarifa As String
    Public CodMoneda As String


    Public Sub New()
        CodTarifa = ""
        NombreTarifa = ""
        CodMoneda = ""
    End Sub



    Public Const zzXML_NombreArchivo As String = "fees"
    Public Const zzXML_EtiquetaRaiz As String = "fees"
    Public Const zzXML_EtiquetaNodo As String = "fee"

    Public Const zzEtiquetaCodTarifa As String = "feeCode"
    Public Const zzEtiquetaCodMoneda As String = "currency"

    Public Const zzXML_EtiquetaRaizArticulos As String = "articles"
    Public Const zzXML_EtiquetaNodoArticulo As String = "article"

    Public Const zzEtiquetaCodigoArticulo As String = "articleCode"
    Public Const zzEtiquetaPrecio As String = "price"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode
        Dim nTarifa As XmlNode
        nTarifa = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

            Dim nCodigo As XmlNode
            nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodTarifa)
            Dim numero As Integer
            If Integer.TryParse(CodTarifa, numero) Then
                nCodigo.InnerText = numero.ToString("000")
            Else
                nCodigo.InnerText = CodTarifa
            End If
            nTarifa.AppendChild(nCodigo)

            Dim nNombre As XmlNode
            nNombre = documentoDondeSeVaAInsertar.CreateElement(ConstantesComunesParaXML.zzEtiquetaDescripcion)
            nNombre.InnerText = NombreTarifa
            nTarifa.AppendChild(nNombre)

            Dim nMoneda As XmlNode
            nMoneda = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodMoneda)
            nMoneda.InnerText = CodMoneda
            nTarifa.AppendChild(nMoneda)

        Return nTarifa
    End Function


    Sub fromXML(ByVal LeerTarifaDesde As XmlNode)
        If LeerTarifaDesde.Name = zzXML_EtiquetaNodo Then
            Dim nodo As XmlNode
            For Each nodo In LeerTarifaDesde.ChildNodes
                Select Case nodo.Name

                  Case zzEtiquetaCodTarifa
                      CodTarifa = nodo.InnerText

                  Case ConstantesComunesParaXML.zzEtiquetaDescripcion
                      NombreTarifa = nodo.InnerText

                  Case zzEtiquetaCodMoneda
                      CodMoneda = nodo.InnerText

                End Select
            Next
        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerTarifaDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If
    End Sub



End Class
