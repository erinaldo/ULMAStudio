Imports System.Xml
Imports Magnitudes


Public Class EstructuraConfiguracionArticulo

    Public CodConfiguracionALaQuePertenece As String

    Public CodArticulo As String
    Public Cantidad As Magnitud
    Public CodModoSuministro As String
    Public NumOrdenParaListar As Integer


    Public Sub New()
        CodConfiguracionALaQuePertenece = ""
        CodArticulo = ""
        Cantidad = New Magnitud(valor:=0.0, unidaddemedida:=Magnitud.DESCONOCIDA)
        CodModoSuministro = ""
        NumOrdenParaListar = 0
    End Sub




    Public Const zzXML_NombreArchivo As String = "configurations"
    Public Const zzXML_EtiquetaRaizConfiguraciones As String = "configurations"

    Public Const zzXML_EtiquetaNodoConfiguracionArticulos As String = "articles"
    Public Const zzXML_EtiquetaNodoArticulo As String = "article"

    Public Const zzEtiquetaCodigoArticulo As String = "articleCode"
    Public Const zzEtiquetaCantidad As String = "quantity"
    Public Const zzEtiquetaModoSuministro As String = "supplyMethod"
    Public Const zzEtiquetaOrden As String = "order"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode
        Dim nArticulo As XmlNode
        nArticulo = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoArticulo)

            Dim nOrden As XmlNode
            nOrden = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaOrden)
            nOrden.InnerText = NumOrdenParaListar.ToString
            nArticulo.AppendChild(nOrden)

            Dim nCodigo As XmlNode
            nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoArticulo)
            nCodigo.InnerText = CodArticulo
            nArticulo.AppendChild(nCodigo)

            nArticulo.AppendChild(Cantidad.toXML(documentoDondeSeVaAInsertar, zzEtiquetaCantidad))

            Dim nModoSuministro As XmlNode
            nModoSuministro = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaModoSuministro)
            nModoSuministro.InnerText = CodModoSuministro
            nArticulo.AppendChild(nModoSuministro)

        Return nArticulo
    End Function


    Public Sub fromXML(ByVal LeerArticuloDesde As XmlNode, ByVal CodConfiguracionALaQueAsignarlo As String)
       If LeerArticuloDesde.Name = zzXML_EtiquetaNodoArticulo Then
            CodConfiguracionALaQuePertenece = CodConfiguracionALaQueAsignarlo
            Dim nodo As XmlNode
            For Each nodo In LeerArticuloDesde.ChildNodes
                Select Case nodo.Name

                   Case zzEtiquetaCodigoArticulo
                      CodArticulo = nodo.InnerText

                  Case zzEtiquetaCantidad
                      Cantidad = New Magnitud(nodo)

                  Case zzEtiquetaModoSuministro
                      CodModoSuministro = nodo.InnerText

                  Case zzEtiquetaOrden
                      Try
                          NumOrdenParaListar = Integer.Parse(nodo.InnerText)
                      Catch ex As Exception
                          NumOrdenParaListar = 0
                      End Try

                End Select
            Next
        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerArticuloDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodoArticulo)
        End If

    End Sub



End Class
