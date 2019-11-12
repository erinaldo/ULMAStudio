Imports System.Xml



Public Class Subgrupo : Inherits Elemento

    Public CodGrupoProducto As String
    Public posicionDeVisualizacion As Integer


    Public Sub New()
        TipoElemento = TiposDeElemento.subgrupo
    End Sub



    Public Function getCodigoUnificadoConGrupo() As String
        Return getCodigoUnificadoConGrupo(Me.CodGrupoProducto, Me.CodElemento)
    End Function
    Public Shared Function getCodigoUnificadoConGrupo(CodigoDeGrupoProducto As String, CodigoDeElemento As String) As String
        Return CodigoDeGrupoProducto.Trim & "-" & CodigoDeElemento.Trim
    End Function


    Public Const zzEtiquetaPosicionDeVisualizacion As String = "visualizationOrder"


    Public Overloads Function ToXML(ByVal documentoDondeSeVaAInsertar As XmlDocument, ByVal incluirDescripciones As Boolean, descripcionesEstiloRepcon As Boolean) As XmlNode

        Dim nElemento As XmlNode
        nElemento = MyBase.toXML(documentoDondeSeVaAInsertar, incluirDescripciones, descripcionesEstiloRepcon)

        Dim nPosicionDeVisualizacion As XmlNode = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaPosicionDeVisualizacion)
        nPosicionDeVisualizacion.InnerText = posicionDeVisualizacion.ToString
        nElemento.AppendChild(nPosicionDeVisualizacion)

        Return nElemento

    End Function



    Public Overloads Sub fromXML(ByVal LeerElementoDesde As XmlNode)

        MyBase.fromXML(LeerElementoDesde)

        Dim nPosicionVisualizacion As XmlNode
        nPosicionVisualizacion = LeerElementoDesde.SelectSingleNode(zzEtiquetaPosicionDeVisualizacion)
        If Not IsNothing(nPosicionVisualizacion) Then
            Integer.TryParse(nPosicionVisualizacion.InnerText, posicionDeVisualizacion)
        End If

    End Sub


    Public Overloads Function Equals(ByVal otroSubgrupo As Subgrupo) As Boolean
        If otroSubgrupo.CodGrupoProducto = Me.CodGrupoProducto _
           And otroSubgrupo.CodElemento = Me.CodElemento Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overloads Function CompareTo(ByVal otroSubgrupo As Subgrupo) As Integer
        If otroSubgrupo.posicionDeVisualizacion = Me.posicionDeVisualizacion Then
            Return 0
        ElseIf otroSubgrupo.posicionDeVisualizacion < Me.posicionDeVisualizacion Then
            Return 1
        Else
            Return -1
        End If
    End Function


End Class
