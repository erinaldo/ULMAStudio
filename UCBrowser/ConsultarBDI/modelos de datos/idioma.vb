Imports System.Xml



Public Class Idioma

    Private m_codigo As String
    Private m_descripcion As String

    Public ReadOnly Property codigo As String
        Get
            Return m_codigo
        End Get
    End Property
    Public ReadOnly Property descripcion As String
        Get
            Return m_descripcion
        End Get
    End Property


    Public Sub New(codigo As String, descripcion As String)
        m_codigo = codigo
        m_descripcion = descripcion
    End Sub


    Public Const zzXML_NombreArchivo As String = "languages"
    Public Const zzXML_EtiquetaRaiz As String = "languages"
    Public Const zzXML_EtiquetaNodo As String = "language"

    Public Const zzEtiquetaCodigo As String = "languageCode"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode

        Dim nIdioma As XmlNode
        nIdioma = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

            Dim nCodigo As XmlNode
            nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigo)
            nCodigo.InnerText = m_codigo.ToString
            nIdioma.AppendChild(nCodigo)

            Dim nDescripcion As XmlNode
            nDescripcion = documentoDondeSeVaAInsertar.CreateElement(ConstantesComunesParaXML.zzEtiquetaDescripcion)
            nDescripcion.InnerText = m_descripcion
            nIdioma.AppendChild(nDescripcion)

        Return nIdioma

    End Function

    Sub fromXML(ByVal LeerIdiomaDesde As XmlNode)
        If LeerIdiomaDesde.Name = zzXML_EtiquetaNodo Then
            Dim nodo As XmlNode
            For Each nodo In LeerIdiomaDesde.ChildNodes
                Select Case nodo.Name

                  Case zzEtiquetaCodigo
                      m_codigo = nodo.InnerText

                  Case ConstantesComunesParaXML.zzEtiquetaDescripcion
                      m_descripcion = nodo.InnerText

                End Select
            Next
        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerIdiomaDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If
    End Sub

End Class
