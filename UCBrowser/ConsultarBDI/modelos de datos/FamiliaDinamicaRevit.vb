Imports Magnitudes
Imports System.Xml



Public Class FamiliaDinamicaRevit : Inherits Elemento : Implements IEquatable(Of Object) : Implements IComparable(Of FamiliaDinamicaRevit)

    Public FAMILY_CODE As String
    Public CodElementoAlQueRepresenta As String
    Public longitud As Magnitud
    Public anchura As Magnitud
    Public altura As Magnitud
    Public tipoGenerico As TiposDeGenerico
    Public nombreFichero As String

    Public Enum TiposDeGenerico
        NOgenerico = 0
        genericoLINEAL = 1
        genericoDeAREA = 2
    End Enum


    Public Sub New()
        TipoElemento = TiposDeElemento.familiaRevit
        Me.Reset()
    End Sub

    Public Sub New(familiaDeLaQueCopiar As FamiliaDinamicaRevit)
        Me.CodElementoAlQueRepresenta = familiaDeLaQueCopiar.CodElementoAlQueRepresenta
        TipoElemento = TiposDeElemento.familiaRevit
        Me.FAMILY_CODE = familiaDeLaQueCopiar.FAMILY_CODE
        Me.CodElementoAlQueRepresenta = familiaDeLaQueCopiar.CodElementoAlQueRepresenta
        Me.longitud = familiaDeLaQueCopiar.longitud
        Me.anchura = familiaDeLaQueCopiar.anchura
        Me.altura = familiaDeLaQueCopiar.altura
        Me.tipoGenerico = familiaDeLaQueCopiar.tipoGenerico
        Me.nombreFichero = familiaDeLaQueCopiar.nombreFichero
    End Sub

    Private Sub Reset()
         CodElementoAlQueRepresenta = ""
         FAMILY_CODE = ""
         CodElementoAlQueRepresenta = ""
         longitud = New Magnitud(0, Magnitud.DESCONOCIDA)
         anchura = New Magnitud(0, Magnitud.DESCONOCIDA)
         altura = New Magnitud(0, Magnitud.DESCONOCIDA)
         tipoGenerico = New TiposDeGenerico
         nombreFichero = ""
    End Sub



   Public Const zzXML_NombreArchivo As String = "revit_dynamicfamilies"
   Public Const zzXML_EtiquetaRaiz As String = "revitDynamicFamilies"
   Public Const zzXML_EtiquetaNodo As String = "family"
   Public Const zzXML_EtiquetaCodigo As String = "revitfamilyCode"
   Public Const zzXML_nombreFichero As String = "fileName"

   Public Const zzXML_FAMILY_CODE As String = "FAMILY_CODE"
   Public Const zzXML_CodElementoAlQueRepresenta As String = "ITEM_CODE"
   Public Const zzXML_longitud As String = "ITEM_LENGTH"
   Public Const zzXML_anchura As String = "ITEM_WIDTH"
   Public Const zzXML_altura As String = "ITEM_HEIGHT"
   Public Const zzXML_TipoGenerico As String = "ITEM_GENERIC"


   Public Overloads Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode
        Dim nFamilia As XmlNode
        nFamilia = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

        Dim nFAMILY_CODE As XmlNode
        nFAMILY_CODE = documentoDondeSeVaAInsertar.CreateElement(zzXML_FAMILY_CODE)
        nFAMILY_CODE.InnerText = FAMILY_CODE
        nFamilia.AppendChild(nFAMILY_CODE)

        Dim nCodigoElementoAlQueRepresenta As XmlNode
        nCodigoElementoAlQueRepresenta = documentoDondeSeVaAInsertar.CreateElement(zzXML_CodElementoAlQueRepresenta)
        nCodigoElementoAlQueRepresenta.InnerText = CodElementoAlQueRepresenta
        nFamilia.AppendChild(nCodigoElementoAlQueRepresenta)

        nFamilia.AppendChild(longitud.ToXML(documentoDondeSeVaAInsertar, zzXML_longitud))

        nFamilia.AppendChild(anchura.ToXML(documentoDondeSeVaAInsertar, zzXML_anchura))

        nFamilia.AppendChild(altura.ToXML(documentoDondeSeVaAInsertar, zzXML_altura))

        Dim nTipoGenerico As XmlNode
        nTipoGenerico = documentoDondeSeVaAInsertar.CreateElement(zzXML_TipoGenerico)
        nTipoGenerico.InnerText = tipoGenerico.GetHashCode.ToString()
        nFamilia.AppendChild(nTipoGenerico)

        Dim nCodigo As XmlNode
        nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaCodigo)
        nCodigo.InnerText = CodElemento
        nFamilia.AppendChild(nCodigo)

        Dim nNombreFichero As XmlNode
        nNombreFichero = documentoDondeSeVaAInsertar.CreateElement(zzXML_nombreFichero)
        nNombreFichero.InnerText = nombreFichero
        nFamilia.AppendChild(nNombreFichero)

        Return nFamilia
   End Function

   Public Overloads Sub fromXML(ByVal LeerElementoDesde As XmlNode)

        If LeerElementoDesde.Name = zzXML_EtiquetaNodo Then

            Me.Reset()

            Dim nFAMILY_CODE As XmlNode
            nFAMILY_CODE = LeerElementoDesde.SelectSingleNode(zzXML_FAMILY_CODE)
            If Not IsNothing(nFAMILY_CODE) Then
                FAMILY_CODE = nFAMILY_CODE.InnerText
            End If

            Dim nCodElementoAlQueRepresenta As XmlNode
            nCodElementoAlQueRepresenta = LeerElementoDesde.SelectSingleNode(zzXML_CodElementoAlQueRepresenta)
            If Not IsNothing(nCodElementoAlQueRepresenta) Then
                CodElementoAlQueRepresenta = nCodElementoAlQueRepresenta.InnerText
            End If

            Dim nLongitud As XmlNode
            nLongitud = LeerElementoDesde.SelectSingleNode(zzXML_longitud)
            If Not IsNothing(nLongitud) Then
                 longitud = New Magnitud(nLongitud)
            End If

            Dim nAnchura As XmlNode
            nAnchura = LeerElementoDesde.SelectSingleNode(zzXML_anchura)
            If Not IsNothing(nAnchura) Then
                 anchura = New Magnitud(nAnchura)
            End If

            Dim nAltura As XmlNode
            nAltura = LeerElementoDesde.SelectSingleNode(zzXML_altura)
            If Not IsNothing(nAltura) Then
                 altura = New Magnitud(nAltura)
            End If

            Dim nTipoGenerico As XmlNode
            nTipoGenerico = LeerElementoDesde.SelectSingleNode(zzXML_TipoGenerico)
            If Not IsNothing(nTipoGenerico) Then
                tipoGenerico = CType([Enum].Parse(GetType(FamiliaDinamicaRevit.TiposDeGenerico), nTipoGenerico.InnerText), FamiliaDinamicaRevit.TiposDeGenerico)
            End If

            Dim nCodigo As XmlNode
            nCodigo = LeerElementoDesde.SelectSingleNode(zzXML_EtiquetaCodigo)
            If Not IsNothing(nCodigo) Then
                CodElemento = nCodigo.InnerText
            End If

            Dim nNombreFichero As XmlNode
            nNombreFichero = LeerElementoDesde.SelectSingleNode(zzXML_nombreFichero)
            If Not IsNothing(nNombreFichero) Then
                nombreFichero = nNombreFichero.InnerText
            End If

        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerElementoDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If

    End Sub


    Public Overloads Function Equals(ByVal otroObjeto As Object) As Boolean
        If otroObjeto.GetType.Equals(Me.GetType) Then
            Dim otroElemento = CType(otroObjeto, FamiliaDinamicaRevit)
            If otroElemento.CodElemento.Equals(Me.CodElemento) _
            And otroElemento.FAMILY_CODE.Equals(Me.FAMILY_CODE) _
            And otroElemento.CodElementoAlQueRepresenta.Equals(Me.CodElementoAlQueRepresenta) _
            And otroElemento.longitud.Equals(Me.longitud) _
            And otroElemento.anchura.Equals(Me.anchura) _
            And otroElemento.altura.Equals(Me.altura) _
            And otroElemento.tipoGenerico.Equals(Me.tipoGenerico) _
            And otroElemento.nombreFichero.Equals(Me.nombreFichero) _
            Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Public Overloads Function CompareTo(ByVal otraFamiliaDinamica As FamiliaDinamicaRevit) As Integer Implements IComparable(Of FamiliaDinamicaRevit).CompareTo
        If otraFamiliaDinamica.Equals(Me) Then
            Return 0
        ElseIf otraFamiliaDinamica.CodElemento < Me.CodElemento Then
            Return 1
        Else
            Return -1
        End If
    End Function


End Class
