Imports Magnitudes
Imports System.Xml



Public Class FamiliaRevit : Inherits Elemento : Implements IEquatable(Of Object) : Implements IComparable(Of FamiliaRevit)

    Public esDinamica As Boolean
    Public esConjunto As Boolean
    Public EsAnnotationSymbol As Boolean
    Public nombreFichero As String


    Public Sub New()
        TipoElemento = TiposDeElemento.familiaRevit
        Me.Reset()
    End Sub

    Public Sub New(familiaDeLaQueCopiar As FamiliaRevit)
        TipoElemento = TiposDeElemento.familiaRevit
        Me.CodElemento = familiaDeLaQueCopiar.CodElemento
        Me.esDinamica = familiaDeLaQueCopiar.esDinamica
        Me.esConjunto = familiaDeLaQueCopiar.esConjunto
        Me.EsAnnotationSymbol = familiaDeLaQueCopiar.EsAnnotationSymbol
        Me.nombreFichero = familiaDeLaQueCopiar.nombreFichero
    End Sub

    Private Sub Reset()
        CodElemento = ""
        esDinamica = False
        esConjunto = False
        EsAnnotationSymbol = False
        nombreFichero = ""
    End Sub



    Public Const zzXML_NombreArchivo As String = "revitfamilies"
    Public Const zzXML_EtiquetaRaiz As String = "revitFamilies"
    Public Const zzXML_EtiquetaNodo As String = "family"
    Public Const zzXML_EtiquetaCodigoFamiliaRevit As String = "revitfamilyCode"

    Public Const zzXML_CodElemento As String = "ITEM_CODE"
    Public Const zzXML_esDinamica As String = "IsDynamic"
    Public Const zzXML_esConjunto As String = "IsASet"
    Public Const zzXML_esAnotacion As String = "IsAnAnnotationSymbol"
    Public Const zzXML_nombreFichero As String = "fileName"


    Public Overloads Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument, incluirDescripciones As Boolean) As XmlNode
        Dim nFamilia As XmlNode
        nFamilia = MyBase.toXML(documentoDondeSeVaAInsertar, incluirDescripciones:=incluirDescripciones, descripcionesEstiloRepcon:=False)

        Dim nDinamica As XmlNode
        nDinamica = documentoDondeSeVaAInsertar.CreateElement(zzXML_esDinamica)
        nDinamica.InnerText = esDinamica.ToString()
        nFamilia.AppendChild(nDinamica)

        Dim nConjunto As XmlNode
        nConjunto = documentoDondeSeVaAInsertar.CreateElement(zzXML_esConjunto)
        nConjunto.InnerText = esConjunto.ToString()
        nFamilia.AppendChild(nConjunto)

        Dim nAnotacion As XmlNode
        nAnotacion = documentoDondeSeVaAInsertar.CreateElement(zzXML_esAnotacion)
        nAnotacion.InnerText = EsAnnotationSymbol.ToString()
        nFamilia.AppendChild(nAnotacion)

        Dim nNombreFichero As XmlNode
        nNombreFichero = documentoDondeSeVaAInsertar.CreateElement(zzXML_nombreFichero)
        nNombreFichero.InnerText = nombreFichero
        nFamilia.AppendChild(nNombreFichero)

        Return nFamilia
    End Function

    Public Overloads Sub fromXML(ByVal LeerElementoDesde As XmlNode)

        If LeerElementoDesde.Name = zzXML_EtiquetaNodo Then

            Me.Reset()

            MyBase.fromXML(LeerElementoDesde)

            Dim nDinamica As XmlNode
            nDinamica = LeerElementoDesde.SelectSingleNode(zzXML_esDinamica)
            If Not IsNothing(nDinamica) Then
                esDinamica = Boolean.Parse(nDinamica.InnerText)
            End If

            Dim nConjunto As XmlNode
            nConjunto = LeerElementoDesde.SelectSingleNode(zzXML_esConjunto)
            If Not IsNothing(nConjunto) Then
                esConjunto = Boolean.Parse(nConjunto.InnerText)
            End If

            Dim nAnotacion As XmlNode
            nAnotacion = LeerElementoDesde.SelectSingleNode(zzXML_esAnotacion)
            If Not IsNothing(nAnotacion) Then
                EsAnnotationSymbol = Boolean.Parse(nAnotacion.InnerText)
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
            Dim otroElemento = CType(otroObjeto, FamiliaRevit)
            If otroElemento.CodElemento.Equals(Me.CodElemento) _
            And otroElemento.esDinamica.Equals(Me.esDinamica) _
            And otroElemento.esConjunto.Equals(Me.esConjunto) _
            And otroElemento.EsAnnotationSymbol.Equals(Me.EsAnnotationSymbol) _
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


    Public Overloads Function CompareTo(ByVal otraFamilia As FamiliaRevit) As Integer Implements IComparable(Of FamiliaRevit).CompareTo
        If otraFamilia.Equals(Me) Then
            Return 0
        ElseIf otraFamilia.CodElemento < Me.CodElemento Then
            Return 1
        Else
            Return -1
        End If
    End Function


End Class
