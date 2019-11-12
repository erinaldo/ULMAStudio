Imports System.Xml



Public Class EstructuraGrupoSubgrupoElemento : Implements IEquatable(Of Object) : Implements IComparable(Of EstructuraGrupoSubgrupoElemento)

    Public TipoElemento As Elemento.TiposDeElemento
    Public CodElemento As String
    Public NumOrdenParaListar As Integer
    Public CodModoSuministro As String
    Public CodGrupoAlQuePertenece As String
    Public CodSubgrupoAlQuePertenece As String

    Public Shared ModosDeSuministroDisponibles As String() = {"A", "V", "VT", "M", "E"}


    Public Sub New()
        TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
        CodElemento = ""
        NumOrdenParaListar = 0
        CodModoSuministro = ""
        CodGrupoAlQuePertenece = ""
        CodSubgrupoAlQuePertenece = ""
    End Sub



    Public Const zzXML_NombreArchivo As String = "groups_subgroups_elements"
    Public Const zzXML_EtiquetaRaiz As String = "groupsSubgroupsElements"
    Public Const zzXML_EtiquetaNodoSubgrupos As String = "subgroups"
    Public Const zzXML_EtiquetaNodoElementos As String = "elements"

    Public Const zzEtiquetaModoSuministro As String = "supplyMethod"
    Public Const zzEtiquetaNumOrden As String = "visualizationOrder"

    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode

        Dim nElemento As XmlNode

        Dim unElemento As Elemento
        Select Case TipoElemento
            Case Elemento.TiposDeElemento.articulo
                unElemento = New Articulo()
            Case Elemento.TiposDeElemento.articuloGenerico
                unElemento = New ArticuloGenerico()
            Case Elemento.TiposDeElemento.configuracion
                unElemento = New Configuracion()
            Case Elemento.TiposDeElemento.bloqueAuxiliarAutocad
                unElemento = New BloqueAuxiliarAutocad()
            Case Elemento.TiposDeElemento.familiaRevit
                unElemento = New FamiliaDinamicaRevit()
            Case Else
                unElemento = New DummyElementoDesconocido()
        End Select

        unElemento.CodElemento = CodElemento
        nElemento = unElemento.toXML(documentoDondeSeVaAInsertar, incluirDescripciones:=False, descripcionesEstiloRepcon:=False)

        Dim nModoSuministro As XmlNode
        nModoSuministro = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaModoSuministro)
        nModoSuministro.InnerText = CodModoSuministro.ToString
        nElemento.AppendChild(nModoSuministro)

        Dim nNumOrden As XmlNode
        nNumOrden = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaNumOrden)
        nNumOrden.InnerText = NumOrdenParaListar.ToString
        nElemento.AppendChild(nNumOrden)

        Return nElemento

    End Function


    Public Sub fromXML(ByVal LeerEstructuraDesde As XmlNode, ByVal CodGrupoAAsignarle As String, ByVal CodSubgrupoAAsignarle As String)

            CodGrupoAlQuePertenece = CodGrupoAAsignarle
            CodSubgrupoAlQuePertenece = CodSubgrupoAAsignarle

            Dim nCodigo As XmlNode
            Select Case LeerEstructuraDesde.Name
                Case Elemento.zzXML_EtiquetaNodoArticulo
                    nCodigo = LeerEstructuraDesde.SelectSingleNode(Elemento.zzEtiquetaCodigoArticulo)
                    TipoElemento = Elemento.TiposDeElemento.articulo
                    CodElemento = nCodigo.InnerText
                Case Elemento.zzXML_EtiquetaNodoArticuloGenerico
                    nCodigo = LeerEstructuraDesde.SelectSingleNode(Elemento.zzEtiquetaCodigoArticuloGenerico)
                    TipoElemento = Elemento.TiposDeElemento.articuloGenerico
                    CodElemento = nCodigo.InnerText
                Case Elemento.zzXML_EtiquetaNodoConfiguracion
                    nCodigo = LeerEstructuraDesde.SelectSingleNode(Elemento.zzEtiquetaCodigoConfiguracion)
                    TipoElemento = Elemento.TiposDeElemento.configuracion
                    CodElemento = nCodigo.InnerText
                Case Elemento.zzXML_EtiquetaNodoBloqueAuxiliar
                    nCodigo = LeerEstructuraDesde.SelectSingleNode(Elemento.zzEtiquetaCodigoBloqueAuxiliar)
                    TipoElemento = Elemento.TiposDeElemento.bloqueAuxiliarAutocad
                    CodElemento = nCodigo.InnerText
                Case FamiliaRevit.zzXML_EtiquetaNodo
                    nCodigo = LeerEstructuraDesde.SelectSingleNode(FamiliaRevit.zzXML_EtiquetaCodigoFamiliaRevit)
                    TipoElemento = Elemento.TiposDeElemento.familiaRevit
                    CodElemento = nCodigo.InnerText
                 Case Else
                    TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
                    CodElemento = ""
            End Select

            Dim nModoSuministro As XmlNode
            nModoSuministro = LeerEstructuraDesde.SelectSingleNode(zzEtiquetaModoSuministro)
            If Not IsNothing(nModoSuministro) Then
                CodModoSuministro = nModoSuministro.InnerText
            Else
                CodModoSuministro = ""
            End If

            Dim nNumOrden As XmlNode
            nNumOrden = LeerEstructuraDesde.SelectSingleNode(zzEtiquetaNumOrden)
            If Not IsNothing(nNumOrden) Then
                If Not Integer.TryParse(nNumOrden.InnerText, NumOrdenParaListar) Then
                    NumOrdenParaListar = 0
                End If
            Else
                NumOrdenParaListar = 0
            End If

    End Sub



    Public Function getCodigoUnificadoConGrupo() As String
        Return CodGrupoAlQuePertenece.Trim & "-" & CodSubgrupoAlQuePertenece.Trim
    End Function


    Public Overrides Function Equals(ByVal otroObjeto As Object) As Boolean Implements IEquatable(Of Object).Equals
        If otroObjeto.GetType.Equals(Me.GetType) Then
            Dim otroElemento = CType(otroObjeto, EstructuraGrupoSubgrupoElemento)
            If otroElemento.TipoElemento.Equals(Me.TipoElemento) _
            And otroElemento.CodElemento.Equals(Me.CodElemento) _
            And otroElemento.NumOrdenParaListar.Equals(Me.NumOrdenParaListar) _
            And otroElemento.CodModoSuministro.Equals(Me.CodModoSuministro) _
            And otroElemento.CodGrupoAlQuePertenece.Equals(Me.CodGrupoAlQuePertenece) _
            And otroElemento.CodSubgrupoAlQuePertenece.Equals(Me.CodSubgrupoAlQuePertenece) _
            Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Public Function CompareTo(ByVal otroElemento As EstructuraGrupoSubgrupoElemento) As Integer Implements IComparable(Of EstructuraGrupoSubgrupoElemento).CompareTo
        If otroElemento.NumOrdenParaListar = Me.NumOrdenParaListar Then
            Return 0
        ElseIf otroElemento.NumOrdenParaListar < Me.NumOrdenParaListar Then
            Return 1
        Else
            Return -1
        End If
    End Function


End Class
