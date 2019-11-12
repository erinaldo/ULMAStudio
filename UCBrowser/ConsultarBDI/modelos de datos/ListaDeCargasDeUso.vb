Imports System.Xml
Imports Magnitudes


Public Class ListaDeCargasDeUso


    Public ReadOnly CodNorma As String
    Public ReadOnly cargas As List(Of DatoDeCargasDeUso)


    Public Sub New(codNorma As String, cargas As List(Of DatoDeCargasDeUso))
        Me.cargas = cargas
        Me.CodNorma = codNorma
    End Sub

    Public Overrides Function ToString() As String
        Return CodNorma
    End Function


    Public Function getSublista(tipoElemento As Elemento.TiposDeElemento, codElemento As String) As ListaDeCargasDeUso
        Dim cargasFiltradas As New List(Of DatoDeCargasDeUso)
        Dim carga As DatoDeCargasDeUso
        For Each carga In cargas
            If carga.TipoElemento = tipoElemento And carga.CodElemento = codElemento Then
                cargasFiltradas.Add(carga)
            End If
        Next
        Return New ListaDeCargasDeUso(codNorma:=CodNorma + " [" + codElemento + "]", cargas:=cargasFiltradas)
    End Function

    Public Function getCargas(tipoElemento As Elemento.TiposDeElemento, codElemento As String, _
                              orientacion As DatoDeCargasDeUso.OrientacionDelTuboInterior, alturaDeUso As Magnitud) As DatoDeCargasDeUso

        Dim alturaInf As Double = 0
        Dim alturaSup As Double = Double.MaxValue

        Dim cargausoSup As Double = 0
        Dim cargausoInf As Double = 0
        Dim cargahormSup As Double = 0
        Dim cargahormInf As Double = 0
        Dim cargareapSup As Double = 0
        Dim cargareapInf As Double = 0
        Dim rigidezSup As Double = 0
        Dim rigidezInf As Double = 0

        Dim alturaBuscada As Double = alturaDeUso.getCopiaConvertidaA("m").valor

        Dim carga As New DatoDeCargasDeUso
        For Each carga In getSublista(tipoElemento, codElemento).cargas
            If carga.TuboInterior = orientacion Or carga.TuboInterior = DatoDeCargasDeUso.OrientacionDelTuboInterior.indistinto Then
                If carga.Altura.valor <= alturaBuscada And carga.Altura.valor >= alturaInf Then
                    alturaInf = carga.Altura.valor
                    cargausoInf = carga.Carga.valor
                    cargahormInf = carga.CargaPlantaHormigonada.valor
                    cargareapInf = carga.CargaReapuntalado.valor
                    rigidezInf = carga.Rigidez.valor
                End If
                If carga.Altura.valor >= alturaBuscada And carga.Altura.valor <= alturaSup Then
                    alturaSup = carga.Altura.valor
                    cargausoSup = carga.Carga.valor
                    cargahormSup = carga.CargaPlantaHormigonada.valor
                    cargareapSup = carga.CargaReapuntalado.valor
                    rigidezSup = carga.Rigidez.valor
                End If
            End If
        Next

        If alturaSup <> Double.MaxValue And alturaInf <> 0 Then
            Dim cargaResultado As New DatoDeCargasDeUso(cargaDeLaQueCopiar:=carga)
            If alturaSup = alturaInf Then
                cargaResultado.Altura = New Magnitud(alturaBuscada, "m")
                cargaResultado.Carga = New Magnitud(cargausoInf, carga.Carga.unidaddemedida)
                cargaResultado.CargaPlantaHormigonada = New Magnitud(cargahormInf, carga.CargaPlantaHormigonada.unidaddemedida)
                cargaResultado.CargaReapuntalado = New Magnitud(cargareapInf, carga.CargaReapuntalado.unidaddemedida)
                cargaResultado.Rigidez = New Magnitud(rigidezInf, carga.Rigidez.unidaddemedida)
            Else
                cargaResultado.Altura = New Magnitud(alturaBuscada, "m")
                cargaResultado.Carga = New Magnitud(cargausoInf + ((cargausoSup - cargausoInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), carga.Carga.unidaddemedida)
                cargaResultado.CargaPlantaHormigonada = New Magnitud(cargahormInf + ((cargahormSup - cargahormInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), carga.CargaPlantaHormigonada.unidaddemedida)
                cargaResultado.CargaReapuntalado = New Magnitud(cargareapInf + ((cargareapSup - cargareapInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), carga.CargaReapuntalado.unidaddemedida)
                cargaResultado.Rigidez = New Magnitud(rigidezInf + ((rigidezSup - rigidezInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), carga.Rigidez.unidaddemedida)
            End If
            Return cargaResultado
        Else
            Return New DatoDeCargasDeUso()
        End If

    End Function

    Public Function getSublista(altura As Magnitud) As ListaDeCargasDeUso

        Dim alturaBuscada As Double
        alturaBuscada = altura.getCopiaConvertidaA("m").valor

        Dim resultados As New List(Of DatoDeCargasDeUso)

        Dim claveAnterior As String
        Dim cargaResultado As DatoDeCargasDeUso

        Dim alturaInf As Double = 0
        Dim alturaSup As Double = Double.MaxValue

        Dim cargausoSup As Double = 0
        Dim cargausoInf As Double = 0
        Dim cargahormSup As Double = 0
        Dim cargahormInf As Double = 0
        Dim cargareapSup As Double = 0
        Dim cargareapInf As Double = 0
        Dim rigidezSup As Double = 0
        Dim rigidezInf As Double = 0

        If cargas.Count > 0 Then

            ''Inicializar para el primer apeo
            claveAnterior = cargas(0).TipoElemento.ToString() & cargas(0).CodElemento.ToString() & cargas(0).TuboInterior.ToString()
            cargaResultado = New DatoDeCargasDeUso(cargaDeLaQueCopiar:=cargas(0))

            ''Procesar los apeos (incluido el primero, excluido el último)
            Dim cargaEnProceso As New DatoDeCargasDeUso
            For Each cargaEnProceso In cargas

                Dim claveEnCurso As String
                claveEnCurso = cargaEnProceso.TipoElemento.ToString() & cargaEnProceso.CodElemento.ToString() & cargaEnProceso.TuboInterior.ToString()

                If claveEnCurso <> claveAnterior And alturaSup <> Double.MaxValue And alturaInf <> 0 Then
                    If alturaSup = alturaInf Then
                        cargaResultado.Altura = New Magnitud(alturaBuscada, "m")
                        cargaResultado.Carga = New Magnitud(cargausoInf, cargaEnProceso.Carga.unidaddemedida)
                        cargaResultado.CargaPlantaHormigonada = New Magnitud(cargahormInf, cargaEnProceso.CargaPlantaHormigonada.unidaddemedida)
                        cargaResultado.CargaReapuntalado = New Magnitud(cargareapInf, cargaEnProceso.CargaReapuntalado.unidaddemedida)
                        cargaResultado.Rigidez = New Magnitud(rigidezInf, cargaEnProceso.Rigidez.unidaddemedida)
                    Else
                        cargaResultado.Altura = New Magnitud(alturaBuscada, "m")
                        cargaResultado.Carga = New Magnitud(cargausoInf + ((cargausoSup - cargausoInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.Carga.unidaddemedida)
                        cargaResultado.CargaPlantaHormigonada = New Magnitud(cargahormInf + ((cargahormSup - cargahormInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.CargaPlantaHormigonada.unidaddemedida)
                        cargaResultado.CargaReapuntalado = New Magnitud(cargareapInf + ((cargareapSup - cargareapInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.CargaReapuntalado.unidaddemedida)
                        cargaResultado.Rigidez = New Magnitud(rigidezInf + ((rigidezSup - rigidezInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.Rigidez.unidaddemedida)
                    End If
                    resultados.Add(cargaResultado)

                    claveAnterior = claveEnCurso
                    cargaResultado = New DatoDeCargasDeUso(cargaDeLaQueCopiar:=cargaEnProceso)
                    alturaInf = 0
                    alturaSup = Double.MaxValue
                    cargausoInf = 0
                    cargausoSup = 0
                    cargahormInf = 0
                    cargahormSup = 0
                    cargareapInf = 0
                    cargareapSup = 0
                    rigidezInf = 0
                    rigidezSup = 0
                End If

                If cargaEnProceso.Altura.valor <= alturaBuscada And cargaEnProceso.Altura.valor >= alturaInf Then
                    alturaInf = cargaEnProceso.Altura.valor
                    cargausoInf = cargaEnProceso.Carga.valor
                    cargahormInf = cargaEnProceso.CargaPlantaHormigonada.valor
                    cargareapInf = cargaEnProceso.CargaReapuntalado.valor
                    rigidezInf = cargaEnProceso.Rigidez.valor
                End If
                If cargaEnProceso.Altura.valor >= alturaBuscada And cargaEnProceso.Altura.valor <= alturaSup Then
                    alturaSup = cargaEnProceso.Altura.valor
                    cargausoSup = cargaEnProceso.Carga.valor
                    cargahormSup = cargaEnProceso.CargaPlantaHormigonada.valor
                    cargareapSup = cargaEnProceso.CargaReapuntalado.valor
                    rigidezSup = cargaEnProceso.Rigidez.valor
                End If

            Next

            ''Procesar el último apeo.
            If alturaSup <> Double.MaxValue And alturaInf <> 0 Then
                cargaResultado.Altura = New Magnitud(alturaBuscada, "m")
                cargaResultado.Carga = New Magnitud(cargausoInf + ((cargausoSup - cargausoInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.Carga.unidaddemedida)
                cargaResultado.CargaPlantaHormigonada = New Magnitud(cargahormInf + ((cargahormSup - cargahormInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.CargaPlantaHormigonada.unidaddemedida)
                cargaResultado.CargaReapuntalado = New Magnitud(cargareapInf + ((cargareapSup - cargareapInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.CargaReapuntalado.unidaddemedida)
                cargaResultado.Rigidez = New Magnitud(rigidezInf + ((rigidezSup - rigidezInf) * (alturaBuscada - alturaInf) / (alturaSup - alturaInf)), cargaEnProceso.Rigidez.unidaddemedida)
                resultados.Add(cargaResultado)
            End If

        End If

        Return New ListaDeCargasDeUso(codNorma:=CodNorma + " [" + altura.ToString + "]", cargas:=resultados)

    End Function



    Public Const zzXML_NombreArchivo As String = "workingloads"
    Public Const zzXML_EtiquetaRaiz As String = "workingloads"
    Public Const zzXML_EtiquetaNodo As String = "standard"

    Public Const zzXML_EtiquetaCodigo As String = "standardCode"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument) As XmlNode

        Dim nLista As XmlNode
        nLista = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

            Dim aCodigo As XmlAttribute
            aCodigo = documentoDondeSeVaAInsertar.CreateAttribute(zzXML_EtiquetaCodigo)
            aCodigo.InnerText = CodNorma.ToString
            nLista.Attributes.Append(aCodigo)

            Dim carga As DatoDeCargasDeUso
            For Each carga In cargas
                nLista.AppendChild(carga.toXML(documentoDondeSeVaAInsertar, omitirDatosVaciosONoValidos:=False))
            Next

        Return nLista

    End Function


    Sub New(ByVal LeerLaListaDesde As XmlNode)
        If LeerLaListaDesde.Name = zzXML_EtiquetaNodo Then

            Dim nCodigo As XmlNode
            nCodigo = LeerLaListaDesde.Attributes.GetNamedItem(zzXML_EtiquetaCodigo)
            If Not IsNothing(nCodigo) Then
                CodNorma = nCodigo.InnerText
            End If

            cargas = New List(Of DatoDeCargasDeUso)
            Dim nCarga As XmlNode
            For Each nCarga In LeerLaListaDesde.SelectNodes(DatoDeCargasDeUso.zzXML_EtiquetaNodo)
                Dim carga As DatoDeCargasDeUso
                carga = New DatoDeCargasDeUso
                carga.fromXML(LeerCargaDeUsoDesde:=nCarga)
                cargas.Add(carga)
            Next

        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerLaListaDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If
    End Sub


End Class
