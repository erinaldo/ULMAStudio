Imports cBDI = ConsultarBDI

Partial Public Class clsBase
    Public Shared CONSULTAS As ConsultarBDI.CONSULTAS
    'Public Shared FamiliasRevit As List(Of cBDI.FamiliaRevit)
    Public Shared oArticulos As New Dictionary(Of String, cBDI.Articulo)
    Public Shared oArticulosGenericos As New Dictionary(Of String, cBDI.ArticuloGenerico)
    Public Shared oCompanias As New List(Of cBDI.Compania)
    Public Shared oConfiguraciones As New Dictionary(Of String, cBDI.Configuracion)
    Public Shared oArticulo_Datosadicionales As New Dictionary(Of String, cBDI.ElementoDatosadicionales)
    Public Shared oConfiguraciones_Datosadicionales As New Dictionary(Of String, cBDI.ElementoDatosadicionales)
    Public Shared oArticulos_EstructuraConfiguraciones As New List(Of cBDI.EstructuraConfiguracionArticulo)
    Public Shared oElementos_EstructuraGruposSubgrupos As New List(Of cBDI.EstructuraGrupoSubgrupoElemento)
    Public Shared oFamiliasRevit As New List(Of cBDI.FamiliaRevit)
    Public Shared oFamiliasDinamicas As New List(Of cBDI.FamiliaDinamicaRevit)
    Public Shared oGrupos As New Dictionary(Of String, cBDI.Grupo)
    Public Shared oSubGrupos As New Dictionary(Of String, cBDI.Subgrupo)
    '
    Public Shared oFamilasRevitExisten As New List(Of cBDI.FamiliaRevit)

    Public Shared Sub CONSULTAS_TODO()
        Dim inicio As DateTime = DateTime.Now
        '
        oArticulos = Articulos()
        System.Threading.Thread.Sleep(0)
        oArticulosGenericos = ArticulosGenericos()
        System.Threading.Thread.Sleep(0)
        oCompanias = Companias()
        System.Threading.Thread.Sleep(0)
        oConfiguraciones = Configuraciones()
        System.Threading.Thread.Sleep(0)
        oArticulo_Datosadicionales = Articulo_Datosadicionales()
        System.Threading.Thread.Sleep(0)
        oConfiguraciones_Datosadicionales = Configuraciones_Datosadicionales()
        System.Threading.Thread.Sleep(0)
        oArticulos_EstructuraConfiguraciones = Articulos_EstructuraConfiguraciones()
        System.Threading.Thread.Sleep(0)
        oElementos_EstructuraGruposSubgrupos = Elementos_EstructuraGruposSubgrupos()    ' CodElemento, CodGrupoAlQuePertenece. Este tiene el codigo del elemento y grupo al que pertenece
        System.Threading.Thread.Sleep(0)
        oFamiliasRevit = FamiliasRevit()
        System.Threading.Thread.Sleep(0)
        oFamiliasDinamicas = FamiliasDinamicas()
        System.Threading.Thread.Sleep(0)
        oGrupos = Grupos()                  ' CodElemento (groupCode)  CodLineProducto(superGroupCode)
        System.Threading.Thread.Sleep(0)
        oSubGrupos = SubGrupos()
        System.Threading.Thread.Sleep(0)
        '
        oFamilasRevitExisten = FamiliasRevitExisten()
        System.Threading.Thread.Sleep(0)
        Dim fin As DateTime = DateTime.Now
        Dim total As Long = (fin - inicio).Ticks
        ' 1,53 segundos me ha dado.
        Debug.Print(New TimeSpan(total).TotalSeconds.ToString)
        'MsgBox(New TimeSpan(total).TotalSeconds.ToString)
    End Sub
    Public Shared Function Articulos() As Dictionary(Of String, cBDI.Articulo)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getArticulos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET), True, True)
    End Function
    Public Shared Function ArticulosGenericos() As Dictionary(Of String, cBDI.ArticuloGenerico)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getArticulosGenericos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET), True, True)
    End Function
    Public Shared Function Companias() As List(Of cBDI.Compania)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getCompanias()
    End Function
    Public Shared Function Configuraciones() As Dictionary(Of String, cBDI.Configuracion)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getConfiguraciones(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function Articulo_Datosadicionales() As Dictionary(Of String, cBDI.ElementoDatosadicionales)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getDatosAdicionalesDeArticulos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function Configuraciones_Datosadicionales() As Dictionary(Of String, cBDI.ElementoDatosadicionales)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getDatosAdicionalesDeConfiguraciones(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function Articulos_EstructuraConfiguraciones() As List(Of cBDI.EstructuraConfiguracionArticulo)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getEstructuraConfiguracionesArticulos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function Elementos_EstructuraGruposSubgrupos() As List(Of cBDI.EstructuraGrupoSubgrupoElemento)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getEstructuraGruposSubgruposElementos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function FamiliasRevit() As List(Of cBDI.FamiliaRevit)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getFamiliasRevit(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function FamiliasDinamicas() As List(Of cBDI.FamiliaDinamicaRevit)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getFamiliasDinamicasRevit(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function Grupos() As Dictionary(Of String, cBDI.Grupo)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getGrupos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function
    Public Shared Function SubGrupos() As Dictionary(Of String, cBDI.Subgrupo)
        If CONSULTAS Is Nothing Then CONSULTAS = New cBDI.CONSULTAS(cBDI.CONSULTAS.ModoDeObtenerLosDatos.offline)
        Return CONSULTAS.getSubgrupos(Convert.ToInt32(DEFAULT_PROGRAM_MARKET))
    End Function

    Public Shared Function FamiliasRevitExisten() As List(Of cBDI.FamiliaRevit)
        Dim FullFilesRFA As String() = IO.Directory.GetFiles(ULMALGFree.clsBase.path_families_base, "*.rfa", IO.SearchOption.AllDirectories)
        Dim NameFilesRFA = From x In FullFilesRFA
                           Select IO.Path.GetFileName(x)
        Dim filtradas = From x In oFamiliasRevit
                        Where NameFilesRFA.Contains(x.nombreFichero)
                        Select x
        Return filtradas.ToList
    End Function
End Class
