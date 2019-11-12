using System;
using System.IO;
using System.Collections.Generic;

using ConsultarBDI;
using System.Linq;

namespace UCBrowser
{

    public class BibliotecaDeFamilias
    {
        private List<LineaDeProducto> lineasDeProducto;
        private List<GrupoDeFamilias> grupos;
        private List<Familia> familias;
        private List<relacionFamiliaGrupo> familiasYgrupos;

        private List<Filial> filiales;
        private List<string> idiomas;
        private string idiomaParaNombres;

 
        public BibliotecaDeFamilias(int idCompaniaBaan, 
                                    ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos modoDeObtenerLosDatosBDI, 
                                    string pathDeLaCarpetaConLosXMLParaOffline,
                                    string idiomaParaLosNombres,
                                    string pathDeLaCarpetaBaseDeArchivosDeFamilia,
                                    string pathDeLaCarpetaBaseDeImagenesThumbnail,
                                    string pathDeLaCarpetaPersonalDeArchivosDeFamilia,
                                    string pathDeLaCarpetaPersonalDeImagenesThumbnail)
        {
            CONSULTAS bdi;
            if (modoDeObtenerLosDatosBDI == CONSULTAS.ModoDeObtenerLosDatos.online || string.IsNullOrWhiteSpace(pathDeLaCarpetaConLosXMLParaOffline))
            {
                bdi = new CONSULTAS(modoDeObtenerLosDatosBDI);
            }
            else
            {
                bdi = new CONSULTAS(new DirectoryInfo(pathDeLaCarpetaConLosXMLParaOffline));
            }

            this.filiales = getFiliales(bdi);
            this.idiomas = getIdiomas(bdi);
            this.idiomaParaNombres = idiomaParaLosNombres;

            this.lineasDeProducto = new List<LineaDeProducto>();
            this.grupos = new List<GrupoDeFamilias>();
            RellenarLineasDeProductoYGrupos(idCompaniaBaan, bdi);
            this.familiasYgrupos = new List<relacionFamiliaGrupo>();
            RellenarRelacionesFamiliaGrupo(idCompaniaBaan, bdi);
            
            this.familias = new List<Familia>();
            RellenarFamilias(idCompaniaBaan, bdi, pathDeLaCarpetaPersonalDeArchivosDeFamilia, pathDeLaCarpetaBaseDeArchivosDeFamilia);
            
            RellenarEstructuraDesdeLasSubcarpetasDe(pathDeLaCarpetaBaseDeArchivosDeFamilia);
            RellenarEstructuraDesdeLasSubcarpetasDe(pathDeLaCarpetaPersonalDeArchivosDeFamilia);

            PonerImagenesThumbnailALasFamilias(pathDeLaCarpetaBaseDeImagenesThumbnail, pathDeLaCarpetaPersonalDeImagenesThumbnail);

        }

 
        private static List<Filial> getFiliales(CONSULTAS bdi)
        {
            List<Compania> companiasBaan = bdi.getCompanias();
            List<Filial> lista = new List<Filial>();
            foreach (Compania compania in companiasBaan)
            {
                lista.Add(new Filial(id: compania.codigo, 
                                     nombre: "[" + compania.codigo + "] " + compania.descripcion,
                                     idioma: compania.idioma + "-" + compania.pais));
            }
            return lista;
        }

        public static List<Filial> getFilialesDirectamenteDeLaBDI(ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos modo)
        {
            List<Filial> lista;
            if (modo == CONSULTAS.ModoDeObtenerLosDatos.offline)
            {
                lista = getFiliales(new CONSULTAS(CONSULTAS.ModoDeObtenerLosDatos.offline));
            }
            else
            {
                try
                {
                    lista = getFiliales(new CONSULTAS(CONSULTAS.ModoDeObtenerLosDatos.online));
                }
                catch
                {
                    lista = getFiliales(new CONSULTAS(CONSULTAS.ModoDeObtenerLosDatos.offline));
                }
            }
            if (lista == null)
            { lista = new List<Filial>(); }
            return lista;
        }
 
        private static List<string> getIdiomas(CONSULTAS bdi)
        {
            List<Idioma> idiomasBaan = bdi.getIdiomas();
            List<string> lista = new List<string>();
            foreach (Idioma idioma in idiomasBaan)
            {
                lista.Add(idioma.codigo);
            }
            return lista;
        }

        public static List<string> getIdiomasDirectamenteDeLaBDI(ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos modo)
        {
            List<string> lista;
            if (modo == CONSULTAS.ModoDeObtenerLosDatos.offline)
            {
                lista = getIdiomas(new CONSULTAS(CONSULTAS.ModoDeObtenerLosDatos.offline));
            }
            else
            {
                try
                {
                    lista = getIdiomas(new CONSULTAS(CONSULTAS.ModoDeObtenerLosDatos.online));
                }
                catch
                {
                    lista = getIdiomas(new CONSULTAS(CONSULTAS.ModoDeObtenerLosDatos.offline));
                }
            }
            if (lista == null)
            { lista = new List<string>(); }
            return lista;
        }


        private void RellenarLineasDeProductoYGrupos(int idCompaniaBaan, CONSULTAS bdi)
        {
            SortedList<string, LineaDeProducto> clasificadorDeLineas = new SortedList<string, LineaDeProducto>();
            foreach (KeyValuePair<string, Grupo> grupoBDI in bdi.getGrupos(idCompaniaBaan))
            {
                if (grupoBDI.Value.estaActivo)
                {
                    if (!clasificadorDeLineas.ContainsKey(grupoBDI.Value.CodLineaProducto))
                    {
                        clasificadorDeLineas.Add(key: grupoBDI.Value.CodLineaProducto,
                                                 value: new LineaDeProducto(id: grupoBDI.Value.CodLineaProducto, nombre: grupoBDI.Value.CodLineaProducto));
                    }

                    this.grupos.Add(new GrupoDeFamilias(id: grupoBDI.Value.CodElemento, 
                                                        nombre: grupoBDI.Value.getNombre(idiomaParaNombres),
                                                        lineaALaQuePertenece: grupoBDI.Value.CodLineaProducto,
                                                        nombreCorto: grupoBDI.Value.getAbreviatura(idiomaParaNombres)));
                }
            }

            foreach (LineaDeProducto linea in clasificadorDeLineas.Values)
            {
                lineasDeProducto.Add(linea);
            }

            grupos.Sort((x, y) => x.nombre.CompareTo(y.nombre));

        }

        private void RellenarRelacionesFamiliaGrupo(int idCompaniaBaan, CONSULTAS bdi)
        {
            List<relacionFamiliaGrupo> listaConPosiblesDuplicados = new List<relacionFamiliaGrupo>();
            foreach (EstructuraGrupoSubgrupoElemento elemento in bdi.getEstructuraGruposSubgruposElementos(idCompaniaBaan))
            {
                if (elemento.TipoElemento == Elemento.TiposDeElemento.familiaRevit
                    || elemento.TipoElemento == Elemento.TiposDeElemento.articulo)
                {
                    listaConPosiblesDuplicados.Add(new relacionFamiliaGrupo(idFamilia: elemento.CodElemento, idGrupoAlQuePertenece: elemento.CodGrupoAlQuePertenece));
                }
            }
            // Una misma familia puede estar en varios subgrupos de un mismo grupo, aqui solo necesitamos que aparezca una vez por grupo.
            familiasYgrupos = listaConPosiblesDuplicados.Distinct().ToList();
        }

        private void RellenarFamilias(int idCompaniaBaan, CONSULTAS bdi,
                                      string pathDeLaCarpetaPersonalDeArchivosDeFamilia,
                                      string pathDeLaCarpetaBaseDeArchivosDeFamilia)
        {
            List<string> listaConPosiblesDuplicados = new List<string>();
            if (System.IO.Directory.Exists(pathDeLaCarpetaPersonalDeArchivosDeFamilia))
            {
                foreach (string archivo in System.IO.Directory.GetFiles(pathDeLaCarpetaPersonalDeArchivosDeFamilia, "*.rfa"))
                {
                    listaConPosiblesDuplicados.Add(System.IO.Path.GetFileName(archivo));
                }
            }
            if (System.IO.Directory.Exists(pathDeLaCarpetaBaseDeArchivosDeFamilia))
            {
                foreach (string archivo in System.IO.Directory.GetFiles(pathDeLaCarpetaBaseDeArchivosDeFamilia, "*.rfa"))
                {
                    listaConPosiblesDuplicados.Add(System.IO.Path.GetFileName(archivo));
                }
            }
            List<string> listaDeNombresDeArchivoDeFamiliasEnLasCarpetas = new List<string>();
            listaDeNombresDeArchivoDeFamiliasEnLasCarpetas = listaConPosiblesDuplicados.Distinct().ToList();

            // en caso de no haber ningún archivo de familia en las carpetas de biblioteca configuradas, mejor no seguir adelante y dejar la lista de familias sin rellenar
            if (listaDeNombresDeArchivoDeFamiliasEnLasCarpetas.Count > 0)
            {
                foreach (FamiliaRevit familia in bdi.getFamiliasRevit(idCompaniaBaan))
                {
                    string nombre = familia.getNombre(idiomaParaNombres);
                    if (string.IsNullOrWhiteSpace(nombre)) { nombre = familia.nombreFichero; }
                    // 2019/10/10 Xabier Calvo: Opción para que se incluyan solo familias disponibles o todos los que estén en los XML de familias Revit
                    Boolean existe = false;
                    Boolean disponibles = true;
                    if (disponibles) { existe = listaConPosiblesDuplicados.Contains(familia.nombreFichero); }
                    //if (listaArticulosEnGrSgr.Contains(familia.CodElemento) && existe)
                    if (existe)
                    {
                            familias.Add(new Familia(id: familia.CodElemento,
                                             nombre: nombre, 
                                             nombreArchivo: familia.nombreFichero,
                                             esDinamica: familia.esDinamica,
                                             esConjunto: familia.esConjunto,
                                             esAnnotationSymbol: familia.EsAnnotationSymbol));
                    }
                    // **************************************
                }

                Dictionary<string, Articulo> listaDeArticulos = bdi.getArticulos(idCompaniaBaan, limitarASoloLosPresentesEnGrupoSubgrupo: true, limitarASoloGruposActivos: false);
                foreach (string nombreArchivo in listaDeNombresDeArchivoDeFamiliasEnLasCarpetas)
                {
                    if (!familias.Exists(x => x.nombreArchivo.Equals(nombreArchivo)))
                    {
                        KeyValuePair<string, Articulo> articuloRelacionado = listaDeArticulos.FirstOrDefault(x => nombreArchivo.Contains(x.Key));
                        if(articuloRelacionado.Key == null)
                        {
                                familias.Add(new Familia(id: System.IO.Path.GetFileNameWithoutExtension(nombreArchivo),
                                                     nombre: System.IO.Path.GetFileNameWithoutExtension(nombreArchivo),
                                                     nombreArchivo: nombreArchivo,
                                                     esDinamica: false, esConjunto: false, esAnnotationSymbol: false));
                        }
                        else
                        {
                                familias.Add(new Familia(id: articuloRelacionado.Value.CodElemento,
                                                         nombre: articuloRelacionado.Value.getNombre(idiomaParaNombres) + " -" + articuloRelacionado.Value.CodElemento.Trim() + "-",
                                                         nombreArchivo: nombreArchivo,
                                                         esDinamica: false, esConjunto: false, esAnnotationSymbol: false));
                        }
                    }
                }

            }
        }

        private void RellenarEstructuraDesdeLasSubcarpetasDe(string pathDeLaCarpeta)
        {
            if(Directory.Exists(pathDeLaCarpeta))
            {
                DirectoryInfo carpetaMadre = new DirectoryInfo(pathDeLaCarpeta);
                foreach (DirectoryInfo carpetaHija in carpetaMadre.GetDirectories())
                {
                    lineasDeProducto.Add(new LineaDeProducto(id:carpetaHija.Name, nombre: "[" + carpetaHija.Name.Substring(startIndex:0, length:3) + "]"));
                    foreach (DirectoryInfo carpetaNieta in carpetaHija.GetDirectories())
                    {
                        string nombreCorto;
                        if (carpetaNieta.Name.Length < 8)
                        { nombreCorto = carpetaNieta.Name; }
                        else
                        { nombreCorto = carpetaNieta.Name.Substring(0,8); }
                        grupos.Add(new GrupoDeFamilias(id: carpetaNieta.Name, nombre: carpetaNieta.Name, lineaALaQuePertenece: carpetaHija.Name, nombreCorto: nombreCorto));

                        foreach (FileInfo archivo in carpetaNieta.GetFiles("*.rfa"))
                        {
                            familiasYgrupos.Add(new relacionFamiliaGrupo(idFamilia: archivo.Name, idGrupoAlQuePertenece: carpetaNieta.Name));
                            Familia familia = new Familia(id:archivo.Name, 
                                                          nombre: archivo.Name, 
                                                          nombreArchivo: archivo.FullName.Replace(oldValue: pathDeLaCarpeta, newValue: "").TrimStart(Path.DirectorySeparatorChar), 
                                                          esDinamica: false, esConjunto: false, esAnnotationSymbol: true);
                            familias.Add(familia);
                        }
                    }
                }
            }
        }

        private void PonerImagenesThumbnailALasFamilias(string pathDeLaCarpetaBaseDeImagenesThumbnail, string pathDeLaCarpetaPersonalDeImagenesThumbnail)
        {
            Dictionary<string, System.Drawing.Bitmap> imagenes = new Dictionary<string, System.Drawing.Bitmap>();
            if (System.IO.Directory.Exists(pathDeLaCarpetaPersonalDeImagenesThumbnail))
            {
                foreach (string archivo in System.IO.Directory.GetFiles(path: pathDeLaCarpetaPersonalDeImagenesThumbnail, searchPattern: "*.png"))
                {
                    // nota: el indice va con la extensión .rfa en lugar de .png porque luego habrá que ir buscando por el nombre de archivo de la familia.
                    imagenes.Add(key: Path.GetFileNameWithoutExtension(archivo) + ".rfa",
                                 value: new System.Drawing.Bitmap(archivo));
                }
            }
            if (System.IO.Directory.Exists(pathDeLaCarpetaBaseDeImagenesThumbnail))
            {
                foreach (string archivo in System.IO.Directory.GetFiles(path: pathDeLaCarpetaBaseDeImagenesThumbnail, searchPattern: "*.png"))
                {
                    if (!imagenes.ContainsKey(Path.GetFileNameWithoutExtension(archivo) + ".rfa"))
                    {
                        // nota: el indice va con la extensión .rfa en lugar de .png porque luego habrá que ir buscando por el nombre de archivo de la familia.
                        imagenes.Add(key: Path.GetFileNameWithoutExtension(archivo) + ".rfa",
                                     value: new System.Drawing.Bitmap(archivo));
                    }
                }
            }

            if (imagenes.Count > 0)
            {
                foreach (Familia familia in familias)
                {
                    System.Drawing.Bitmap imagen;
                    if (imagenes.TryGetValue(familia.nombreArchivo, out imagen))
                    {
                        familia.thumbnail = imagen;
                    }
                    else
                    {
                        familia.thumbnail = null;   // Properties.Resources._default;
                    }
                }
            }

        }


        //Chapu: Esta funcion de ObtenerImagen existe porque, en la lista de favoritos, se necesita para 
        //obtener la imagen de las familias favoritas durante la recuperacion de estas desde disco.
        //_pendiente_: Seria conveniente hacerlo de otra manera y prescindir de esta función.
        public static System.Drawing.Bitmap ObtenerImagen(string nombreArchivoFamilia,
                                                          string pathDeLaCarpetaPersonalDeImagenesThumbnail,
                                                          string pathDeLaCarpetaBaseDeImagenesThumbnail)
        {
            string archivoDeImagen = Path.Combine(pathDeLaCarpetaPersonalDeImagenesThumbnail,
                                     Path.ChangeExtension(Path.GetFileName(nombreArchivoFamilia), "png"));
            if (File.Exists(archivoDeImagen))
            {
                return new System.Drawing.Bitmap(archivoDeImagen);
            }
            else
            {
                archivoDeImagen = Path.Combine(pathDeLaCarpetaBaseDeImagenesThumbnail,
                                  Path.ChangeExtension(Path.GetFileName(nombreArchivoFamilia), "png"));
                if (File.Exists(archivoDeImagen))
                {
                    return new System.Drawing.Bitmap(archivoDeImagen);
                }
                else
                {
                    return Properties.Resources._default;
                }
            }
        }



        public List<LineaDeProducto> getLineasDeProducto()
        {
            return lineasDeProducto;
        }

        public List<GrupoDeFamilias> getGruposDeLaLinea(string idLineaDeProducto)
        {
            List<GrupoDeFamilias> resultado = new List<GrupoDeFamilias>();
            foreach (GrupoDeFamilias grupo in grupos.FindAll(x => x.lineaALaQuePertenece.Equals(idLineaDeProducto)))
            {
                resultado.Add(grupo);
            }
            return resultado;
        }

        public enum filtroFamilia
        {
            All,
            Dyn,
            SET,
            Unit,
            AnnSymb,
            ANN_,
            DET_
        }
        public List<Familia> getFamiliasDelGrupo(string idGrupo, filtroFamilia filtro)
        {
            List<Familia> resultado = new List<Familia>();
            if (idGrupo != null)
            {
                foreach (relacionFamiliaGrupo relacion in familiasYgrupos.FindAll(x => x.idGrupoAlQuePertenece.Equals(idGrupo)))
                {
                    Familia familia = familias.Find(x => x.id.Equals(relacion.idFamilia));
                    if (familia != null)
                    {
                        if (filtro == filtroFamilia.All
                            || (filtro == filtroFamilia.Dyn && familia.esDinamica)
                            || (filtro == filtroFamilia.SET && familia.esConjunto)
                            || (filtro == filtroFamilia.Unit && !familia.esAnnotationSymbol && !familia.esDinamica && !familia.esConjunto)
                            || (filtro == filtroFamilia.AnnSymb && familia.esAnnotationSymbol)
                            || (filtro == filtroFamilia.ANN_ && System.IO.Path.GetFileName(familia.nombreArchivo).StartsWith("ANN_"))
                            || (filtro == filtroFamilia.DET_ && System.IO.Path.GetFileName(familia.nombreArchivo).StartsWith("DET_"))
                           )
                        {
                            resultado.Add(familia);
                        }
                    }
                }
                resultado.Sort((x,y) => x.descripcion.CompareTo(y.descripcion));
            }
            return resultado;
        }

        public List<Familia> getFamiliasCuyoNombreContenga(string patronABuscar)
        {
            List<Familia> familiasConDuplicados = familias.FindAll(x => x.descripcion.ToUpper().Contains(patronABuscar.ToUpper()) || x.id.ToUpper().Contains(patronABuscar.ToUpper()));
            ComparadorDeFamilias comparador = new ComparadorDeFamilias();
            return familiasConDuplicados.Distinct(comparador).ToList();
        }


        public void UsarNombresCortosEnLosGrupos()
        {
            foreach (GrupoDeFamilias grupo in grupos)
            {
                if(grupo.nombreCorto == null)
                {
                    grupo.descripcion = "z";
                }
                else
                {
                    grupo.descripcion = grupo.nombreCorto;
                }
            }
            grupos.Sort((x,y) => x.descripcion.CompareTo(y.descripcion));
        }
        public void UsarNombresNormalesEnLosGrupos()
        {
            foreach (GrupoDeFamilias grupo in grupos)
            {
                grupo.descripcion = grupo.nombre;
            }
            grupos.Sort((x, y) => x.nombre.CompareTo(y.nombre));
        }

        public void UsarNombresDeArchivoEnLasFamilias()
        {
            foreach (Familia familia in familias)
            {
                if (familia.nombreArchivo == null)
                {
                    familia.descripcion = "z";
                }
                else
                {
                    familia.descripcion = System.IO.Path.GetFileName(familia.nombreArchivo);
                }
            }
        }
        public void UsarNombresNormalesEnLasFamilias()
        {
            foreach (Familia familia in familias)
            {
                familia.descripcion = familia.nombre;
            }
        }


    }

}
