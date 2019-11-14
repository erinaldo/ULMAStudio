using System;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;


namespace UCBrowser
{ 

    public class Main_viewmodel : INotifyPropertyChanged
    {
        private string _tituloDeLaVentana;
        public string tituloDeLaVentana
        {
            get { return _tituloDeLaVentana; }
            set
            {
                _tituloDeLaVentana = value;
                NotifyPropertyChanged("tituloDeLaVentana");
            }
        }

        private BibliotecaDeFamilias biblioteca;

        private Opciones opciones;
        private Favoritos favoritos;

        private ProcesadorDeComandosRevit procesadorDeComandosRevit;
        private Autodesk.Revit.UI.ExternalEvent lanzarProcesadorDeComandosRevit;

        public Main_viewmodel(ProcesadorDeComandosRevit procesadorDeComandosRevit, 
                              Autodesk.Revit.UI.ExternalEvent lanzarProcesadorDeComandosRevit)
        {
            this.procesadorDeComandosRevit = procesadorDeComandosRevit;
            procesadorDeComandosRevit.comandoAEjecutar = ProcesadorDeComandosRevit.ComandosDisponibles.noHacerNada;
            this.lanzarProcesadorDeComandosRevit = lanzarProcesadorDeComandosRevit;
            InicializarDatos();
        }
        internal void InicializarDatos()
        {
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                _lineasDeProducto = new List<LineaDeProducto>();
                NotifyPropertyChanged("lineasDeProducto");
                _grupos = new List<GrupoDeFamilias>();
                NotifyPropertyChanged("grupos");
                _familias = new List<Familia>();
                NotifyPropertyChanged("familias");

                opciones = Opciones.getOpcionesAlmacenadas(mostrarAvisoEnCasoDeError: true);

                tituloDeLaVentana = "Browser";
                                    //         [" + opciones.modoDeObtenerLosDatosBDI + "] [" + opciones.idFilialConLaQueTrabajar
                                    //+ "] [" + opciones.idiomaParaLosNombres + "]   "
                                    //+ "   (v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString()
                                    //+ "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString()
                                    //+ " c" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString() + ") ";

                biblioteca = new BibliotecaDeFamilias(idCompaniaBaan: opciones.idFilialConLaQueTrabajar, 
                                                      modoDeObtenerLosDatosBDI: opciones.modoDeObtenerLosDatosBDI,
                                                      pathDeLaCarpetaConLosXMLParaOffline: opciones.pathDeLaCarpetaConLosXMLParaOffline,
                                                      idiomaParaLosNombres: opciones.idiomaParaLosNombres,
                                                      pathDeLaCarpetaBaseDeArchivosDeFamilia: opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia,
                                                      pathDeLaCarpetaBaseDeImagenesThumbnail: opciones.pathDeLaCarpetaBaseDeImagenesThumbnail,
                                                      pathDeLaCarpetaPersonalDeArchivosDeFamilia: opciones.pathDeLaCarpetaPersonalDeArchivosDeFamilia,
                                                      pathDeLaCarpetaPersonalDeImagenesThumbnail: opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail);
                NotifyPropertyChanged("familias");

                if (opciones.usarNombresCortosEnLosGrupos)
                { 
                    biblioteca.UsarNombresCortosEnLosGrupos(); 
                }
                else
                {
                    biblioteca.UsarNombresNormalesEnLosGrupos();
                }
                if (opciones.usarNombresDeArchivoEnLasFamilias) 
                { 
                    biblioteca.UsarNombresDeArchivoEnLasFamilias(); 
                }
                else
                {
                    biblioteca.UsarNombresNormalesEnLasFamilias();
                }

                _lineasDeProducto = biblioteca.getLineasDeProducto();
                NotifyPropertyChanged("lineasDeProducto");

                //if (BibliotecaDeFamilias.familiasimg == null) { BibliotecaDeFamilias.familiasimg = new List<Familia>(); };
                //BibliotecaDeFamilias.familiasimg.Clear();
                //for (int x = 0; x < _familias.Count - 1; x++)
                //{
                //    if (_familias[x].thumbnail != null && _familias[x].thumbnail != Properties.Resources._default)
                //    {
                //        BibliotecaDeFamilias.familiasimg.Add(_familias[x]);
                //    }
                //}
                //_familias.Clear();
                //_familias = BibliotecaDeFamilias.familiasimg;
                //BibliotecaDeFamilias.familiasimg.Clear();
                //NotifyPropertyChanged("familiasimg");

                //Main_window.mWindows.actualizar();

                favoritos = new Favoritos();
                favoritos.RecuperarDelDisco(opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail, opciones.pathDeLaCarpetaBaseDeImagenesThumbnail);

                filtroAll = true;
                filtroDyn = false;
                filtroSET = false;
                filtroUnit = false;
                filtroAnnSymb = false;
                filtroANN_ = false;
                filtroDET_ = false;

                if (opciones.orientacionParaLaListaDeGrupos == "V")
                {
                    anguloDeRotacionTextBlockGrupos = -90;
                }
                else
                {
                    anguloDeRotacionTextBlockGrupos = 0;
                }

            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }
        
        public Main_viewmodel() { InicializarDatos(); } // Este constructor auxiliar, sin parametros, es solo para TEST; no usarlo para nada más.


        private List<LineaDeProducto> _lineasDeProducto;
        public List<LineaDeProducto> lineasDeProducto { get { return _lineasDeProducto; } }
        private string _lineaDeProductoSeleccionada;
        public string lineaDeProductoSeleccionada
        {
            get { return _lineaDeProductoSeleccionada; }
            set
            {
                _lineaDeProductoSeleccionada = value;
                NotifyPropertyChanged("lineaDeProductoSeleccionada");

                _grupos = biblioteca.getGruposDeLaLinea(lineaDeProductoSeleccionada);
                NotifyPropertyChanged("grupos");

                _familias = new List<Familia>();
                NotifyPropertyChanged("familias");
            }
        }

        private List<GrupoDeFamilias> _grupos;
        public List<GrupoDeFamilias> grupos { get { return _grupos; } }
        private string _grupoSeleccionado;
        public string grupoSeleccionado
        {
            get { return _grupoSeleccionado; }
            set
            {
                _grupoSeleccionado = value;
                NotifyPropertyChanged("grupoSeleccionado");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }

        private double _anguloDeRotacionTextBlockGrupos;
        public double anguloDeRotacionTextBlockGrupos
        {
            get { return _anguloDeRotacionTextBlockGrupos; }
            set
            {
                _anguloDeRotacionTextBlockGrupos = value;
                NotifyPropertyChanged("anguloDeRotacionTextBlockGrupos");
            }
        }

        public double? limiteDeTamainoElementoGrupo
        { 
            get
            {
                double limite;
                if (double.TryParse(opciones.limiteDeTamainoElementoGrupo, out limite))
                { return limite; }
                else
                { return double.NaN; }
            }
        }

        private List<Familia> _familias;
        public List<Familia> familiasimg;
        public List<Familia> familias { get { return _familias; } }
        private Familia _familiaSeleccionada;
        public Familia familiaSeleccionada
        {
            get { return _familiaSeleccionada; }
            set
            {
                _familiaSeleccionada = value;
                NotifyPropertyChanged("familiaSeleccionada");
            }
        }

        public void InsertarFamiliaSeleccionada()
        {
            InsertarFamilia(familiaSeleccionada);
        }
        public void InsertarFamilia(Familia familia)
        {
            procesadorDeComandosRevit.comandoAEjecutar = ProcesadorDeComandosRevit.ComandosDisponibles.InsertarFamilia;
            if (File.Exists(Path.Combine(opciones.pathDeLaCarpetaPersonalDeArchivosDeFamilia, familia.nombreArchivo)))
            {
                procesadorDeComandosRevit.pathArchivoFamilia = Path.Combine(opciones.pathDeLaCarpetaPersonalDeArchivosDeFamilia, familia.nombreArchivo);
            }
            else
            {
                procesadorDeComandosRevit.pathArchivoFamilia = Path.Combine(opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia, familia.nombreArchivo);
            }
            lanzarProcesadorDeComandosRevit.Raise();
        }
        //nota: no olvidar mirar su funcion siamesa que inserta familias desde ResultadosBusqueda_viewmodel.


        private ICommand _MostrarImagenBigDeLaFamiliaSeleccionada;
        private bool _sePermiteMostrarImagenBigDeLaFamiliaSeleccionada = true;
        public ICommand MostrarImagenBigDeLaFamiliaSeleccionada
        {
            get { return _MostrarImagenBigDeLaFamiliaSeleccionada ?? (_MostrarImagenBigDeLaFamiliaSeleccionada = new GestorDeComandos(() => Accion_MostrarImagenBigDeLaFamiliaSeleccionada(), _sePermiteMostrarImagenBigDeLaFamiliaSeleccionada)); }
        }
        public void Accion_MostrarImagenBigDeLaFamiliaSeleccionada()
        {
            MostrarImagenBigDeLaFamilia(familiaSeleccionada);
        }
        public void MostrarImagenBigDeLaFamilia(Familia familia)
        {
            procesadorDeComandosRevit.comandoAEjecutar = ProcesadorDeComandosRevit.ComandosDisponibles.MostrarImagenBigDeLaFamilia;
            string archivoImagenBIG = System.IO.Path.GetFileNameWithoutExtension(familia.nombreArchivo) + "_BIG" + ".png";
            if (File.Exists(Path.Combine(opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail, archivoImagenBIG)))
            {
                procesadorDeComandosRevit.pathArchivoImagenBIG = Path.Combine(opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail, archivoImagenBIG);
            }
            else
            {
                procesadorDeComandosRevit.pathArchivoImagenBIG = Path.Combine(opciones.pathDeLaCarpetaBaseDeImagenesThumbnail, archivoImagenBIG);
            }
            lanzarProcesadorDeComandosRevit.Raise();
        }


        private string _patronABuscar;
        public string patronABuscar
        {
            get { return _patronABuscar; }
            set
            {
                _patronABuscar = value;
                NotifyPropertyChanged("familiaABuscar");
            }
        }
        private ICommand _BuscarFamilias;
        private bool _sePermiteBuscarFamilias = true;
        public ICommand BuscarFamilias
        {
            get { return _BuscarFamilias ?? (_BuscarFamilias = new GestorDeComandos(() => Accion_BuscarFamilias(), _sePermiteBuscarFamilias)); }
        }
        internal void Accion_BuscarFamilias()
        {
            if (Main.cLcsv != null && patronABuscar != null)
            {
                Main.cLcsv.PonLog_ULMA("BROWSER_SEARCH", NOTES:patronABuscar, EApp: ULMALGFree.queApp.UCBROWSER);
            }
            List<Familia> familiasEncontradas = biblioteca.getFamiliasCuyoNombreContenga(patronABuscar);
            ResultadosBusqueda_window ventana = new ResultadosBusqueda_window();
            ventana.DataContext = new ResultadosBusqueda_viewmodel(familiasEncontradas, procesadorDeComandosRevit, lanzarProcesadorDeComandosRevit);
            ventana.Title = "\"..." + patronABuscar.ToUpper() + "...\"";
            ventana.ShowDialog();
        }

        private bool? _filtroAll;
        public bool? filtroAll
        {
            get { return _filtroAll; }
            set
            {
                _filtroAll = value;
                NotifyPropertyChanged("filtroAll");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private bool? _filtroSET;
        public bool? filtroSET
        {
            get { return _filtroSET; }
            set
            {
                _filtroSET = value;
                NotifyPropertyChanged("filtroSET");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private bool? _filtroAnnSymb;
        public bool? filtroAnnSymb
        {
            get { return _filtroAnnSymb; }
            set
            {
                _filtroAnnSymb = value;
                NotifyPropertyChanged("filtroAnn");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private bool? _filtroDyn;
        public bool? filtroDyn
        {
            get { return _filtroDyn; }
            set
            {
                _filtroDyn = value;
                NotifyPropertyChanged("filtroDyn");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private bool? _filtroUnit;
        public bool? filtroUnit
        {
            get { return _filtroUnit; }
            set
            {
                _filtroUnit = value;
                NotifyPropertyChanged("filtroUnit");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private bool? _filtroANN_;
        public bool? filtroANN_
        {
            get { return _filtroANN_; }
            set
            {
                _filtroANN_ = value;
                NotifyPropertyChanged("filtroANN_");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private bool? _filtroDET_;
        public bool? filtroDET_
        {
            get { return _filtroDET_; }
            set
            {
                _filtroDET_ = value;
                NotifyPropertyChanged("filtroDET_");

                _familias = biblioteca.getFamiliasDelGrupo(grupoSeleccionado, filtroSeleccionado);
                if (Main.cLcsv != null && grupoSeleccionado != null)
                {
                    Main.cLcsv.PonLog_ULMA("BROWSER_NAVIGATE", NOTES: grupoSeleccionado, EApp: ULMALGFree.queApp.UCBROWSER);
                }
                NotifyPropertyChanged("familias");
            }
        }
        private BibliotecaDeFamilias.filtroFamilia filtroSeleccionado
        {
            get
            {
                if (filtroAll == true) { return BibliotecaDeFamilias.filtroFamilia.All; }
                if (filtroSET == true) { return BibliotecaDeFamilias.filtroFamilia.SET; }
                if (filtroDyn == true) { return BibliotecaDeFamilias.filtroFamilia.Dyn; }
                if (filtroUnit == true) { return BibliotecaDeFamilias.filtroFamilia.Unit; }
                if (filtroAnnSymb == true) { return BibliotecaDeFamilias.filtroFamilia.AnnSymb; }
                if (filtroANN_ == true) { return BibliotecaDeFamilias.filtroFamilia.ANN_; }
                if (filtroDET_ == true) { return BibliotecaDeFamilias.filtroFamilia.DET_; }
                throw new InvalidEnumArgumentException("UCBrowser_viewmodel.filtroSeleccionado");
            }
        }


        public void AgregarAFavoritos(Familia familia)
        {
            favoritos.familiasFavoritas.Add(familia);
            favoritos.GuardarEnDisco();
        }
        public void QuitarDeFavoritos(Familia familia)
        {
            favoritos.familiasFavoritas.Remove(familia);
            favoritos.GuardarEnDisco();
            Accion_MostrarFamiliasFavoritas();
        }

        private ICommand _AgregarAFavoritosLaFamiliaSeleccionada;
        private bool _sePermiteAgregarAFavoritosLaFamiliaSeleccionada = true;
        public ICommand AgregarAFavoritosLaFamiliaSeleccionada
        {
            get { return _AgregarAFavoritosLaFamiliaSeleccionada ?? (_AgregarAFavoritosLaFamiliaSeleccionada = new GestorDeComandos(() => Accion_AgregarAFavoritosLaFamiliaSeleccionada(), _sePermiteAgregarAFavoritosLaFamiliaSeleccionada)); }
        }
        private void Accion_AgregarAFavoritosLaFamiliaSeleccionada()
        {
            AgregarAFavoritos(familiaSeleccionada);
        }

        private ICommand _QuitarDeFavoritosLaFamiliaSeleccionada;
        private bool _sePermiteQuitarDeFavoritosLaFamiliaSeleccionada = true;
        public ICommand QuitarDeFavoritosLaFamiliaSeleccionada
        {
            get { return _QuitarDeFavoritosLaFamiliaSeleccionada ?? (_QuitarDeFavoritosLaFamiliaSeleccionada = new GestorDeComandos(() => Accion_QuitarDeFavoritosLaFamiliaSeleccionada(), _sePermiteQuitarDeFavoritosLaFamiliaSeleccionada)); }
        }
        private void Accion_QuitarDeFavoritosLaFamiliaSeleccionada()
        {
            QuitarDeFavoritos(familiaSeleccionada);
        }

        private ICommand _MostrarFamiliasFavoritas;
        private bool _sePermiteMostrarFamiliasFavoritas = true;
        public ICommand MostrarFamiliasFavoritas
        {
            get { return _MostrarFamiliasFavoritas ?? (_MostrarFamiliasFavoritas = new GestorDeComandos(() => Accion_MostrarFamiliasFavoritas(), _sePermiteMostrarFamiliasFavoritas)); }
        }
        private void Accion_MostrarFamiliasFavoritas()
        {
            filtroAll = true;
            grupoSeleccionado = null;
            _familias = favoritos.familiasFavoritas;
            NotifyPropertyChanged("familias");
        }


        private ICommand _ActivarEditorDeOpciones;
        private bool _sePermiteActivarEditorDeOpciones = true;
        public ICommand ActivarEditorDeOpciones
        {
            get { return _ActivarEditorDeOpciones ?? (_ActivarEditorDeOpciones = new GestorDeComandos(() => Accion_ActivarEditorDeOpciones(), _sePermiteActivarEditorDeOpciones)); }
        }
        private void Accion_ActivarEditorDeOpciones()
        {
            Opciones_window ventanaOpciones = new Opciones_window();
            ventanaOpciones.DataContext = new Opciones_viewmodel(enlaceConLaVentanaUCBrowser: this);
            ventanaOpciones.ShowDialog();
        }



        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

    }





}