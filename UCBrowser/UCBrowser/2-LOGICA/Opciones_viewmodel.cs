using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;

namespace UCBrowser
{
    public class Opciones_viewmodel : INotifyPropertyChanged
    {
        private Main_viewmodel enlaceConLaVentanaUCBrowser;
        public Opciones_viewmodel(Main_viewmodel enlaceConLaVentanaUCBrowser)
        {
            this.enlaceConLaVentanaUCBrowser = enlaceConLaVentanaUCBrowser;

            Opciones opciones = Opciones.getOpcionesAlmacenadas(mostrarAvisoEnCasoDeError: false);
            pathDeLaCarpetaConLosXMLParaOffline = opciones.pathDeLaCarpetaConLosXMLParaOffline;
            pathDeLaCarpetaBaseDeArchivosDeFamilia = opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia;
            pathDeLaCarpetaBaseDeImagenesThumbnail = opciones.pathDeLaCarpetaBaseDeImagenesThumbnail;
            pathDeLaCarpetaPersonalDeArchivosDeFamilia = opciones.pathDeLaCarpetaPersonalDeArchivosDeFamilia;
            pathDeLaCarpetaPersonalDeImagenesThumbnail = opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail;

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            try
            {
                _filiales = BibliotecaDeFamilias.getFilialesDirectamenteDeLaBDI(opciones.modoDeObtenerLosDatosBDI);
                NotifyPropertyChanged("filiales");
                filialSeleccionada = opciones.idFilialConLaQueTrabajar;

                _idiomas = new ReadOnlyObservableCollection<string>(new ObservableCollection<string>(BibliotecaDeFamilias.getIdiomasDirectamenteDeLaBDI(opciones.modoDeObtenerLosDatosBDI)));
                NotifyPropertyChanged("idiomas");
                idiomaSeleccionado = opciones.idiomaParaLosNombres;
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }

            if(opciones.modoDeObtenerLosDatosBDI == ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.online)
            {
                leerBDIonline = true;
                leerBDIoffline = false;
            }
            else
            {
                leerBDIonline = false;
                leerBDIoffline = true;
            }

            if(opciones.orientacionParaLaListaDeGrupos == "V")
            {
                gruposEnVertical = true;
                gruposEnHorizontal = false;
            }
            else
            {
                gruposEnVertical = false;
                gruposEnHorizontal = true;
            }

            double limite;
            if (double.TryParse(opciones.limiteDeTamainoElementoGrupo, out limite))
            {
                limitarTamainoElementoGrupo = true;
                limiteDeTamainoElementoGrupo = opciones.limiteDeTamainoElementoGrupo;
            }
            else
            {
                limitarTamainoElementoGrupo = false;
                limiteDeTamainoElementoGrupo = "";
            }

            usarNombresCortosEnLosGrupos = opciones.usarNombresCortosEnLosGrupos;

            if (opciones.usarNombresDeArchivoEnLasFamilias)
            {
                mostrarNombresDescriptivosEnLasFamilias = false;
                mostrarNombresDeArchivoEnLasFamilias = true;
            }
            else
            {
                mostrarNombresDescriptivosEnLasFamilias = true;
                mostrarNombresDeArchivoEnLasFamilias = false;
            }

        }

        private List<Filial> _filiales;
        public List<Filial> filiales { get { return _filiales; } }
        private int _filialSeleccionada;
        public int filialSeleccionada
        {
            get { return _filialSeleccionada; }
            set
            {
                _filialSeleccionada = value;
                NotifyPropertyChanged("filialSeleccionada");
            }
        }

        private ReadOnlyObservableCollection<string> _idiomas;
        public ReadOnlyObservableCollection<string> idiomas { get { return _idiomas; } }
        private string _idiomaSeleccionado;
        public string idiomaSeleccionado
        {
            get { return _idiomaSeleccionado; }
            set
            {
                _idiomaSeleccionado = value;
                NotifyPropertyChanged("filialIdiomaSeleccionada");
            }
        }

        private bool? _leerBDIonline;
        public bool? leerBDIonline
        {
            get { return _leerBDIonline; }
            set
            {
                _leerBDIonline = value;
                NotifyPropertyChanged("leerBDIonline");
                NotifyPropertyChanged("seVisualizaElSelectorDeLaUbicacionCarpetaConLosXMLParaOffline");
            }
        }
        private bool? _leerBDIoffline;
        public bool? leerBDIoffline
        {
            get { return _leerBDIoffline; }
            set
            {
                _leerBDIoffline = value;
                NotifyPropertyChanged("leerBDIoffline");
                NotifyPropertyChanged("seVisualizaElSelectorDeLaUbicacionCarpetaConLosXMLParaOffline");
            }
        }
        public ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos modoDeObtenerLosDatosBDI
        {
            get 
            {
                if (leerBDIonline == true)
                {
                    return ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.online;
                }
                else
                {
                    return ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.offline;
                }
            }
        }
        public System.Windows.Visibility seVisualizaElSelectorDeLaUbicacionCarpetaConLosXMLParaOffline
        {
            get 
            {
                if (leerBDIonline == true)
                {
                    return System.Windows.Visibility.Hidden;
                }
                else
                {
                    return System.Windows.Visibility.Visible;
                }
            }
        }
        private string _pathDeLaCarpetaConLosXMLParaOffline;
        public string pathDeLaCarpetaConLosXMLParaOffline
        {
            get { return _pathDeLaCarpetaConLosXMLParaOffline; }
            set
            {
                _pathDeLaCarpetaConLosXMLParaOffline = value;
                NotifyPropertyChanged("pathDeLaCarpetaConLosXMLParaOffline");
            }
        }
        private ICommand _SeleccionarCarpetaConLosXMLParaOffline;
        private bool _sePermiteSeleccionarCarpetaConLosXMLParaOffline = true;
        public ICommand SeleccionarCarpetaConLosXMLParaOffline
        {
            get { return _SeleccionarCarpetaConLosXMLParaOffline ?? (_SeleccionarCarpetaConLosXMLParaOffline = new GestorDeComandos(() => Accion_SeleccionarCarpetaConLosXMLParaOffline(), _sePermiteSeleccionarCarpetaConLosXMLParaOffline)); }
        }
        public void Accion_SeleccionarCarpetaConLosXMLParaOffline()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
                pathDeLaCarpetaConLosXMLParaOffline = dialogo.FileName;
            }
        }

        private string _pathDeLaCarpetaBaseDeArchivosDeFamilia;
        public string pathDeLaCarpetaBaseDeArchivosDeFamilia
        {
            get { return _pathDeLaCarpetaBaseDeArchivosDeFamilia; } 
            set
            {
                _pathDeLaCarpetaBaseDeArchivosDeFamilia = value;
                _pathDeLaCarpetaBaseDeImagenesThumbnail = value + "_images";
                NotifyPropertyChanged("pathDeLaCarpetaBaseDeArchivosDeFamilia");
                NotifyPropertyChanged("pathDeLaCarpetaBaseDeImagenesThumbnail");
            }
        }
        private string _pathDeLaCarpetaBaseDeImagenesThumbnail;
        public string pathDeLaCarpetaBaseDeImagenesThumbnail
        {
            get { return _pathDeLaCarpetaBaseDeImagenesThumbnail; }
            set
            {
                _pathDeLaCarpetaBaseDeImagenesThumbnail = value;
                NotifyPropertyChanged("pathDeLaCarpetaBaseDeImagenesThumbnail");
            }
        }
        private ICommand _SeleccionarCarpetaBaseDeArchivosDeFamilia;
        private bool _sePermiteSeleccionarCarpetaBaseDeArchivosDeFamilia = true;
        public ICommand SeleccionarCarpetaBaseDeArchivosDeFamilia
        {
            get { return _SeleccionarCarpetaBaseDeArchivosDeFamilia ?? (_SeleccionarCarpetaBaseDeArchivosDeFamilia = new GestorDeComandos(() => Accion_SeleccionarCarpetaBaseDeArchivosDeFamilia(), _sePermiteSeleccionarCarpetaBaseDeArchivosDeFamilia)); }
        }
        public void Accion_SeleccionarCarpetaBaseDeArchivosDeFamilia()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
               pathDeLaCarpetaBaseDeArchivosDeFamilia = dialogo.FileName;
            }
        }
        private ICommand _SeleccionarCarpetaBaseDeImagenesThumbnail;
        private bool _sePermiteSeleccionarCarpetaBaseDeImagenesThumbnail = true;
        public ICommand SeleccionarCarpetaBaseDeImagenesThumbnail
        {
            get { return _SeleccionarCarpetaBaseDeImagenesThumbnail ?? (_SeleccionarCarpetaBaseDeImagenesThumbnail = new GestorDeComandos(() => Accion_SeleccionarCarpetaBaseDeImagenesThumbnail(), _sePermiteSeleccionarCarpetaBaseDeImagenesThumbnail)); }
        }
        public void Accion_SeleccionarCarpetaBaseDeImagenesThumbnail()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
                pathDeLaCarpetaBaseDeImagenesThumbnail = dialogo.FileName;
            }
        }

        private string _pathDeLaCarpetaPersonalDeArchivosDeFamilia;
        public string pathDeLaCarpetaPersonalDeArchivosDeFamilia
        {
            get { return _pathDeLaCarpetaPersonalDeArchivosDeFamilia; }
            set
            {
                _pathDeLaCarpetaPersonalDeArchivosDeFamilia = value;
                _pathDeLaCarpetaPersonalDeImagenesThumbnail = value + "_images";
                NotifyPropertyChanged("pathDeLaCarpetaPersonalDeArchivosDeFamilia");
                NotifyPropertyChanged("pathDeLaCarpetaPersonalDeImagenesThumbnail");
            }
        }
        private string _pathDeLaCarpetaPersonalDeImagenesThumbnail;
        public string pathDeLaCarpetaPersonalDeImagenesThumbnail
        {
            get { return _pathDeLaCarpetaPersonalDeImagenesThumbnail; }
            set
            {
                _pathDeLaCarpetaPersonalDeImagenesThumbnail = value;
                NotifyPropertyChanged("pathDeLaCarpetaPersonalDeImagenesThumbnail");
            }
        }
        private ICommand _SeleccionarCarpetaPersonalDeArchivosDeFamilia;
        private bool _sePermiteSeleccionarCarpetaPersonalDeArchivosDeFamilia = true;
        public ICommand SeleccionarCarpetaPersonalDeArchivosDeFamilia
        {
            get { return _SeleccionarCarpetaPersonalDeArchivosDeFamilia ?? (_SeleccionarCarpetaPersonalDeArchivosDeFamilia = new GestorDeComandos(() => Accion_SeleccionarCarpetaPersonalDeArchivosDeFamilia(), _sePermiteSeleccionarCarpetaPersonalDeArchivosDeFamilia)); }
        }
        public void Accion_SeleccionarCarpetaPersonalDeArchivosDeFamilia()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
                pathDeLaCarpetaPersonalDeArchivosDeFamilia = dialogo.FileName;
                pathDeLaCarpetaPersonalDeImagenesThumbnail = pathDeLaCarpetaPersonalDeArchivosDeFamilia + "_images";
                Main.cIni.IniWrite(ULMALGFree.clsBase._IniFull, "PATHS", "path_families_custom", pathDeLaCarpetaPersonalDeArchivosDeFamilia);
                Main.cIni.IniWrite(ULMALGFree.clsBase._IniFull, "PATHS", "path_families_custom_images", pathDeLaCarpetaPersonalDeArchivosDeFamilia + "_images");
            }
        }
        private ICommand _SeleccionarCarpetaPersonalDeImagenesThumbnail;
        private bool _sePermiteSeleccionarCarpetaPersonalDeImagenesThumbnail = true;
        public ICommand SeleccionarCarpetaPersonalDeImagenesThumbnail
        {
            get { return _SeleccionarCarpetaPersonalDeImagenesThumbnail ?? (_SeleccionarCarpetaPersonalDeImagenesThumbnail = new GestorDeComandos(() => Accion_SeleccionarCarpetaPersonalDeImagenesThumbnail(), _sePermiteSeleccionarCarpetaPersonalDeImagenesThumbnail)); }
        }
        public void Accion_SeleccionarCarpetaPersonalDeImagenesThumbnail()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
                pathDeLaCarpetaPersonalDeImagenesThumbnail = dialogo.FileName;
            }
        }

        private bool? _gruposEnHorizontal;
        public bool? gruposEnHorizontal
        {
            get { return _gruposEnHorizontal; }
            set
            {
                _gruposEnHorizontal = value;
                NotifyPropertyChanged("gruposEnHorizontal");
            }
        }
        private bool? _gruposEnVertical;
        public bool? gruposEnVertical
        {
            get { return _gruposEnVertical; }
            set
            {
                _gruposEnVertical = value;
                NotifyPropertyChanged("gruposEnVertical");
            }
        }

        private bool? _limitarTamainoElementoGrupo;
        public bool? limitarTamainoElementoGrupo
        {
            get { return _limitarTamainoElementoGrupo; }
            set
            {
                _limitarTamainoElementoGrupo = value;
                NotifyPropertyChanged("limitarTamainoElementoGrupo");
                if (value == false)
                { limiteDeTamainoElementoGrupo = ""; }
                else
                { limiteDeTamainoElementoGrupo = "125"; }

            }
        }
        private string _limiteDeTamainoElementoGrupo;
        public string limiteDeTamainoElementoGrupo
        {
            get { return _limiteDeTamainoElementoGrupo; }
            set
            {
                _limiteDeTamainoElementoGrupo = value;
                NotifyPropertyChanged("limiteDeTamainoElementoGrupo");
            }
        }
 
        private bool? _usarNombresCortosEnLosGrupos;
        public bool? usarNombresCortosEnLosGrupos
        {
            get { return _usarNombresCortosEnLosGrupos; }
            set
            {
                _usarNombresCortosEnLosGrupos = value;
                NotifyPropertyChanged("usarNombresCortosEnLosGrupos");
            }
        }

        private bool? _mostrarNombresDescriptivosEnLasFamilias;
        public bool? mostrarNombresDescriptivosEnLasFamilias
        {
            get { return _mostrarNombresDescriptivosEnLasFamilias; }
            set
            {
                _mostrarNombresDescriptivosEnLasFamilias = value;
                NotifyPropertyChanged("mostrarNombresDescriptivosEnLasFamilias");
            }
        }
        private bool? _mostrarNombresDeArchivoEnLasFamilias;
        public bool? mostrarNombresDeArchivoEnLasFamilias
        {
            get { return _mostrarNombresDeArchivoEnLasFamilias; }
            set
            {
                _mostrarNombresDeArchivoEnLasFamilias = value;
                NotifyPropertyChanged("mostrarNombresDeArchivoEnLasFamilias");
            }
        }
        private bool usarNombresDeArchivoEnLasFamilias
        {
            get
            {
                if (mostrarNombresDeArchivoEnLasFamilias == true)
                { return true; }
                else
                { return false; }
            }
        }

        private ICommand _SalirYGuardarOpciones;
        private bool _sePermiteSalirYGuardarOpciones = true;
        public ICommand SalirYGuardarOpciones
        {
            get { return _SalirYGuardarOpciones ?? (_SalirYGuardarOpciones = new GestorDeComandos(() => Accion_SalirYGuardarOpciones(), _sePermiteSalirYGuardarOpciones)); }
        }
        public void Accion_SalirYGuardarOpciones()
        {
            string orientacionGrupos;
            if (gruposEnVertical == true)
            {
                orientacionGrupos = "V";
            }
            else
            {
                orientacionGrupos = "H";
            }
            bool nombresCortos;
            if (usarNombresCortosEnLosGrupos == true)
            { nombresCortos = true; }
            else
            { nombresCortos = false; }
            bool nombresDeArchivo;
            if (usarNombresDeArchivoEnLasFamilias == true)
            { nombresDeArchivo = true; }
            else
            { nombresDeArchivo = false; }

            Opciones opciones = Opciones.getOpcionesAlmacenadas(mostrarAvisoEnCasoDeError: false);
            opciones.idFilialConLaQueTrabajar = filialSeleccionada;
            opciones.modoDeObtenerLosDatosBDI = modoDeObtenerLosDatosBDI;
            opciones.pathDeLaCarpetaConLosXMLParaOffline = pathDeLaCarpetaConLosXMLParaOffline;
            opciones.idiomaParaLosNombres = idiomaSeleccionado;
            opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia = pathDeLaCarpetaBaseDeArchivosDeFamilia;
            opciones.pathDeLaCarpetaBaseDeImagenesThumbnail = pathDeLaCarpetaBaseDeImagenesThumbnail;
            opciones.pathDeLaCarpetaPersonalDeArchivosDeFamilia = pathDeLaCarpetaPersonalDeArchivosDeFamilia;
            opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail = pathDeLaCarpetaPersonalDeImagenesThumbnail;
            opciones.orientacionParaLaListaDeGrupos = orientacionGrupos;
            opciones.limiteDeTamainoElementoGrupo = limiteDeTamainoElementoGrupo;
            opciones.usarNombresCortosEnLosGrupos = nombresCortos;
            opciones.usarNombresDeArchivoEnLasFamilias = nombresDeArchivo;
            opciones.AlmacenarOpciones();
            enlaceConLaVentanaUCBrowser.InicializarDatos();
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
