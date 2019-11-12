using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Input;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCFamilyProcessor
{

    public class MainViewModel : INotifyPropertyChanged
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
        private UIApplication app;
        public MainViewModel(UIApplication app)
        {
            tituloDeLaVentana = "UCFamilyProcessor"
                                + "  (v" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString()
                                + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString()
                                + " c" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString() + ") ";
            this.app = app;
            procesarSoloFamiliasSinImagen = true;
            procesarTodasLasFamilias = false;
        }
        public MainViewModel() //Este constructor sin parametros es SOLO PARA TEST; no usarlo para nada mas;
        {
            procesarSoloFamiliasSinImagen = true;
        } //Este constructor sin parametros es SOLO PARA TEST; no usarlo para nada mas;

        private string _pathCarpetaAProcesar;
        public string pathCarpetaAProcesar { get { return _pathCarpetaAProcesar; } }
        private ICommand _SeleccionarCarpetaAProcesar;
        private bool _sePermiteSeleccionarCarpetaAProcesar = true;
        public ICommand SeleccionarCarpetaAProcesar
        {
            get { return _SeleccionarCarpetaAProcesar ?? (_SeleccionarCarpetaAProcesar = new GestorDeComandos(() => Accion_SeleccionarCarpetaAProcesar(), _sePermiteSeleccionarCarpetaAProcesar)); }
        }
        public void Accion_SeleccionarCarpetaAProcesar()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
                _pathCarpetaAProcesar = dialogo.FileName;
                NotifyPropertyChanged("pathCarpetaAProcesar");
                _pathCarpetaDestino = _pathCarpetaAProcesar + "_images";
                NotifyPropertyChanged("pathCarpetaDestino");
            }
        }

        private string _pathCarpetaDestino;
        public string pathCarpetaDestino { get { return _pathCarpetaDestino; } }
        private ICommand _SeleccionarCarpetaDestino;
        private bool _sePermiteSeleccionarCarpetaDestino = true;
        public ICommand SeleccionarCarpetaDestino
        {
            get { return _SeleccionarCarpetaDestino ?? (_SeleccionarCarpetaDestino = new GestorDeComandos(() => Accion_SeleccionarCarpetaDestino(), _sePermiteSeleccionarCarpetaDestino)); }
        }
        public void Accion_SeleccionarCarpetaDestino()
        {
            WPFFolderBrowser.WPFFolderBrowserDialog dialogo = new WPFFolderBrowser.WPFFolderBrowserDialog();
            if (dialogo.ShowDialog() == true)
            {
                _pathCarpetaDestino = dialogo.FileName;
                NotifyPropertyChanged("pathCarpetaDestino");
            }
        }


        private bool? _procesarSoloFamiliasSinImagen;
        public bool? procesarSoloFamiliasSinImagen
        {
            get { return _procesarSoloFamiliasSinImagen; }
            set
            {
                _procesarSoloFamiliasSinImagen = value;
                NotifyPropertyChanged("procesarSoloFamiliasSinImagen");
            }
        }
        private bool? _procesarTodasLasFamilias;
        public bool? procesarTodasLasFamilias
        {
            get { return _procesarTodasLasFamilias; }
            set
            {
                _procesarTodasLasFamilias = value;
                NotifyPropertyChanged("procesarTodasLasFamilias");
            }
        }

        private ICommand _ProcesarFamilias;
        private bool _sePermiteProcesarFamilias = true;
        public ICommand ProcesarFamilias
        {
            get { return _ProcesarFamilias ?? (_ProcesarFamilias = new GestorDeComandos(() => Accion_ProcesarFamilias(), _sePermiteProcesarFamilias)); }
        }
        public void Accion_ProcesarFamilias()
        {
            mensajeEnBarraDeEstado = Traducciones.MainWindow.InicioDelProcesamiento + " " + DateTime.Now.ToString("s");
            if (System.IO.Directory.Exists(pathCarpetaAProcesar))
            {
                if (!System.IO.Directory.Exists(pathCarpetaDestino))
                {
                    System.IO.Directory.CreateDirectory(pathCarpetaDestino);
                }
                Document documentoActivo = app.ActiveUIDocument.Document;
                ProcesarFamiliasEnLaCarpeta(documentoActivo, pathCarpetaAProcesar, pathCarpetaDestino);
            }
            mensajeEnBarraDeEstado = mensajeEnBarraDeEstado + "  ... " + Traducciones.MainWindow.Terminado + " " + DateTime.Now.ToString("s") + "  .";
        }

        private void ProcesarFamiliasEnLaCarpeta(Document documentoActivo, string carpetaEnProceso, string carpetaDestino)
        {
            foreach (string pathArchivoFamilia in System.IO.Directory.GetFiles(carpetaEnProceso, "*.rfa"))
            {
                string pathArchivoImagen = System.IO.Path.Combine(carpetaDestino, System.IO.Path.GetFileName(System.IO.Path.ChangeExtension(pathArchivoFamilia, ".png")));
                if (!(procesarSoloFamiliasSinImagen == true && System.IO.File.Exists(pathArchivoImagen))
                    || procesarTodasLasFamilias == true)
                    try
                    {
                        string nombreFamilia = System.IO.Path.GetFileNameWithoutExtension(pathArchivoFamilia);
                        Family familia;
                        using (Transaction trans = new Transaction(documentoActivo, "Load family for UCFamilyProcessor"))
                        {
                            trans.Start();
                            FailureHandlingOptions enCasoDeFalloOAviso = trans.GetFailureHandlingOptions();
                            enCasoDeFalloOAviso.SetFailuresPreprocessor(new ProcesadorDeWarnings());
                            trans.SetFailureHandlingOptions(enCasoDeFalloOAviso);

                            documentoActivo.LoadFamily(pathArchivoFamilia, new OpcionesDeSobreescrituraDeFamiliasAnidadasYaExistentesEnElDocumento(), out familia);
                                
                            trans.Commit();
                        }
                        ElementId idPrimerSimboloEnLaFamilia = familia.GetFamilySymbolIds().First();
                        FamilySymbol simbolo = (FamilySymbol)familia.Document.GetElement(idPrimerSimboloEnLaFamilia);

                        System.Drawing.Bitmap thumbnail = simbolo.GetPreviewImage(new System.Drawing.Size(100, 100));
                        if (thumbnail != null)
                        {
                            thumbnail.Save(filename: pathArchivoImagen, format: System.Drawing.Imaging.ImageFormat.Png);
                        }
                        else
                        {
                            System.IO.File.Create(System.IO.Path.ChangeExtension(pathArchivoImagen, "dummy"));
                        }

                    }
                    catch (Exception)
                    {
                        continue;
                    }
            }

            //Este bucle recursivo es una chapu para incorporar ANNOTATIONS y lo que se tercie...
            foreach (string subcarpeta in System.IO.Directory.GetDirectories(carpetaEnProceso))
            {
                ProcesarFamiliasEnLaCarpeta(documentoActivo, subcarpeta, pathCarpetaDestino);
            }
        }

        private string _mensajeEnBarraDeEstado;
        public string mensajeEnBarraDeEstado
        {
            get { return _mensajeEnBarraDeEstado; }
            set
            {
                _mensajeEnBarraDeEstado = value;
                NotifyPropertyChanged("mensajeEnBarraDeEstado");
            }
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


    class OpcionesDeSobreescrituraDeFamiliasAnidadasYaExistentesEnElDocumento : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }







}
