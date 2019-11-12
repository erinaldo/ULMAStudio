using System;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{
    public class ResultadosBusqueda_viewmodel : INotifyPropertyChanged
    {
        private ProcesadorDeComandosRevit procesadorDeComandosRevit;
        private Autodesk.Revit.UI.ExternalEvent lanzarProcesadorDeComandosRevit;
        private Opciones opciones;

        public ResultadosBusqueda_viewmodel(List<Familia> familiasEncontradas,
                                            ProcesadorDeComandosRevit procesadorDeComandosRevit, 
                                            Autodesk.Revit.UI.ExternalEvent lanzarProcesadorDeComandosRevit)
        {
            this.procesadorDeComandosRevit = procesadorDeComandosRevit;
            procesadorDeComandosRevit.comandoAEjecutar = ProcesadorDeComandosRevit.ComandosDisponibles.noHacerNada;
            this.lanzarProcesadorDeComandosRevit = lanzarProcesadorDeComandosRevit;

            // ALBERTO. Que solo muestre las familias .rfa que tengan imagen asociada.
            //if (opciones == null) { opciones = Opciones.getOpcionesAlmacenadas(mostrarAvisoEnCasoDeError: false); };
            //if (familiasimg == null) { familiasimg = new List<Familia>(); };
            //familiasimg.Clear();
            //for (int x = 0; x < familiasEncontradas.Count - 1; x++)
            //{
            //    //if (System.IO.Path.Combine(pathDeLaCarpetaPersonalDeImagenesThumbnail, _familias[x].nombreArchivo))
            //    if (File.Exists(System.IO.Path.Combine(opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia, familiasEncontradas[x].nombreArchivo)))
            //    {
            //        familiasimg.Add(familiasEncontradas[x]);
            //    }
            //}
            //familiasEncontradas.Clear();
            //familiasEncontradas = familiasimg;
            // ************************************************************************

            _familias = familiasEncontradas;
            NotifyPropertyChanged("familias");

            opciones = Opciones.getOpcionesAlmacenadas(mostrarAvisoEnCasoDeError: false);
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
        //nota: no olvidar mirar su funcion siamesa que inserta familias desde UCBrowser_viewmodel.



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
