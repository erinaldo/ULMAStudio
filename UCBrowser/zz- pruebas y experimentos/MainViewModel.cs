using System.ComponentModel;
using System.Windows.Input;


namespace zz__pruebas_y_experimentos
{
    class MainViewModel : INotifyPropertyChanged
    {
        private string _resultados;
        public string resultados
        {
            get { return _resultados; }
            set
            {
                _resultados = value;
                NotifyPropertyChanged("resultados");
            }
        }


        private ICommand _LanzarPrueba;
        private bool sePermiteLanzarPrueba = true;
        public ICommand LanzarPrueba
        {
            get { return _LanzarPrueba ?? (_LanzarPrueba = new GestorDeComandos(() => LanzarPrueba_accion(), sePermiteLanzarPrueba)); }
        }
        public void LanzarPrueba_accion()
        {
            resultados = "prueba ejecutada";
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
