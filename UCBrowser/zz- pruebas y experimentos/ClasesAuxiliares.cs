using System;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.IO;




namespace zz__pruebas_y_experimentos
{

    public class GestorDeComandos : ICommand
    {
        private Action _accionAEjecutar;
        private bool _sePermiteEjecutarElComando;

        public GestorDeComandos(Action accionAEjecutar, bool esPosibleEjecutarElComando)
        {
            _accionAEjecutar = accionAEjecutar;
            _sePermiteEjecutarElComando = esPosibleEjecutarElComando;
        }

        public bool CanExecute(object parameter)
        {
            return _sePermiteEjecutarElComando;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            _accionAEjecutar();
        }
    }




//using System.ComponentModel;
//using System.Windows.Input;

//public event PropertyChangedEventHandler PropertyChanged;

//private void NotifyPropertyChanged(string propertyName)
//{
//    if (PropertyChanged != null)
//    {
//        PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
//    }
//}




}
