using System;
using System.Windows.Data;
using System.Windows.Input;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Media.Imaging;
using System.IO;
using Autodesk.Revit.DB;

namespace UCFamilyProcessor
{

    public class ProcesadorDeWarnings : Autodesk.Revit.DB.IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor accesoALosFallosDetectados)
        {
            accesoALosFallosDetectados.DeleteAllWarnings();

            return FailureProcessingResult.Continue;
        }
    }


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


    public class ConversorDeBitmaps : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is Bitmap)
            {
                var stream = new MemoryStream();
                ((Bitmap)value).Save(stream, ImageFormat.Png);

                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = stream;
                bitmap.EndInit();

                return bitmap;
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

}
