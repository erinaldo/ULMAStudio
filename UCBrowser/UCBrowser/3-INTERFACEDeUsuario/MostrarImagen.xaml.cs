using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace UCBrowser
{
    /// <summary>
    /// Interaction logic for MostrarImagen.xaml
    /// </summary>
    public partial class MostrarImagen : Window
    {
        public MostrarImagen()
        {
            InitializeComponent();
        }


        public void putImagenAMostrar(System.Windows.Media.Imaging.BitmapImage imagenAMostrar, string titulo)
        {
            this.visorDeImagen.Source = imagenAMostrar;
            this.Title = titulo;
        }

    
    }
}
