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
    /// Interaction logic for Opciones_window.xaml
    /// </summary>
    public partial class Opciones_window : Window
    {
        public Opciones_window()
        {
            InitializeComponent();
        }

        private void btnAceptar_click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

    }
}
