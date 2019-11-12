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
    /// Interaction logic for Busqueda_window.xaml
    /// </summary>
    public partial class ResultadosBusqueda_window : Window
    {
        public ResultadosBusqueda_window()
        {
            InitializeComponent();
        }

        private void InsertarLaFamiliaSeleccionadaConDobleClic(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            ((ResultadosBusqueda_viewmodel)this.DataContext).InsertarFamiliaSeleccionada();
            DialogResult = true;
        }


        private void InsertarFamiliaConArrastrarYSoltar(object sender, DragEventArgs e)
        {
            Familia datosRecibidos = (Familia)e.Data.GetData(typeof(Familia));
            ((ResultadosBusqueda_viewmodel)this.DataContext).InsertarFamilia(familia: datosRecibidos);
            DialogResult = true;
        }

        private void AlIniciarMovimientoDeRatonEnListaDeFamilias(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                ListView lista = (ListView)sender;
                if (lista != null && lista.SelectedItem != null)
                {
                    DragDrop.DoDragDrop(dragSource: lista, data: lista.SelectedItem, allowedEffects: DragDropEffects.Copy);
                }
            }

        }



        private void ScrollConLaRuedaDelRaton(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            ScrollViewer scroll = (ScrollViewer)sender;
            scroll.ScrollToVerticalOffset(scroll.VerticalOffset - e.Delta);
            e.Handled = true;
        }

    }
}
