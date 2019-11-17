using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using UCBrowser;

//namespace UCBrowser
//{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class Main_window : Window
    {
        //public static Main_window mWindows;
        public Main_window() 
        {
           InitializeComponent();

           Closing += AlCerrarEstaVentana;
        //mWindows = this;
        WindowManager.ClosingWindows += AlCerrarEstaVentana;
    }

        public void actualizar()
        {
            int indice = lfamilies.SelectedIndex;
            lfamilies.Items.Refresh();
            lfamilies.SelectedIndex = -1;
            lfamilies.SelectedIndex = indice;
        }

        public void AlCerrarEstaVentana(object sender, CancelEventArgs e)
        {
            if (Main.cLcsv != null)
            {
                Main.cLcsv.PonLog_ULMA("BROWSER_CLOSE", EApp: ULMALGFree.queApp.UCBROWSER);
            }
        this.Visibility = System.Windows.Visibility.Hidden;
            e.Cancel = true;
        }
        public void AlCerrarEstaVentana(object sender, System.EventArgs e)
        {
            this.Visibility = System.Windows.Visibility.Hidden;
            if (Main.cLcsv != null)
            {
                Main.cLcsv.PonLog_ULMA("BROWSER_CLOSE", EApp: ULMALGFree.queApp.UCBROWSER);
            }
    }

    private void ScrollConLaRuedaDelRaton(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            ScrollViewer scroll = (ScrollViewer)sender;
            scroll.ScrollToVerticalOffset(scroll.VerticalOffset - e.Delta);
            e.Handled = true;
        }

        private void IniciarBusquedaSiSePulsaEnter(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                this.btnBuscarFamilias.Focus();
                ((UCBrowser.Main_viewmodel)this.DataContext).Accion_BuscarFamilias();
            }
        }

        private void InsertarLaFamiliaSeleccionadaConDobleClic(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            ((UCBrowser.Main_viewmodel)this.DataContext).InsertarFamiliaSeleccionada();
        }

        private void InsertarFamiliaConArrastrarYSoltar(object sender, DragEventArgs e)
        {
        UCBrowser.Familia datosRecibidos = (UCBrowser.Familia)e.Data.GetData(typeof(UCBrowser.Familia));
            ((UCBrowser.Main_viewmodel)this.DataContext).InsertarFamilia(familia: datosRecibidos);

        }

        private void AlIniciarMovimientoDeRatonEnListaDeFamilias(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                ListView lista = (ListView)sender;
                if (lista != null)
                {
                    DragDrop.DoDragDrop(dragSource: lista, data: lista.SelectedItem, allowedEffects: DragDropEffects.Copy);
                }
            }

        }

        private void AlRecibirDropEnBotonFavoritos(object sender, DragEventArgs e)
        {
        UCBrowser.Familia datosRecibidos = (UCBrowser.Familia)e.Data.GetData(typeof(UCBrowser.Familia));
            ((UCBrowser.Main_viewmodel)this.DataContext).AgregarAFavoritos(familia: datosRecibidos);
        }

        private void AlRecibirDropEnAlgunSitioFueraDeFavoritos(object sender, DragEventArgs e)
        {
        UCBrowser.Familia datosRecibidos = (UCBrowser.Familia)e.Data.GetData(typeof(UCBrowser.Familia));
            ((UCBrowser.Main_viewmodel)this.DataContext).QuitarDeFavoritos(familia: datosRecibidos);
        }


    }