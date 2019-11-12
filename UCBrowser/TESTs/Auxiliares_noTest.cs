using Microsoft.VisualStudio.TestTools.UnitTesting;

using UCBrowser;
using UCFamilyProcessor;

namespace TESTs
{
    [TestClass]
    public class Auxiliares_noTest
    {

        [TestMethod]
        public void auxiliarParaProbarManualmenteLaNAVEGACIONEnUCBrowserSinNecesidadDeArrancarRevit()
        {
            Main_window ventana = new Main_window();
            ventana.DataContext = new Main_viewmodel();
            ventana.ShowDialog();
        }

        [TestMethod]
        public void auxiliarParaProbarManualmenteElDialogoDeOPCIONESSinNecesidadDeArrancarRevit()
        {
            Opciones_window ventana = new Opciones_window();
            ventana.DataContext = new Opciones_viewmodel(enlaceConLaVentanaUCBrowser: new Main_viewmodel());
            ventana.ShowDialog();
        }

        [TestMethod]
        public void auxiliarParaProbarManualmenteElDialogoFamilyPROCESSORSinNecesidadDeArrancarRevit()
        {
            MainWindow ventana = new MainWindow();
            ventana.DataContext = new MainViewModel();
            ventana.ShowDialog();
        }

    }
}
