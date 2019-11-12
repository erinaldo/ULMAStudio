using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using UCBrowser;

namespace UCBrowser_TESTs
{
    [TestClass]
    public class Opciones_tests
    {
        [TestMethod]
        public void AlmacenarLasOpcionesPorDefectoYVolverALeerlasDesdeArchivo()
        {
            Opciones_viewmodel gestorA = new Opciones_viewmodel(enlaceConLaVentanaUCBrowser: new Main_viewmodel());

            gestorA.filialSeleccionada = 120;
            gestorA.leerBDIonline = true;
            gestorA.idiomaSeleccionado = "es-ES";

            gestorA.pathDeLaCarpetaConLosXMLParaOffline = "C:/Users/jmurua/AppData/Roaming/Autodesk/Revit/Addins/2018/UCREVIT2018/offlineBDIdata";
            gestorA.pathDeLaCarpetaBaseDeArchivosDeFamilia = "G:/08_RECURSOS DE DISEÑO PARA APLICACIÓN/REVIT_FAMILY_MASTER/REVIT_FAMILY_MASTER_CENTRAL_120";
            gestorA.pathDeLaCarpetaPersonalDeArchivosDeFamilia = "C:/Users/Public/FamiliasRevit";

            gestorA.gruposEnHorizontal = true;
            gestorA.limitarTamainoElementoGrupo = true;
            gestorA.limiteDeTamainoElementoGrupo = "100";
            gestorA.usarNombresCortosEnLosGrupos = true;
            gestorA.mostrarNombresDeArchivoEnLasFamilias = true;

            gestorA.Accion_SalirYGuardarOpciones();


            Opciones_viewmodel gestorB = new Opciones_viewmodel(enlaceConLaVentanaUCBrowser: new Main_viewmodel());

            Assert.AreEqual(gestorA.filialSeleccionada, gestorB.filialSeleccionada);
            Assert.AreEqual(gestorA.leerBDIonline, gestorB.leerBDIonline);
            Assert.AreEqual(gestorA.idiomaSeleccionado, gestorB.idiomaSeleccionado);

            Assert.AreEqual(gestorA.pathDeLaCarpetaConLosXMLParaOffline, gestorB.pathDeLaCarpetaConLosXMLParaOffline);
            Assert.AreEqual(gestorA.pathDeLaCarpetaBaseDeArchivosDeFamilia, gestorB.pathDeLaCarpetaBaseDeArchivosDeFamilia);
            Assert.AreEqual(gestorA.pathDeLaCarpetaPersonalDeArchivosDeFamilia, gestorB.pathDeLaCarpetaPersonalDeArchivosDeFamilia);

            Assert.AreEqual(gestorA.gruposEnHorizontal, gestorB.gruposEnHorizontal);
            Assert.AreEqual(gestorA.limitarTamainoElementoGrupo, gestorB.limitarTamainoElementoGrupo);
            Assert.AreEqual(gestorA.limiteDeTamainoElementoGrupo, gestorB.limiteDeTamainoElementoGrupo);
            Assert.AreEqual(gestorA.usarNombresCortosEnLosGrupos, gestorB.usarNombresCortosEnLosGrupos);
            Assert.AreEqual(gestorA.mostrarNombresDeArchivoEnLasFamilias, gestorB.mostrarNombresDeArchivoEnLasFamilias);
        }
    }
}
