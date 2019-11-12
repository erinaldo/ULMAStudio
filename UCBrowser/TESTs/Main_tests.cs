using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UCBrowser;

namespace UCBrowser_TESTs
{
    [TestClass]
    public class Main_tests
    {
        Main_viewmodel main;

        public Main_tests()
        {
            main = new Main_viewmodel();
        }

        private TestContext testContextInstance;

        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion


        [TestMethod]
        public void LosDatosInicialesSeHanInicializadoCorrectamente()
        {
            Assert.IsTrue(main.lineasDeProducto.Count > 0);
            //Assert.IsNotNull(main.opciones);
        }

        [TestMethod]
        public void AlSeleccionarUnaLineaDeProductoSeActualizaLaListaDeGruposYAlSeleccionarUnGrupoSeActualizaLaListaDeFamilias()
        {
            int cuentaGruposAnterior = 0;
            foreach(LineaDeProducto producto in main.lineasDeProducto)
            {
                main.lineaDeProductoSeleccionada = producto.id;
                Assert.IsTrue(main.grupos.Count != cuentaGruposAnterior);
                cuentaGruposAnterior = main.grupos.Count;
                int cuentaFamiliasAnterior = 0;
                foreach(GrupoDeFamilias grupo in main.grupos)
                {
                    main.grupoSeleccionado = grupo.id;
                    Assert.IsTrue(main.familias.Count != cuentaFamiliasAnterior);
                    cuentaFamiliasAnterior = main.familias.Count;
                }
            }
        }
    }
}
