using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UCBrowser;


namespace UCBrowser_TESTs
{
    [TestClass]
    public class BibliotecaDeFamilias_tests
    {
        public BibliotecaDeFamilias_tests()
        {
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
        public void LecturaDeFilialesDirectamenteDeLaBDI_online()
        {
            List<Filial> lista = BibliotecaDeFamilias.getFilialesDirectamenteDeLaBDI(ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.online);
            Assert.IsNotNull(lista);
            Assert.IsTrue(lista.Count > 0);
        }

        [TestMethod]
        public void LecturaDeFilialesDirectamenteDeLaBDI_offline()
        {
            List<Filial> lista = BibliotecaDeFamilias.getFilialesDirectamenteDeLaBDI(ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.offline);
            Assert.IsNotNull(lista);
            Assert.IsTrue(lista.Count > 0);
        }

        [TestMethod]
        public void LecturaDeIdiomasDirectamenteDeLaBDI_online()
        {
            List<string> lista = BibliotecaDeFamilias.getIdiomasDirectamenteDeLaBDI(ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.online);
            Assert.IsNotNull(lista);
            Assert.IsTrue(lista.Count > 0);
        }

        [TestMethod]
        public void LecturaDeIdiomasDirectamenteDeLaBDI_offline()
        {
            List<string> lista = BibliotecaDeFamilias.getIdiomasDirectamenteDeLaBDI(ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.offline);
            Assert.IsNotNull(lista);
            Assert.IsTrue(lista.Count > 0);
        }




    }
}

