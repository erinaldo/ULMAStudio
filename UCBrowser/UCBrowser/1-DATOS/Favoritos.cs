using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace UCBrowser
{
    class Favoritos
    {
        //private static string archivoFavoritos = (Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + Path.DirectorySeparatorChar
        //                                          + "UCBrowser" + Path.DirectorySeparatorChar + "UCBrowserFavorites.xml");
        private static string archivoFavoritos = (System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                                                 + Path.DirectorySeparatorChar + "UCBrowserFavorites.xml");

        public List<Familia> familiasFavoritas;

        public Favoritos()
        {
            familiasFavoritas = new List<Familia>();
        }

        private const string ETIQUETAraizFAVORITOS = "Favorits";
        private const string ETIQUETAfamilia = "family";

        public void GuardarEnDisco()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlNode Nraiz = doc.CreateElement(ETIQUETAraizFAVORITOS);

                foreach (Familia familia in familiasFavoritas)
                {
                    XmlNode Nfamilia = doc.CreateElement(ETIQUETAfamilia);

                    XmlNode nID = doc.CreateElement("id");
                    nID.InnerText = familia.id;
                    Nfamilia.AppendChild(nID);

                    XmlNode nName = doc.CreateElement("name");
                    nName.InnerText = familia.nombre;
                    Nfamilia.AppendChild(nName);

                    XmlNode nFile = doc.CreateElement("file");
                    nFile.InnerText = familia.nombreArchivo;
                    Nfamilia.AppendChild(nFile);

                    Nraiz.AppendChild(Nfamilia);
                }

                doc.AppendChild(Nraiz);

                Directory.CreateDirectory(Path.GetDirectoryName(archivoFavoritos));
                FileStream archivo = File.Create(archivoFavoritos);
                doc.Save(archivo);
                archivo.Flush();
                archivo.Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Some problem saving favoritsThe favorits maybe not saved properly." + Environment.NewLine + Environment.NewLine
                                               + archivoFavoritos + Environment.NewLine + Environment.NewLine
                                               + ex.Message + Environment.NewLine + Environment.NewLine + ex.StackTrace,"ATTENTION",
                                               System.Windows.MessageBoxButton.OK, 
                                               System.Windows.MessageBoxImage.Warning);
            }
        }


        public void RecuperarDelDisco(string pathDeLaCarpetaPersonalParaLlegarHastaLosThumbnails, string pathDeLaCarpetaBaseParaLlegarHastaLosThumbnails)
        {
            familiasFavoritas.Clear();
            try
            {
                if (File.Exists(archivoFavoritos))
                {
                    XmlDocument doc = new XmlDocument();
                    FileStream archivo = File.OpenRead(archivoFavoritos);
                    doc.Load(archivo);
                    archivo.Close();

                    foreach (XmlNode nFamilia in doc.SelectNodes("//" + ETIQUETAfamilia))
                    {
                        Familia familia = new Familia(id: nFamilia.SelectSingleNode("id").InnerText,
                                                      nombre: nFamilia.SelectSingleNode("name").InnerText,
                                                      nombreArchivo: nFamilia.SelectSingleNode("file").InnerText,
                                                      esDinamica: false, esConjunto: false, esAnnotationSymbol: false);
                        familia.thumbnail = BibliotecaDeFamilias.ObtenerImagen(nombreArchivoFamilia: familia.nombreArchivo,
                                                                               pathDeLaCarpetaPersonalDeImagenesThumbnail: pathDeLaCarpetaPersonalParaLlegarHastaLosThumbnails,
                                                                               pathDeLaCarpetaBaseDeImagenesThumbnail: pathDeLaCarpetaBaseParaLlegarHastaLosThumbnails);
                        familiasFavoritas.Add(familia);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Some problem reading favorits." + Environment.NewLine + Environment.NewLine
                                               + archivoFavoritos + Environment.NewLine + ex.Message, "ATTENTION",
                                               System.Windows.MessageBoxButton.OK,
                                               System.Windows.MessageBoxImage.Warning);
            }
        }


    }
}
