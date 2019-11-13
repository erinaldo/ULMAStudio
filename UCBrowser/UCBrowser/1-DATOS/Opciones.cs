using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using uf = ULMALGFree.clsBase;

namespace UCBrowser
{

    public class Opciones
    {
        public int idFilialConLaQueTrabajar;
        public ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos modoDeObtenerLosDatosBDI;
        public string pathDeLaCarpetaConLosXMLParaOffline;
        public string idiomaParaLosNombres;

        public string pathDeLaCarpetaBaseDeArchivosDeFamilia;
        public string pathDeLaCarpetaBaseDeImagenesThumbnail;
        public string pathDeLaCarpetaPersonalDeArchivosDeFamilia;
        public string pathDeLaCarpetaPersonalDeImagenesThumbnail;

        public string orientacionParaLaListaDeGrupos;
        public string limiteDeTamainoElementoGrupo;
        public bool usarNombresCortosEnLosGrupos;

        public bool usarNombresDeArchivoEnLasFamilias;


        private Opciones() { }
        //private static string ARCHIVO_OPCIONES = (Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + Path.DirectorySeparatorChar
        //                                          + "UCBrowser" + Path.DirectorySeparatorChar + "UCBrowserOptions.xml");
        private static string ARCHIVO_OPCIONES = (System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                                                  + Path.DirectorySeparatorChar + "UCBrowserOptions.xml");
    
        public static Opciones getOpcionesPorDefecto()
        {
            // Alberto. Coger las opciones de uf (ULMALGFree.clsBase)
            Opciones op = new Opciones();
            op.idFilialConLaQueTrabajar = Convert.ToInt32(uf.DEFAULT_PROGRAM_MARKET);
            op.modoDeObtenerLosDatosBDI = ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.offline;
            op.pathDeLaCarpetaConLosXMLParaOffline = System.IO.Path.GetFullPath(uf.path_offlineBDIdata);
            op.idiomaParaLosNombres = uf.DEFAULT_PROGRAM_LANGUAGE;
            // Alberto.
            string baseFamilia = uf.path_families_base.StartsWith(".\\") ? uf.path_families_base.Replace(".\\", uf._LgFullFolder) : uf.path_families_base;
            string baseFamiliaImagen = uf.path_families_base_images.StartsWith(".\\") ? uf.path_families_base_images.Replace(".\\", uf._LgFullFolder) : uf.path_families_base_images;
            string baseCustom = uf.path_families_custom.StartsWith(".\\") ? uf.path_families_custom.Replace(".\\", uf._LgFullFolder) : uf.path_families_custom;
            string baseCustomImagen = uf.path_families_custom_images.StartsWith(".\\") ? uf.path_families_custom_images.Replace(".\\", uf._LgFullFolder) : uf.path_families_custom_images;
            op.pathDeLaCarpetaBaseDeArchivosDeFamilia = baseFamilia;
            op.pathDeLaCarpetaBaseDeImagenesThumbnail = baseFamiliaImagen;
            op.pathDeLaCarpetaPersonalDeArchivosDeFamilia = baseCustom;
            op.pathDeLaCarpetaPersonalDeImagenesThumbnail = baseCustomImagen;
            //
            op.orientacionParaLaListaDeGrupos = "H";
            op.limiteDeTamainoElementoGrupo = "125";
            op.usarNombresCortosEnLosGrupos = true;
            op.usarNombresDeArchivoEnLasFamilias = true;
            op.AlmacenarOpciones();
            return op;
        }
        //public static Opciones getOpcionesPorDefectoDeJuan()
        //{
        //    Opciones op = new Opciones();
        //    op.idFilialConLaQueTrabajar = 120;
        //    op.modoDeObtenerLosDatosBDI = ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos.offline;
        //    op.pathDeLaCarpetaConLosXMLParaOffline = (System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        //                                              + Path.DirectorySeparatorChar + "offlineBDIdata");
        //    op.idiomaParaLosNombres = "en";
        //    op.pathDeLaCarpetaBaseDeArchivosDeFamilia = (System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        //                                              + Path.DirectorySeparatorChar + "families");
        //    op.pathDeLaCarpetaBaseDeImagenesThumbnail = (System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        //                                              + Path.DirectorySeparatorChar + "families_images");
        //    op.pathDeLaCarpetaPersonalDeArchivosDeFamilia = "";
        //    op.pathDeLaCarpetaPersonalDeImagenesThumbnail = "";
        //    op.orientacionParaLaListaDeGrupos = "H";
        //    op.limiteDeTamainoElementoGrupo = "125";
        //    op.usarNombresCortosEnLosGrupos = true;
        //    op.usarNombresDeArchivoEnLasFamilias = true;
        //    return op;
        //}

        public static Opciones getOpcionesAlmacenadas(bool mostrarAvisoEnCasoDeError)
        {
            Opciones opciones = new Opciones();
            Opciones pordefecto = Opciones.getOpcionesPorDefecto();

            XmlDocument doc = new XmlDocument();

            string errores = "";
            string error1 = "";
            try
            {
                FileStream archivo = File.OpenRead(ARCHIVO_OPCIONES);
                doc.Load(archivo);
                archivo.Close();
            }
            catch (Exception ex)
            {
                opciones = pordefecto;
                errores = errores + "    (" + ex.Message + ")" + Environment.NewLine + Environment.NewLine; 
            }

            try
            {
                opciones.idFilialConLaQueTrabajar = int.Parse(doc.SelectSingleNode("//" + ETIQUETAidFilialConLaQueTrabajar).InnerText);
            }
            catch (Exception)
            {
                opciones.idFilialConLaQueTrabajar = pordefecto.idFilialConLaQueTrabajar;
                errores = errores + ETIQUETAidFilialConLaQueTrabajar + Environment.NewLine;
            }
            
            try
            {
                Enum.TryParse<ConsultarBDI.CONSULTAS.ModoDeObtenerLosDatos>(doc.SelectSingleNode("//" + ETIQUETAmodoDeObtenerLosDatosBDI).InnerText,
                                                                            out opciones.modoDeObtenerLosDatosBDI);
            }
            catch (Exception)
            { 
                opciones.modoDeObtenerLosDatosBDI = pordefecto.modoDeObtenerLosDatosBDI;
                errores = errores + ETIQUETAmodoDeObtenerLosDatosBDI + Environment.NewLine; 
            }

            //try
            //{
            //    opciones.pathDeLaCarpetaConLosXMLParaOffline = doc.SelectSingleNode("//" + ETIQUETApathDeLaCarpetaConLosXMLParaOffline).InnerText;
            //}
            //catch (Exception)
            //{
            //    opciones.pathDeLaCarpetaConLosXMLParaOffline = pordefecto.pathDeLaCarpetaConLosXMLParaOffline;
            //    errores = errores + ETIQUETApathDeLaCarpetaConLosXMLParaOffline + Environment.NewLine;
            //}
            if (string.IsNullOrWhiteSpace(opciones.pathDeLaCarpetaConLosXMLParaOffline))
            {
                opciones.pathDeLaCarpetaConLosXMLParaOffline = pordefecto.pathDeLaCarpetaConLosXMLParaOffline;
            }

            try
            {
                opciones.idiomaParaLosNombres = doc.SelectSingleNode("//" + ETIQUETAidiomaParaLosNombres).InnerText;
            }
            catch (Exception)
            { 
                opciones.idiomaParaLosNombres = pordefecto.idiomaParaLosNombres;
                //errores = errores + ETIQUETAidiomaParaLosNombres + Environment.NewLine; 
            }

            //try
            //{
            //    opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia = doc.SelectSingleNode("//" + ETIQUETApathDeLaCarpetaBaseDeArchivosDeFamilia).InnerText;
            //}
            //catch (Exception)
            //{ 
            //    opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia = pordefecto.pathDeLaCarpetaBaseDeArchivosDeFamilia;
            //    errores = errores + ETIQUETApathDeLaCarpetaBaseDeArchivosDeFamilia + Environment.NewLine; 
            //}
            if (string.IsNullOrWhiteSpace(opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia))
            {
                opciones.pathDeLaCarpetaBaseDeArchivosDeFamilia = pordefecto.pathDeLaCarpetaBaseDeArchivosDeFamilia;
                opciones.pathDeLaCarpetaBaseDeImagenesThumbnail = pordefecto.pathDeLaCarpetaBaseDeImagenesThumbnail;
            }

            //try
            //{
            //    opciones.pathDeLaCarpetaBaseDeImagenesThumbnail = doc.SelectSingleNode("//" + ETIQUETApathDeLaCarpetaBaseDeImagenesThumbnail).InnerText;
            //}
            //catch (Exception)
            //{
            //    opciones.pathDeLaCarpetaBaseDeImagenesThumbnail = pordefecto.pathDeLaCarpetaBaseDeImagenesThumbnail;
            //    errores = errores + ETIQUETApathDeLaCarpetaBaseDeImagenesThumbnail + Environment.NewLine;
            //}

            try
            {
                opciones.pathDeLaCarpetaPersonalDeArchivosDeFamilia = doc.SelectSingleNode("//" + ETIQUETApathDeLaCarpetaPersonalDeArchivosDeFamilia).InnerText;
            }
            catch (Exception)
            { errores = errores + ETIQUETApathDeLaCarpetaPersonalDeArchivosDeFamilia + Environment.NewLine; }

            try
            {
                opciones.pathDeLaCarpetaPersonalDeImagenesThumbnail = doc.SelectSingleNode("//" + ETIQUETApathDeLaCarpetaPersonalDeImagenesThumbnail).InnerText;
            }
            catch (Exception)
            { errores = errores + ETIQUETApathDeLaCarpetaPersonalDeImagenesThumbnail + Environment.NewLine; }

            try
            {
                opciones.orientacionParaLaListaDeGrupos = doc.SelectSingleNode("//" + ETIQUETAorientacionParaLaListaDeGrupos).InnerText;
            }
            catch (Exception)
            { errores = errores + ETIQUETAorientacionParaLaListaDeGrupos + Environment.NewLine; }


            try
            {
                opciones.limiteDeTamainoElementoGrupo = doc.SelectSingleNode("//" + ETIQUETAlimiteDeTamainoElementoGrupo).InnerText;
            }
            catch (Exception)
            { 
                opciones.limiteDeTamainoElementoGrupo = pordefecto.limiteDeTamainoElementoGrupo;
                errores = errores + ETIQUETAlimiteDeTamainoElementoGrupo + Environment.NewLine; 
            }


            try
            {
                bool.TryParse(doc.SelectSingleNode("//" + ETIQUETAusarNombresCortosEnLosGrupos).InnerText,
                              out opciones.usarNombresCortosEnLosGrupos);

            }
            catch (Exception)
            { 
                opciones.usarNombresCortosEnLosGrupos = pordefecto.usarNombresCortosEnLosGrupos;
                errores = errores + ETIQUETAusarNombresCortosEnLosGrupos + Environment.NewLine; 
            }

            try
            {
                bool.TryParse(doc.SelectSingleNode("//" + ETIQUETAusarNombresDeArchivoEnLasFamilias).InnerText,
                              out opciones.usarNombresDeArchivoEnLasFamilias);
            }
            catch (Exception)
            {
                opciones.usarNombresDeArchivoEnLasFamilias = pordefecto.usarNombresDeArchivoEnLasFamilias;
                errores = errores + ETIQUETAusarNombresDeArchivoEnLasFamilias + Environment.NewLine;
            }


            if (!string.IsNullOrWhiteSpace(errores))
            {
                opciones.AlmacenarOpciones();

                if (mostrarAvisoEnCasoDeError)
                {
                    errores = "Some problem reading settings file." + Environment.NewLine + Environment.NewLine
                              + ARCHIVO_OPCIONES + Environment.NewLine + Environment.NewLine
                              + "=======================================" + Environment.NewLine + Environment.NewLine
                              + errores;
                    System.Windows.MessageBox.Show(errores, "UCBrowser", 
                        System.Windows.MessageBoxButton.OK, 
                        System.Windows.MessageBoxImage.Warning);
                }
            }
            if (!string.IsNullOrWhiteSpace(error1))
            {
                if (mostrarAvisoEnCasoDeError)
                {
                    System.Windows.MessageBox.Show(error1, "UCBrowser", 
                        System.Windows.MessageBoxButton.OK, 
                        System.Windows.MessageBoxImage.Information);
                }
            }
            
            return opciones;
        }

 
        private const string ETIQUETAraizOPCIONES = "UCBrowserOptions";
        private const string ETIQUETAidFilialConLaQueTrabajar = "Market";
        private const string ETIQUETAmodoDeObtenerLosDatosBDI = "BDIdataRetrievalMode";
        private const string ETIQUETApathDeLaCarpetaConLosXMLParaOffline = "OfflineBDIdataFolderPath";
        private const string ETIQUETApathDeLaCarpetaBaseDeArchivosDeFamilia = "FamilyFilesBaseFolderPath";
        private const string ETIQUETApathDeLaCarpetaBaseDeImagenesThumbnail = "ThumbnailFilesBaseFolderPath";
        private const string ETIQUETApathDeLaCarpetaPersonalDeArchivosDeFamilia = "FamilyFilesCustomFolderPath";
        private const string ETIQUETApathDeLaCarpetaPersonalDeImagenesThumbnail = "ThumbnailFilesCustomFolderPath";
        private const string ETIQUETAidiomaParaLosNombres = "LanguajeForNames";
        private const string ETIQUETAorientacionParaLaListaDeGrupos = "GroupListItemOrientation";
        private const string ETIQUETAlimiteDeTamainoElementoGrupo = "GroupListItemCropLimit";
        private const string ETIQUETAusarNombresCortosEnLosGrupos = "UseShortNamesForGroups";
        private const string ETIQUETAusarNombresDeArchivoEnLasFamilias = "UseFilenamesForFamilies";

        public void AlmacenarOpciones()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlNode nRaiz = doc.CreateElement(ETIQUETAraizOPCIONES);

                XmlNode NidFilialConLaQueTrabajar = doc.CreateElement(ETIQUETAidFilialConLaQueTrabajar);
                NidFilialConLaQueTrabajar.InnerText = this.idFilialConLaQueTrabajar.ToString();
                nRaiz.AppendChild(NidFilialConLaQueTrabajar);

                XmlNode NmodoDeObtenerLosDatosBDI = doc.CreateElement(ETIQUETAmodoDeObtenerLosDatosBDI);
                NmodoDeObtenerLosDatosBDI.InnerText = this.modoDeObtenerLosDatosBDI.ToString();
                nRaiz.AppendChild(NmodoDeObtenerLosDatosBDI);

                XmlNode NpathDeLaCarpetaConLosXMLParaOffline = doc.CreateElement(ETIQUETApathDeLaCarpetaConLosXMLParaOffline);
                NpathDeLaCarpetaConLosXMLParaOffline.InnerText = this.pathDeLaCarpetaConLosXMLParaOffline;
                nRaiz.AppendChild(NpathDeLaCarpetaConLosXMLParaOffline);

                XmlNode NidiomaParaLosNombres = doc.CreateElement(ETIQUETAidiomaParaLosNombres);
                NidiomaParaLosNombres.InnerText = this.idiomaParaLosNombres;
                nRaiz.AppendChild(NidiomaParaLosNombres);

                XmlNode NpathDeLaCarpetaBaseDeArchivosDeFamilia = doc.CreateElement(ETIQUETApathDeLaCarpetaBaseDeArchivosDeFamilia);
                NpathDeLaCarpetaBaseDeArchivosDeFamilia.InnerText = this.pathDeLaCarpetaBaseDeArchivosDeFamilia;
                nRaiz.AppendChild(NpathDeLaCarpetaBaseDeArchivosDeFamilia);

                XmlNode NpathDeLaCarpetaBaseDeImagenesThumbnail = doc.CreateElement(ETIQUETApathDeLaCarpetaBaseDeImagenesThumbnail);
                NpathDeLaCarpetaBaseDeImagenesThumbnail.InnerText = this.pathDeLaCarpetaBaseDeImagenesThumbnail;
                nRaiz.AppendChild(NpathDeLaCarpetaBaseDeImagenesThumbnail);

                XmlNode NpathDeLaCarpetaPersonalDeArchivosDeFamilia = doc.CreateElement(ETIQUETApathDeLaCarpetaPersonalDeArchivosDeFamilia);
                NpathDeLaCarpetaPersonalDeArchivosDeFamilia.InnerText = this.pathDeLaCarpetaPersonalDeArchivosDeFamilia;
                nRaiz.AppendChild(NpathDeLaCarpetaPersonalDeArchivosDeFamilia);

                XmlNode NpathDeLaCarpetaPersonalDeImagenesThumbnail = doc.CreateElement(ETIQUETApathDeLaCarpetaPersonalDeImagenesThumbnail);
                NpathDeLaCarpetaPersonalDeImagenesThumbnail.InnerText = this.pathDeLaCarpetaPersonalDeImagenesThumbnail;
                nRaiz.AppendChild(NpathDeLaCarpetaPersonalDeImagenesThumbnail);

                XmlNode NorientacionParaLaListaDeGrupos = doc.CreateElement(ETIQUETAorientacionParaLaListaDeGrupos);
                NorientacionParaLaListaDeGrupos.InnerText = this.orientacionParaLaListaDeGrupos;
                nRaiz.AppendChild(NorientacionParaLaListaDeGrupos);

                XmlNode NlimiteDeTamainoElementoGrupo = doc.CreateElement(ETIQUETAlimiteDeTamainoElementoGrupo);
                NlimiteDeTamainoElementoGrupo.InnerText = this.limiteDeTamainoElementoGrupo;
                nRaiz.AppendChild(NlimiteDeTamainoElementoGrupo);

                XmlNode NusarNombresCortosEnLosGrupos = doc.CreateElement(ETIQUETAusarNombresCortosEnLosGrupos);
                NusarNombresCortosEnLosGrupos.InnerText = this.usarNombresCortosEnLosGrupos.ToString();
                nRaiz.AppendChild(NusarNombresCortosEnLosGrupos);

                XmlNode NusarNombresDeArchivoEnLasFamilias = doc.CreateElement(ETIQUETAusarNombresDeArchivoEnLasFamilias);
                NusarNombresDeArchivoEnLasFamilias.InnerText = this.usarNombresDeArchivoEnLasFamilias.ToString();
                nRaiz.AppendChild(NusarNombresDeArchivoEnLasFamilias);

                doc.AppendChild(nRaiz);

                Directory.CreateDirectory(Path.GetDirectoryName(ARCHIVO_OPCIONES));
                FileStream archivo = File.Create(ARCHIVO_OPCIONES);
                doc.Save(archivo);
                archivo.Flush();
                archivo.Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Some problem saving options.The options maybe not saved properly." + Environment.NewLine + Environment.NewLine
                                               + ARCHIVO_OPCIONES + Environment.NewLine + Environment.NewLine
                                               + ex.Message + Environment.NewLine + Environment.NewLine + ex.StackTrace,
                                               caption:"UCBrowser", 
                                               button: System.Windows.MessageBoxButton.OK, 
                                               icon: System.Windows.MessageBoxImage.Warning);
            }
        }


    }

}
