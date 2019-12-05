using System;
using System.Linq;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace UCBrowser
{
 
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [JournalingAttribute(JournalingMode.NoCommandData)]

    public class Main : IExternalApplication
    {
        public static Main aplicacion;
        public Main_window ventanaUCBrowser;
        private ProcesadorDeComandosRevit procesadorDeComandosRevit;
        private ExternalEvent lanzarProcesadorDeComandosRevit;
        public static ULMALGFree.clsINI cIni = new ULMALGFree.clsINI();
        public static ULMALGFree.clsBase cLcsv = new ULMALGFree.clsBase(typeof(UCBrowser).Assembly);

        public Result OnStartup(UIControlledApplication app)
        { 
            aplicacion = this;
            ventanaUCBrowser = null;
            if (cLcsv == null) { cLcsv = new ULMALGFree.clsBase(typeof(UCBrowser).Assembly); }
            //RibbonPanel panelBrowser = null;
            //try
            //{
            //    List<RibbonPanel> panelesExistentes = app.GetRibbonPanels(tabName: "ULMA");
            //    panelBrowser = panelesExistentes.FirstOrDefault<RibbonPanel>(x => x.Name.Equals("Browser"));
            //}
            //catch (Exception)
            //{
            //    app.CreateRibbonTab("ULMA");
            //}
            //try
            //{
            //    if (panelBrowser == null)
            //    {
            //        panelBrowser = app.CreateRibbonPanel(panelName: "Browser", tabName: "ULMA");
            //    }
            //    else
            //    {
            //        panelBrowser.AddSeparator();
            //    }
            //    PushButton btnUCBrowse = (PushButton)panelBrowser.AddItem(new PushButtonData(name: "Browser", text: "family\nbrowser/picker",
            //                                                                                 assemblyName: System.Reflection.Assembly.GetExecutingAssembly().Location,
            //                                                                                 className: "UCBrowser.UCBrowser"));
            //    btnUCBrowse.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            //                                                                          + System.IO.Path.DirectorySeparatorChar + "ULMA.png"));
            return Result.Succeeded;
            //}
            //catch(Exception ex)
            //{
            //    TaskDialog mensajero = new TaskDialog(title:"UCBrowser.Main.OnStartup");
            //    mensajero.CommonButtons = TaskDialogCommonButtons.Ok;
            //    mensajero.MainContent = ex.ToString() + Environment.NewLine + ex.StackTrace;
            //    mensajero.Show();
            //    return Result.Failed;
            //}
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            if (ventanaUCBrowser != null)
            {   
                lanzarProcesadorDeComandosRevit = null;
                procesadorDeComandosRevit = null;
                ventanaUCBrowser.Close();
            }
            return Result.Succeeded;
        }

        public void MostrarVentanaUCBrowser(UIApplication app)
        {
            if (cLcsv == null) { cLcsv = new ULMALGFree.clsBase(typeof(UCBrowser).Assembly); }
            if (cLcsv != null)
            {
                Main.cLcsv.PonLog_ULMA(ULMALGFree.ACTION.BROWSER_OPEN, EApp: ULMALGFree.queApp.ULMASTUDIO);
            }
            if (ventanaUCBrowser == null || (ventanaUCBrowser != null && ULMALGFree.clsBase._recargarBrowser))
            {
                procesadorDeComandosRevit = new ProcesadorDeComandosRevit();
                lanzarProcesadorDeComandosRevit = ExternalEvent.Create(procesadorDeComandosRevit);

                ventanaUCBrowser = new Main_window();                
                ventanaUCBrowser.DataContext = new Main_viewmodel(procesadorDeComandosRevit, lanzarProcesadorDeComandosRevit);
                ventanaUCBrowser.Show();
                ULMALGFree.clsBase._recargarBrowser = false;
            }
            else
            {
                ventanaUCBrowser.Visibility = System.Windows.Visibility.Visible;
                ventanaUCBrowser.Focus();  
            }
        }

    }



    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [JournalingAttribute(JournalingMode.NoCommandData)]

    class UCBrowser : IExternalCommand
    {

        public Result Execute(ExternalCommandData datosDeEntrada, ref string mensaje, ElementSet elementos)
        {
            try
            {
                Main.aplicacion.MostrarVentanaUCBrowser(datosDeEntrada.Application);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog mensajero = new TaskDialog(title: "UCBrowser.Execute");
                mensajero.CommonButtons = TaskDialogCommonButtons.Ok;
                mensajero.MainContent = ex.ToString() + Environment.NewLine + ex.StackTrace;
                mensajero.Show();
                return Result.Failed;
            }
        }
    }



    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [JournalingAttribute(JournalingMode.NoCommandData)]

    public class ProcesadorDeComandosRevit : IExternalEventHandler
    {
        public enum ComandosDisponibles
        {
            InsertarFamilia,
            MostrarImagenBigDeLaFamilia,
            noHacerNada
        }
        private ComandosDisponibles _comandoAEjecutar = ComandosDisponibles.noHacerNada;
        public ComandosDisponibles comandoAEjecutar { set { _comandoAEjecutar = value; } }

        public void Execute(UIApplication app)
        {
            switch (_comandoAEjecutar)
            {
                case ComandosDisponibles.InsertarFamilia:
                    InsertarFamilia(app);
                    break;

                case ComandosDisponibles.MostrarImagenBigDeLaFamilia:
                    VisualizarImagenBIG();
                    break;
            }
        }


        private string _pathArchivoFamilia = "";
        public  string pathArchivoFamilia { set { _pathArchivoFamilia = value; } }

        private void InsertarFamilia(UIApplication app)
        {
            string nombreFamilia = System.IO.Path.GetFileNameWithoutExtension(_pathArchivoFamilia);
            string pathArchivoFamiliaDesencriptado = "";
            //pathArchivoFamiliaDesencriptado = crip2aCAD.clsCR.Fichero_DesencriptaEnTempDame(_pathArchivoFamilia);
            if (string.IsNullOrWhiteSpace(pathArchivoFamiliaDesencriptado))
            {
                pathArchivoFamiliaDesencriptado = _pathArchivoFamilia;
            }
            if (System.IO.File.Exists(pathArchivoFamiliaDesencriptado))
            {
                if (app.ActiveUIDocument.Document.IsFamilyDocument == true){
                    System.Windows.MessageBox.Show("No open Projects to Load Family into", "ATTENTION",
                        System.Windows.MessageBoxButton.OK, 
                        System.Windows.MessageBoxImage.Warning);
                    return;
                }
                Document documentoActivo = app.ActiveUIDocument.Document;
                Family familia;
                Element encontrado = new FilteredElementCollector(documentoActivo).OfClass(typeof(Family)).FirstOrDefault<Element>(x => x.Name.Equals(nombreFamilia));
                if (encontrado == null)
                {
                    using (Transaction trans = new Transaction(documentoActivo, "Load family from UCBrowser"))
                    {
                        trans.Start();
                        ULMALGFree.clsBase._registraLoadInsert = false;
                        documentoActivo.LoadFamily(pathArchivoFamiliaDesencriptado, new OpcionesDeSobreescrituraDeFamiliasAnidadasYaExistentesEnElDocumento(), out familia);
                        trans.Commit();
                        if (Main.cLcsv != null)
                        {
                            Main.cLcsv.PonLog_ULMA(ULMALGFree.ACTION.BROWSER_LOAD_FAMILY, FILENAME:pathArchivoFamiliaDesencriptado, FAMILY:familia.Name, EApp:ULMALGFree.queApp.ULMASTUDIO);
                        }
                        ULMALGFree.clsBase._registraLoadInsert = true;
                    }
                }
                else
                {
                    familia = (Family)encontrado;

                    List<string> tipos = new List<string>();
                    string pathCatalogoDeTipos = System.IO.Path.ChangeExtension(_pathArchivoFamilia, "txt");
                    if (System.IO.File.Exists(pathCatalogoDeTipos))
                    {
                        System.IO.StreamReader catalogo = System.IO.File.OpenText(pathCatalogoDeTipos);
                        string linea;
                        System.Text.RegularExpressions.Regex patronABuscar = new System.Text.RegularExpressions.Regex("(^\"([^\"]+?)\",)|(^([^\",]+?),)");
                        System.Text.RegularExpressions.Match resultadoBusqueda;
                        while ((linea = catalogo.ReadLine()) != null)
                        {
                            resultadoBusqueda = patronABuscar.Match(linea);
                            if (resultadoBusqueda.Success)
                            {
                                if (resultadoBusqueda.Groups[2].Value == "")
                                {
                                    tipos.Add(resultadoBusqueda.Groups[4].Value);
                                }
                                else
                                {
                                    tipos.Add(resultadoBusqueda.Groups[2].Value);
                                }
                            }
                        }
                        if (tipos.Count > familia.GetFamilySymbolIds().Count)
                        {
                            using (Transaction trans = new Transaction(documentoActivo, "Reload family types from UCBrowser"))
                            {
                                trans.Start();
                                foreach (string tipo in tipos)
                                {
                                    //Esto del try...catch no es una solucion elegante; pero es mas efectiva que intentar cargar solo los tipos que no estan en la familia.
                                    try
                                    {
                                        documentoActivo.LoadFamilySymbol(pathArchivoFamiliaDesencriptado, tipo);
                                    }
                                    catch (Exception)
                                    {
                                        continue;
                                    }
                                }
                                trans.Commit();
                            }
                        }
                    }
                }
                if (familia != null)
                {
                    ElementId idPrimerSimboloEnLaFamilia = familia.GetFamilySymbolIds().First();
                    FamilySymbol simbolo = (FamilySymbol)familia.Document.GetElement(idPrimerSimboloEnLaFamilia);
                    UIDocument interfaceConElDocumento = new UIDocument(documentoActivo);
                    ULMALGFree.clsBase._registraLoadInsert = false;
                    interfaceConElDocumento.PostRequestForElementTypePlacement(simbolo);
                    if (Main.cLcsv != null)
                    {
                        Main.cLcsv.PonLog_ULMA(ULMALGFree.ACTION.BROWSER_INSERT_FAMILY, FILENAME:simbolo.Family.Document.PathName,FAMILY:simbolo.Name, EApp: ULMALGFree.queApp.ULMASTUDIO);
                    }
                    ULMALGFree.clsBase._registraLoadInsert = true;
                }
            }
            else
            {
                TaskDialog.Show(title: "UCBrowser", mainInstruction: "Family file not found. Please, go to settings dialog and check the family folder set in it."
                                                                     + Environment.NewLine + Environment.NewLine
                                                                     + nombreFamilia);
            }
        }


        private string _pathArchivoImagenBIG = "";
        public string pathArchivoImagenBIG { set { _pathArchivoImagenBIG = value; } }

        private void VisualizarImagenBIG()
        {
            if(System.IO.File.Exists(_pathArchivoImagenBIG))
            {
                MostrarImagen ventanaImagen = new MostrarImagen();
                ventanaImagen.putImagenAMostrar(imagenAMostrar: new System.Windows.Media.Imaging.BitmapImage(new Uri(_pathArchivoImagenBIG)), 
                                                titulo: System.IO.Path.GetFileNameWithoutExtension(_pathArchivoImagenBIG));
                ventanaImagen.Show();
            }
            else
            {
                System.Windows.MessageBox.Show(messageBoxText: "File not found:"
                                                               + Environment.NewLine + Environment.NewLine
                                                               + _pathArchivoImagenBIG,
                                               caption: "UCBrowser - BIG image viewer", 
                                               System.Windows.MessageBoxButton.OK,
                                               System.Windows.MessageBoxImage.Error);
                //TaskDialog.Show(title: "UCBrowser - BIG image viewer",
                //                mainInstruction: "File not found:"
                //                                 + Environment.NewLine + Environment.NewLine
                //                                 + _pathArchivoImagenBIG);

               //MsgBox("Utility only for .rvt files", MsgBoxStyle.Critical, "UCBrowser - BIG image viewer");
            }
        }


        public string GetName()
        {
            return "Procesador de comandos Revit para el add-in UCBrowser";
        }

    }


    class OpcionesDeSobreescrituraDeFamiliasAnidadasYaExistentesEnElDocumento : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }



}

