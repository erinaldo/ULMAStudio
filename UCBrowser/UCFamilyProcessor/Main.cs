using System;
using System.Linq;
using System.Collections.Generic;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace UCFamilyProcessor
{
 
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [JournalingAttribute(JournalingMode.NoCommandData)]

    public class Main : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        { 
            RibbonPanel panelBrowser = null;
            try
            {
                List<RibbonPanel> panelesExistentes = app.GetRibbonPanels(tabName: "ULMA");
                panelBrowser = panelesExistentes.FirstOrDefault<RibbonPanel>(x => x.Name.Equals("Browser"));
            }
            catch(Exception)
            {
                app.CreateRibbonTab("ULMA");
            }
            try
            {
                if (panelBrowser == null)
                {
                    panelBrowser = app.CreateRibbonPanel(panelName: "Browser", tabName: "ULMA");
                }
                else
                {
                    panelBrowser.AddSeparator();
                }

                PushButton btnUCFamilyProcessor = (PushButton)panelBrowser.AddItem(new PushButtonData(name: "UCFamilyProcessor", text: "family\nprocessor",
                                                                                                   assemblyName: System.Reflection.Assembly.GetExecutingAssembly().Location,
                                                                                                   className: "UCFamilyProcessor.UCFamilyProcessor"));
                btnUCFamilyProcessor.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                                                                                               + System.IO.Path.DirectorySeparatorChar + "UCFamilyProcessor.png"));
                return Result.Succeeded;
            }
            catch(Exception ex)
            {
                TaskDialog mensajero = new TaskDialog(title:"UCFamilyProcessor.Main.OnStartup");
                mensajero.CommonButtons = TaskDialogCommonButtons.Ok;
                mensajero.MainContent = ex.ToString() + Environment.NewLine + ex.StackTrace;
                mensajero.Show();
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

    }



    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    [JournalingAttribute(JournalingMode.NoCommandData)]

    class UCFamilyProcessor : IExternalCommand
    {
        public Result Execute(ExternalCommandData datosDeEntrada, ref string mensaje, ElementSet elementos)
        {
            try
            {
                MainWindow ventana = new MainWindow();
                ventana.DataContext = new MainViewModel(datosDeEntrada.Application);
                ventana.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog mensajero = new TaskDialog(title: "UCFamilyProcessor.Execute");
                mensajero.CommonButtons = TaskDialogCommonButtons.Ok;
                mensajero.MainContent = ex.ToString() + Environment.NewLine + ex.StackTrace;
                mensajero.Show();
                return Result.Failed;
            }
        }
    }



}

