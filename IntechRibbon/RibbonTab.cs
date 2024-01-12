using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit;
using System.Diagnostics;
using System.IO;
using System.Windows.Media;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Events;


namespace IntechRibbon
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class RibbonTab : IExternalApplication
    {
        // ExternalCommands assembly path
        static string AddInPath = typeof(RibbonTab).Assembly.Location;
        // Button icons directory
        static string ButtonIconsFolder = Path.GetDirectoryName(AddInPath);
        // uiApplication
        static UIApplication uiApplication = null;

        #region IExternalApplication Members
        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // create customer Ribbon Items
                CreateRibbonTab(application);

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Intech Ribbon", ex.ToString());

                return Autodesk.Revit.UI.Result.Failed;
            }
        }

 
        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }
        #endregion

        private void CreateRibbonTab(UIControlledApplication application)
        {
            String tabName = "Intech Ribbon";
            application.CreateRibbonTab(tabName);

            // create a Ribbon panel which contains three stackable buttons and one single push button.
            string firstPanelName = "Export";
            
            RibbonPanel ribbonSamplePanel = application.CreateRibbonPanel(tabName, firstPanelName);

            PushButtonData b1Data = new PushButtonData("BOMExport", "BOM Export", AddInPath, "IntechRibbon.ExportSchedulesToCSV");
            b1Data.ToolTip = "Export all schedules into a single CSV file.";
            b1Data.Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "icon.png"), UriKind.Absolute)); ;
            PushButton pb1 = ribbonSamplePanel.AddItem(b1Data) as PushButton;


            ribbonSamplePanel.AddSeparator();
            PushButtonData b2Data = new PushButtonData("TigerExport", "Tiger Export", AddInPath, "IntechRibbon.TigerExport");

            b2Data.ToolTip = "Export all schedules into individual CSV files.";
            BitmapImage pb2Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "icon.png"), UriKind.Absolute));
            b2Data.Image = pb2Image;
            PushButton pb2 = ribbonSamplePanel.AddItem(b2Data) as PushButton;

    
        }


    }
}
