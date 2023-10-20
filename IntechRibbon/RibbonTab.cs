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
        //static UIApplication uiApplication = null;

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
            //remove events
            //List<RibbonPanel> myPanels = application.GetRibbonPanels();
            //Autodesk.Revit.UI.ComboBox comboboxLevel = (Autodesk.Revit.UI.ComboBox)(myPanels[0].GetItems()[2]);
            //application.ControlledApplication.DocumentCreated -= new EventHandler<
            //   Autodesk.Revit.DB.Events.DocumentCreatedEventArgs>(DocumentCreated);
            //Autodesk.Revit.UI.TextBox textBox = myPanels[0].GetItems()[5] as Autodesk.Revit.UI.TextBox;
            //textBox.EnterPressed -= new EventHandler<
            //   Autodesk.Revit.UI.Events.TextBoxEnterPressedEventArgs>(SetTextBoxValue);

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



            #region Create a SplitButton for user to create Non-Structural or Structural Wall
            SplitButtonData splitButtonData = new SplitButtonData("BOM Export", "Tiger Export");
            SplitButton splitButton = ribbonSamplePanel.AddItem(splitButtonData) as SplitButton;
            PushButton pushButton = splitButton.AddPushButton(new PushButtonData("BOMExport", "BOMExport", AddInPath, "IntechRibbon.ExportSchedulesToCSV"));
            //pushButton.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "CreateWall.png"), UriKind.Absolute));
            //pushButton.Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "CreateWall-S.png"), UriKind.Absolute)); //add intech.png instead
            //pushButton.ToolTip = "Creates a partition wall in the building model.";
            //pushButton.ToolTipImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "CreateWallTooltip.bmp"), UriKind.Absolute));
            pushButton = splitButton.AddPushButton(new PushButtonData("TigerExport", "Tiger Export", AddInPath, "IntechRibbon.ExportSchedulesToCSV"));//need to implement exportschedulrtocsv method
            //pushButton.LargeImage = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "StrcturalWall.png"), UriKind.Absolute));
            //pushButton.Image = new BitmapImage(new Uri(Path.Combine(ButtonIconsFolder, "StrcturalWall-S.png"), UriKind.Absolute));//add intech.png instead
            #endregion

            ribbonSamplePanel.AddSeparator();


         
        }

        
    }
}
