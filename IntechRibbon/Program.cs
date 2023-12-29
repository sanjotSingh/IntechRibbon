using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace IntechRibbon
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ExportSchedulesToCSV : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            if (doc == null)
            {
                TaskDialog.Show("Error", "No active document found.");
                return Result.Failed;
            }

            string baseFolder = string.Empty;
            bool openExported = false;
            bool inclHeaders = false;
            List<ViewSchedule> schedules;//list of scheuled
            List<ViewSchedule> selected = new List<ViewSchedule>();  //List of selected schedules

            Dictionary<ViewSchedule, string> displayNames = new Dictionary<ViewSchedule, string>(); //to display proper name in the checkedlistbox.

            // Show a dialog to select the destination folder and export options
            using (TaskDialog taskDialog = new TaskDialog("Export Schedules to CSV"))
            {
                taskDialog.MainInstruction = "Select destination and export options";
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "My Desktop");
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Where Revit Model Is");
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "My Documents");
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "User Select");

                TaskDialogResult result = taskDialog.Show();

                if (result == TaskDialogResult.CommandLink1)
                {
                    baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                else if (result == TaskDialogResult.CommandLink2)
                {
                    baseFolder = Path.GetDirectoryName(doc.PathName); // Get the folder of the active Revit document
                }
                else if (result == TaskDialogResult.CommandLink3)
                {
                    baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                else if (result == TaskDialogResult.CommandLink4)
                {
                    using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
                    {
                        folderBrowser.Description = "Select a folder";
                        folderBrowser.RootFolder = Environment.SpecialFolder.Desktop;
                        DialogResult folderResult = folderBrowser.ShowDialog();

                        if (folderResult == DialogResult.OK)
                        {
                            baseFolder = folderBrowser.SelectedPath;
                        }
                        else
                        {
                            return Result.Cancelled;
                        }
                    }
                }

                openExported = result == TaskDialogResult.CommandLink1 || result == TaskDialogResult.CommandLink2;
                inclHeaders = result == TaskDialogResult.CommandLink1 || result == TaskDialogResult.CommandLink3;
            }

            if (string.IsNullOrEmpty(baseFolder))
            {
                return Result.Cancelled;
            }


            //we now have the location where the file will be saved.
            // Create a form to select schedules.
            DialogResult result2 = System.Windows.Forms.DialogResult.None;
            using (SelectionForm selectionForm = new SelectionForm())
            {
                schedules = utility.filterSchedules(doc);
                foreach (ViewSchedule w in schedules)
                {
                    selectionForm.checkedListBox.Items.Add(w.Name);
                }
                result2 = selectionForm.ShowDialog();
                // Determine if there are any items checked.  
                if (selectionForm.checkedListBox.CheckedItems.Count != 0)
                {
                    // If so, loop through all checked items   

                    for (int x = 0; x < selectionForm.checkedListBox.CheckedItems.Count; x++)
                    {
                        System.Windows.Forms.MessageBox.Show(selectionForm.checkedListBox.CheckedItems[x].ToString());
                        //compare schedule name
                        foreach (ViewSchedule w in schedules)
                        {
                            //if schedule name is in the CheckedItems list, add schedule to selected list.
                            if (w.Name.Equals(selectionForm.checkedListBox.CheckedItems[x]))
                            {
                                selected.Add(w);
                            }
                        }
                    }
                    try
                    {
                        utility.tigerExport(doc, ref message, baseFolder, result2, schedules, selected);
                    }
                    catch (Exception ex)
                    {
                        // If any error, give error information and return failed
                        message = ex.Message;
                        return Autodesk.Revit.UI.Result.Failed;
                    }


                }

            }

            return Result.Succeeded;

        }



    }


    //configure this to only export new worksheet per schedule
    public class TigerExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            if (doc == null)
            {
                TaskDialog.Show("Error", "No active document found.");
                return Result.Failed;
            }

            string baseFolder = string.Empty;
            bool openExported = false;
            bool inclHeaders = false;
            List<ViewSchedule> schedules;//list of scheuled
            List<ViewSchedule> selected = new List<ViewSchedule>();  //List of selected schedules

            Dictionary<ViewSchedule, string> displayNames = new Dictionary<ViewSchedule, string>(); //to display proper name in the checkedlistbox.

            // Show a dialog to select the destination folder and export options
            using (TaskDialog taskDialog = new TaskDialog("Export Schedules to CSV"))
            {
                taskDialog.MainInstruction = "Select destination and export options";
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "My Desktop");
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Where Revit Model Is");
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "My Downloads");
                taskDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "User Select");

                TaskDialogResult result = taskDialog.Show();

                if (result == TaskDialogResult.CommandLink1)
                {
                    baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                else if (result == TaskDialogResult.CommandLink2)
                {
                    baseFolder = Path.GetDirectoryName(doc.PathName); // Get the folder of the active Revit document
                }
                else if (result == TaskDialogResult.CommandLink3)
                {
                    baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                else if (result == TaskDialogResult.CommandLink4)
                {
                    using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
                    {
                        folderBrowser.Description = "Select a folder";
                        folderBrowser.RootFolder = Environment.SpecialFolder.Desktop;
                        DialogResult folderResult = folderBrowser.ShowDialog();

                        if (folderResult == DialogResult.OK)
                        {
                            baseFolder = folderBrowser.SelectedPath;
                        }
                        else
                        {
                            return Result.Cancelled;
                        }
                    }
                }

                openExported = result == TaskDialogResult.CommandLink1 || result == TaskDialogResult.CommandLink2;
                inclHeaders = result == TaskDialogResult.CommandLink1 || result == TaskDialogResult.CommandLink3;
            }

            if (string.IsNullOrEmpty(baseFolder))
            {
                return Result.Cancelled;
            }


            //we now have the location where the file will be saved.
            // Create a form to select schedules.
            DialogResult result2 = System.Windows.Forms.DialogResult.None;
            using (SelectionForm selectionForm = new SelectionForm())
            {
                schedules = utility.filterSchedules(doc);
                foreach (ViewSchedule w in schedules)
                {
                    selectionForm.checkedListBox.Items.Add(w.Name);
                }
                result2 = selectionForm.ShowDialog();
                // Determine if there are any items checked.  
                if (selectionForm.checkedListBox.CheckedItems.Count != 0)
                {
                    // If so, loop through all checked items   

                    for (int x = 0; x < selectionForm.checkedListBox.CheckedItems.Count; x++)
                    {
                        System.Windows.Forms.MessageBox.Show(selectionForm.checkedListBox.CheckedItems[x].ToString());
                        //compare schedule name
                        foreach (ViewSchedule w in schedules)
                        {
                            //if schedule name is in the CheckedItems list, add schedule to selected list.
                            if (w.Name.Equals(selectionForm.checkedListBox.CheckedItems[x]))
                            {
                                selected.Add(w);
                            }
                        }
                    }
                    try
                    {
                        utility.tigerExport(doc, ref message, baseFolder, result2, schedules, selected);
                    }
                    catch (Exception ex)
                    {
                        // If any error, give error information and return failed
                        message = ex.Message;
                        return Autodesk.Revit.UI.Result.Failed;
                    }


                }

            }

            return Result.Succeeded;

        }

    }
    public static class utility
    {
        public static List<ViewSchedule> filterSchedules(Document doc)
        {
            List<ViewSchedule> coll = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .ToList();

            List<ViewSchedule> keyschedules = new List<ViewSchedule>();
            List<ViewSchedule> revschedules = new List<ViewSchedule>();
            List<ViewSchedule> schedules = new List<ViewSchedule>();


            foreach (ViewSchedule s in coll)
            {
                if (s.Definition.IsKeySchedule)
                {
                    keyschedules.Add(s);
                }
                else if (s.Name.Contains("<Revision Schedule>"))
                {
                    revschedules.Add(s);
                }
                else
                {
                    schedules.Add(s);
                }
            }
            return schedules;
        }
        public static char IndexToColumn(int a)
        {
            char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
#pragma warning disable CS0168 // Variable is declared but never used
            try
            {
                return alpha[a - 1];
            }
            catch (IndexOutOfRangeException e)
            {
                //if this error appears. Change the array above to include alphabets such as "AA", "AB" and so on.
                TaskDialog.Show("Error", "This program was coded to handle 26 rows or less. Please edit source code to fix this error");
                return 'A';
            }
#pragma warning restore CS0168 // Variable is declared but never used



        }
        public static void adjustOrientation()
        {

        }
        public static Result tigerExport(Document doc, ref string message, string baseFolder, DialogResult result,
                                List<ViewSchedule> schedules,
                                List<ViewSchedule> selected
                               )
        {

            while (result == DialogResult.None || result == DialogResult.Retry)
            {
                //now we have a list of selected schedules.
                //we need to export these schedules
                ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
                opt.FieldDelimiter = ","; //csv file is seperated by a comma
                var roamingApplicationPath = Environment.ExpandEnvironmentVariables("%appdata%");
                var fullPath = roamingApplicationPath + @"\Autodesk\Revit\temp";
                Directory.CreateDirectory(fullPath);
                TaskDialog.Show("PATH", fullPath);
                foreach (ViewSchedule vs in selected)
                {
                    //change this to selected directory in step 1
                    vs.Export(fullPath, vs.Name + ".csv", opt);


                    string csvFileName = fullPath + @"\" + vs.Name + @".csv";
                    string excelFileName = baseFolder + vs.Name;
                    string template = @"C:\\Program Files\\Autodesk\\Revit 2022\\AddIns\\IntechRibbon\\template.xlsx";
                    if (!File.Exists(template))
                    {
                        TaskDialog.Show("Error", "The template file does not exist.");
                    }
                    var format = new ExcelTextFormat();
                    format.Delimiter = ',';
                    format.TextQualifier = '"';     // format.TextQualifier = '"';
                    format.EOL = "\r";              // DEFAULT IS "\r\n";


                    using (ExcelPackage templatePackage = new ExcelPackage(new FileInfo(template)))
                    {
                        using (ExcelPackage newPackage = new ExcelPackage())
                        {
                            // Copy the single worksheet from the template workbook to the new workbook.
                            ExcelWorksheet templateWorksheet = templatePackage.Workbook.Worksheets[0]; // Assuming it's the first worksheet
                            ExcelWorksheet copiedWorksheet = newPackage.Workbook.Worksheets.Add("NewWorksheetName", templateWorksheet);


                            // Define the range where you want to start loading the data (e.g., C1)
                            ExcelRangeBase startCell = copiedWorksheet.Cells["A1"];

                            // Load data from the CSV, skipping the first row and setting the second row as the column headers.
                            var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format);
                            ////var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Light8, firstRowIsHeader);


                            //add's image inside the header
                            var img = copiedWorksheet.HeaderFooter.OddHeader.InsertPicture(
                                new FileInfo(@"c:\\temp\\test\\Header.png"), PictureAlignment.Centered
                                );

                            // Iterate through rows until an empty row is encountered.
                            int currentColumn = 1; // Start from the fourth row
                            while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[2, currentColumn].Text))
                            {
                                currentColumn++; // Move to the next row
                            }

                            int currentRow = 3; // Start from the fourth row
                            while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[currentRow, 1].Text))
                            {
                                currentRow++; // Move to the next row
                            }



                            String merger = "A1:" + IndexToColumn(currentColumn - 1) + "1"; //Format example "A1:E1"
                            copiedWorksheet.Cells[merger].Merge = true;
                            // Create a new ExcelStyle object to define cell formatting.
                            ExcelStyle cellStyle = copiedWorksheet.Cells[merger].Style;

                            // Set cell properties.
                            cellStyle.Font.Bold = true;
                            cellStyle.Font.Name = "Calibri";
                            cellStyle.Font.Size = 14;
                            cellStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            cellStyle.Border.Top.Style = ExcelBorderStyle.Thin;
                            cellStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            cellStyle.Border.Left.Style = ExcelBorderStyle.Thin;
                            cellStyle.Border.Right.Style = ExcelBorderStyle.Thin;

                            //Add a table onto the data

                            String merger2 = "A2:" + IndexToColumn(currentColumn - 1) + (currentRow - 1); //get data range
                            TaskDialog.Show("Hello", merger2);
                            var dataRange = copiedWorksheet.Cells[merger2];
                            ExcelTable table = copiedWorksheet.Tables.Add(dataRange, "Table");
                            table.TableStyle = TableStyles.Light8;
                            table.ShowHeader = true;




                            //autofit cell columns
                            copiedWorksheet.Cells.AutoFitColumns();





                            // Get the height of the row in points.

                            double rowHeight = 0;
                            for (int i = currentColumn; i != 1; i--)
                            {
                                ExcelColumn row = copiedWorksheet.Column(i);
                                rowHeight += row.Width;
                            }
                            if (rowHeight < 100)
                            {
                                TaskDialog.Show("Error", rowHeight.ToString());
                            }
                            else
                                TaskDialog.Show("Error", rowHeight.ToString());




                            try
                            {
                                newPackage.SaveAs(new FileInfo(excelFileName));
                            }
                            catch (Exception e)
                            {
                                TaskDialog.Show("Error", e.ToString());
                            }
                        }
                    }
                    Console.WriteLine("Finished!");//do a task dialog instead
                    Console.ReadLine();
                }

            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
        public static Result exportSchedulesToCSV(Document doc, ref string message, string baseFolder)
        {
            List<ViewSchedule> schedules;//list of scheuled
            List<ViewSchedule> selected = new List<ViewSchedule>();  //List of selected schedules
            try
            {

                // Create a form to select objects.
                DialogResult result = System.Windows.Forms.DialogResult.None;
                while (result == DialogResult.None || result == DialogResult.Retry)
                {
                    // Show the selection form?.

                    using (SelectionForm selectionForm = new SelectionForm())
                    {
                        schedules = filterSchedules(doc);
                        foreach (ViewSchedule w in schedules)
                        {
                            selectionForm.checkedListBox.Items.Add(w.Name);
                        }
                        result = selectionForm.ShowDialog();
                        // Determine if there are any items checked.  
                        if (selectionForm.checkedListBox.CheckedItems.Count != 0)
                        {
                            // If so, loop through all checked items   

                            for (int x = 0; x < selectionForm.checkedListBox.CheckedItems.Count; x++)
                            {
                                System.Windows.Forms.MessageBox.Show(selectionForm.checkedListBox.CheckedItems[x].ToString());
                                //compare schedule name
                                foreach (ViewSchedule w in schedules)
                                {
                                    //if schedule name is in the CheckedItems list, add schedule to selected list.
                                    if (w.Name.Equals(selectionForm.checkedListBox.CheckedItems[x]))
                                    {
                                        selected.Add(w);
                                    }
                                }
                            }

                        }
                        //now we have a list of selected schedules.
                        //we need to export these schedules
                        ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
                        opt.FieldDelimiter = ","; //csv file is seperated by a comma
                        var roamingApplicationPath = Environment.ExpandEnvironmentVariables("%appdata%");
                        var fullPath = roamingApplicationPath + @"\Autodesk\Revit\temp";
                        Directory.CreateDirectory(fullPath);
                        TaskDialog.Show("PATH", fullPath);
                        foreach (ViewSchedule vs in selected)
                        {
                            //change this to selected directory in step 1
                            vs.Export(fullPath, vs.Name + ".csv", opt);


                            string csvFileName = fullPath + @"\" + vs.Name + @".csv";
                            string excelFileName = baseFolder + vs.Name;
                            string template = @"C:\\Program Files\\Autodesk\\Revit 2022\\AddIns\\IntechRibbon\\template.xlsx";
                            if (!File.Exists(template))
                            {
                                TaskDialog.Show("Error", "The template file does not exist.");
                            }
                            var format = new ExcelTextFormat();
                            format.Delimiter = ',';
                            format.TextQualifier = '"';     // format.TextQualifier = '"';
                            format.EOL = "\r";              // DEFAULT IS "\r\n";


                            using (ExcelPackage templatePackage = new ExcelPackage(new FileInfo(template)))
                            {
                                using (ExcelPackage newPackage = new ExcelPackage())
                                {
                                    // Copy the single worksheet from the template workbook to the new workbook.
                                    ExcelWorksheet templateWorksheet = templatePackage.Workbook.Worksheets[0]; // Assuming it's the first worksheet
                                    ExcelWorksheet copiedWorksheet = newPackage.Workbook.Worksheets.Add("NewWorksheetName", templateWorksheet);


                                    // Define the range where you want to start loading the data (e.g., C1)
                                    ExcelRangeBase startCell = copiedWorksheet.Cells["A1"];

                                    // Load data from the CSV, skipping the first row and setting the second row as the column headers.
                                    var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format);
                                    ////var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Light8, firstRowIsHeader);


                                    //add's image inside the header
                                    var img = copiedWorksheet.HeaderFooter.OddHeader.InsertPicture(
                                        new FileInfo(@"c:\\temp\\test\\Header.png"), PictureAlignment.Centered
                                        );

                                    // Iterate through rows until an empty row is encountered.
                                    int currentColumn = 1; // Start from the fourth row
                                    while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[2, currentColumn].Text))
                                    {
                                        currentColumn++; // Move to the next row
                                    }

                                    int currentRow = 3; // Start from the fourth row
                                    while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[currentRow, 1].Text))
                                    {
                                        currentRow++; // Move to the next row
                                    }



                                    String merger = "A1:" + IndexToColumn(currentColumn - 1) + "1"; //Format example "A1:E1"
                                    copiedWorksheet.Cells[merger].Merge = true;
                                    // Create a new ExcelStyle object to define cell formatting.
                                    ExcelStyle cellStyle = copiedWorksheet.Cells[merger].Style;

                                    // Set cell properties.
                                    cellStyle.Font.Bold = true;
                                    cellStyle.Font.Name = "Calibri";
                                    cellStyle.Font.Size = 14;
                                    cellStyle.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    cellStyle.Border.Top.Style = ExcelBorderStyle.Thin;
                                    cellStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    cellStyle.Border.Left.Style = ExcelBorderStyle.Thin;
                                    cellStyle.Border.Right.Style = ExcelBorderStyle.Thin;

                                    //Add a table onto the data

                                    String merger2 = "A2:" + IndexToColumn(currentColumn - 1) + (currentRow - 1); //get data range
                                    TaskDialog.Show("Hello", merger2);
                                    var dataRange = copiedWorksheet.Cells[merger2];
                                    ExcelTable table = copiedWorksheet.Tables.Add(dataRange, "Table");
                                    table.TableStyle = TableStyles.Light8;
                                    table.ShowHeader = true;




                                    //autofit cell columns
                                    copiedWorksheet.Cells.AutoFitColumns();





                                    // Get the height of the row in points.

                                    double rowHeight = 0;
                                    for (int i = currentColumn; i != 1; i--)
                                    {
                                        ExcelColumn row = copiedWorksheet.Column(i);
                                        rowHeight += row.Width;
                                    }
                                    if (rowHeight < 100)
                                    {
                                        TaskDialog.Show("Error", rowHeight.ToString());
                                    }
                                    else
                                        TaskDialog.Show("Error", rowHeight.ToString());




                                    try
                                    {
                                        newPackage.SaveAs(new FileInfo(excelFileName));
                                    }
                                    catch (Exception e)
                                    {
                                        TaskDialog.Show("Error", e.ToString());
                                    }
                                }
                            }
                            Console.WriteLine("Finished!");//do a task dialog instead
                            Console.ReadLine();
                        }
                    }
                }

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                // If any error, give error information and return failed
                message = ex.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
        }
    }
}
//to do: Finish tiger export button. 
