using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static Autodesk.Revit.DB.SpecTypeId;

namespace IntechRibbon
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]//EVERY COMMAND REQUIRES THIS!
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
            List<ViewSchedule> schedules;//list of scheuled
            List<ViewSchedule> selected = new List<ViewSchedule>();  //List of selected schedules
            Dictionary<ViewSchedule, string> displayNames = new Dictionary<ViewSchedule, string>(); //to display proper name in the checkedlistbox.

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
                result2 = selectionForm.ShowDialog(); //shows dialog selection windoe
                // Determine if there are any items checked.  
                if (selectionForm.checkedListBox.CheckedItems.Count != 0)
                {
                    // If so, loop through all checked items   

                    for (int x = 0; x < selectionForm.checkedListBox.CheckedItems.Count; x++)
                    {
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

                    //prompt user to select file save location
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog1.FilterIndex = 1; // Set the default filter to Excel files
                    saveFileDialog1.DefaultExt = "xlsx"; // Set the default extension

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        baseFolder = saveFileDialog1.FileName;


                    }
                    if (string.IsNullOrEmpty(baseFolder))
                    {
                        return Result.Cancelled;
                    }


                    try
                    {
                        //convert selected schexdule to excel
                        utility.exportSchedulesToCSV(doc, ref message, baseFolder, result2, schedules, selected);
                    }
                    catch (Exception ex)
                    {
                        // If any error, give error information and return failed
                        TaskDialog.Show("Error", ex.ToString());
                        return Autodesk.Revit.UI.Result.Failed;
                    }


                }

            }

            return Result.Succeeded;

        }        
    }


    //configure this to only export new worksheet per schedule
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

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
            List<ViewSchedule> schedules;//list of scheuled
            List<ViewSchedule> selected = new List<ViewSchedule>();  //List of selected schedules
            Dictionary<ViewSchedule, string> displayNames = new Dictionary<ViewSchedule, string>(); //to display proper name in the checkedlistbox.




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

                    //prompt user to select file save location
                    var folderBrowserDialog1 = new FolderBrowserDialog();
                    DialogResult result = folderBrowserDialog1.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        baseFolder = folderBrowserDialog1.SelectedPath;


                    }
                    if (string.IsNullOrEmpty(baseFolder))
                    {
                        return Result.Cancelled;
                    }


                    try
                    {
                        //convert selected schexdule to excel
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



        }
        public static Result tigerExport(Document doc, ref string message, string saveFolder, DialogResult result,
                                List<ViewSchedule> schedules,
                                List<ViewSchedule> selected
                               )
        {
            if (result == DialogResult.None || result == DialogResult.Retry || result == DialogResult.Cancel)
            {

                //now we have a list of selected schedules.
                //we need to export these schedules
                ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
                opt.FieldDelimiter = ","; //csv file is seperated by a comma
                var roamingApplicationPath = Environment.ExpandEnvironmentVariables("%appdata%");
                var fullPath = roamingApplicationPath + @"\Autodesk\Revit\temp";
                Directory.CreateDirectory(fullPath);

                foreach (ViewSchedule vs in selected)
                {
                    //change this to selected directory in step 1
                    vs.Export(fullPath, vs.Name + ".csv", opt);


                    string csvFileName = fullPath + @"\" + vs.Name + @".csv";
                    string excelFileName = saveFolder + @"\" + vs.Name + @".xlsx";
                    string template = fullPath + @"\\template.xlsx";
                    if (!File.Exists(template))
                    {
                        TaskDialog.Show("Error", "The template file does not exist.");
                    }
                    var format = new ExcelTextFormat();
                    format.Delimiter = ',';
                    format.TextQualifier = '"';     // format.TextQualifier = '"';
                    format.EOL = "\r";              // DEFAULT IS "\r\n";

                    string pattern = @"[\[\]:*?/\\]";
                    string worksheetName;

                    using (ExcelPackage templatePackage = new ExcelPackage(new FileInfo(template)))
                    {
                        using (ExcelPackage newPackage = new ExcelPackage())
                        {
                            worksheetName = Regex.Replace(vs.Name, pattern, "");

                            // Copy the single worksheet from the template workbook to the new workbook.
                            ExcelWorksheet templateWorksheet = templatePackage.Workbook.Worksheets[0]; // Assuming it's the first worksheet
                            ExcelWorksheet copiedWorksheet = newPackage.Workbook.Worksheets.Add(worksheetName, templateWorksheet);

                            // Define the range where you want to start loading the data (e.g., C1)
                            ExcelRangeBase startCell = copiedWorksheet.Cells["A1"];

                            // Load data from the CSV, skipping the first row and setting the second row as the column headers.
                            var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format);


                            //add's image inside the header
                            var img = copiedWorksheet.HeaderFooter.OddHeader.InsertPicture(
                                new FileInfo(fullPath + @"\\Header.png"), PictureAlignment.Centered
                                );

                            // Iterate through rows until an empty row is encountered.
                            int currentColumn = 1; // Start from the first colums
                            while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[2, currentColumn].Text))
                            {
                                currentColumn++; // Move to the next row
                            }

                            int currentRow = 4; // Start from the fourth row 
                            while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[currentRow, 1].Text))
                            {
                                currentRow++; // Move to the next row
                            }

                            System.String merger = "A1:" + IndexToColumn(currentColumn - 1) + "1"; //Format example "A1:E1"
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
                            System.String merger2 = "A2:" + IndexToColumn(currentColumn - 1) + (currentRow - 1); //get data range
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
                                copiedWorksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                            }
                            else
                            {
                                copiedWorksheet.PrinterSettings.Orientation = eOrientation.Portrait;
                            }

                            File.Delete(csvFileName); //Remove Temporary File
                            try
                            {
                                newPackage.SaveAs(new FileInfo(excelFileName));
                                //need to delete csv file.
                            }
                            catch (Exception e)
                            {
                                TaskDialog.Show("Error", e.ToString());
                            }
                        }
                    }
                }
            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
        public static Result exportSchedulesToCSV(Document doc, ref string message, string saveFolder, DialogResult result,
                                List<ViewSchedule> schedules,
                                List<ViewSchedule> selected
                               )
        {
            if (result == DialogResult.None || result == DialogResult.Retry || result == DialogResult.Cancel)
            {
                //now we have a list of selected schedules.
                //we need to export these schedules
                ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
                opt.FieldDelimiter = ","; //csv file is seperated by a comma
                var roamingApplicationPath = Environment.ExpandEnvironmentVariables("%appdata%");
                var fullPath = roamingApplicationPath + @"\Autodesk\Revit\temp";
                Directory.CreateDirectory(fullPath);
                string excelFileName = saveFolder;

                //change this to selected directory in step 1

                string template = fullPath + @"\\template.xlsx";
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
                    // Copy the single worksheet from the template workbook to the new workbook.
                    ExcelWorksheet templateWorksheet = templatePackage.Workbook.Worksheets[0]; // Assuming it's the first worksheet
                    int tableNum = 0;
                    foreach (ViewSchedule vs2 in selected)
                    {
                        using (ExcelPackage newPackage = new ExcelPackage(excelFileName))
                        {
                            
                            ExcelWorksheet copiedWorksheet = newPackage.Workbook.Worksheets.Add(vs2.Name, templateWorksheet);
                        
                            //export Schedules To CSV 
                            vs2.Export(fullPath, vs2.Name + ".csv", opt);//Temp file. This will be deleted later
                            string csvFileName = fullPath + @"\" + vs2.Name + @".csv";
                            // Copy the single worksheet from the template workbook to the new workbook.


                            // Define the range where you want to start loading the data (e.g., C1)
                            ExcelRangeBase startCell = copiedWorksheet.Cells["A1"];
                            // Load data from the CSV, skipping the first row and setting the second row as the column headers.
                            var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format);
                            ////var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Light8, firstRowIsHeader);

                            //add's image inside the header
                            var img = copiedWorksheet.HeaderFooter.OddHeader.InsertPicture(
                                new FileInfo(fullPath + @"\\Header.png"), PictureAlignment.Centered
                                );
                            // Iterate through rows until an empty row is encountered.
                            int currentColumn = 1; // Start from the first colums
                            while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[2, currentColumn].Text))
                            {
                                currentColumn++; // Move to the next row
                            }
                            int currentRow = 4; // Start from the fourth row 
                            while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[currentRow, 1].Text))
                            {
                                currentRow++; // Move to the next row
                            }

                            System.String merger = "A1:" + IndexToColumn(currentColumn - 1) + "1"; //Format example "A1:E1"
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

                            System.String merger2 = "A2:" + IndexToColumn(currentColumn - 1) + (currentRow - 1); //get data range

                            var dataRange = copiedWorksheet.Cells[merger2];
                            ExcelTable table = copiedWorksheet.Tables.Add(dataRange, "Table" + tableNum);
                            tableNum += 1;
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
                                copiedWorksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                            }
                            else
                            {
                                copiedWorksheet.PrinterSettings.Orientation = eOrientation.Portrait;
                            }

                            File.Delete(csvFileName); //Remove Temporary File
                            try
                            {
                                newPackage.Save();
                                TaskDialog.Show("Sucess", "File saved");
                            }
                            catch (Exception e)
                            {
                                TaskDialog.Show("Error", e.ToString());
                            }
                        }
                    }
                }
                Console.WriteLine("Finished!");//do a task dialog instead
                Console.ReadLine();


            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }

    }
}
