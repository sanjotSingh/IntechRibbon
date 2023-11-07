using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using System.IO;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using OfficeOpenXml;



//using Excel = Microsoft.Office.Interop.Excel; //for the excel conversion

//unused so far
using System.Text;
using System.Threading.Tasks;
using System.Windows;//s
using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices; //s
using Autodesk.Revit.DB.Structure; //s
using Autodesk.Revit.UI.Selection; //s
using System.Diagnostics;
using System.Collections;







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

            // The rest of your code for exporting schedules
        

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
                                    //if schedule name is inclHeaders the CheckedItems list, add schedule to selected list.
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

                        foreach (ViewSchedule vs in selected)
                        {
                            Directory.CreateDirectory(@"c:\\temp\\test"); //change this to selected directory in step 1
                            vs.Export(@"c:\\temp\\test", vs.Name + ".csv", opt);//instead of this, write directly to new excel file
                            
                            //ConvertToXlsx(@"c:\\temp\\test\\"+ vs.Name + ".txt", @"c:\\temp\\test\\revitexcel.xlsx");
                            string csvFileName = @"c:\\temp\\test\\" + vs.Name + ".csv";
                            string excelFileName = @"c:\\temp\\test\\workbook.xlsx";
                            string template = @"C:\\Program Files\\Autodesk\\Revit 2022\\AddIns\\IntechRibbon\\template.xlsx";
                            if (!File.Exists(template))
                            {
                                TaskDialog.Show("Error", "The template file does not exist.");
                            }
                            //string worksheetsName = "TEST";

                            bool firstRowIsHeader = false;

                            var format = new ExcelTextFormat();
                            format.Delimiter = ',';
                            format.TextQualifier = '"';
                            format.EOL = "\r";              // DEFAULT IS "\r\n";
                                                            // format.TextQualifier = '"';

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
                                    var range = copiedWorksheet.Cells[startCell.Address].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Light8, firstRowIsHeader);


                                    //add's image inside the header
                                    var img = copiedWorksheet.HeaderFooter.OddHeader.InsertPicture(
                                        new FileInfo(@"c:\\temp\\test\\Header.png"), PictureAlignment.Centered
                                        );

                                    // Iterate through rows until an empty row is encountered.
                                   // int currentColumn = 1; // Start from the fourth row
                                   // while (!string.IsNullOrWhiteSpace(copiedWorksheet.Cells[1, currentColumn].Text))
                                    //{
                                        // Access and process data in the current row.
                                        //string cellValue = copiedWorksheet.Cells[currentColumn, 1].Text;
                                        // Process other cells in the row as needed.
                                      //  currentColumn++; // Move to the next row
                                   // }
                                    
                                    
                                    //merge cells
                                    //string mergeCells = "A1:A"+ currentColumn;
                                    //copiedWorksheet.Cells[mergeCells].Merge = true;



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




                            Console.WriteLine("Finished!");
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

        //ExportSchedulesToRows(Doc);

        // return Result.Succeeded;
        public List<ViewSchedule> filterSchedules(Document doc)
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
        /*
        void ConvertToXlsx(string sourcefile, string destfile)
        {
            int i, j;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel._Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            string[] lines, cells;
            lines = File.ReadAllLines(sourcefile);
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel._Worksheet)xlWorkBook.ActiveSheet;
            for (i = 0; i < lines.Length; i++)
            {
                cells = lines[i].Split(new Char[] { '\t', ';' });
                for (j = 0; j < cells.Length; j++)
                    xlWorkSheet.Cells[i + 1, j + 1] = cells[j];
            }
            xlWorkBook.SaveAs(destfile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }
        */
        public List<List<string>> ExportSchedulesToRows(ViewSchedule schedule)
        {
            
            //ExportScheduleToRows(schedules);
            // Define the function to export a schedule to a list of rows
            
            List<List<string>> rows = new List<List<string>>();
            string header = schedule.Name;
            rows.Add(new List<string> { header });

            TableSectionData scheduleData = schedule.GetTableData().GetSectionData(SectionType.Body);

            for (int i = 0; i < scheduleData.NumberOfRows; i++)
            {
                List<string> currentRow = new List<string>();

                for (int j = 0; j < scheduleData.NumberOfColumns; j++)
                {
                    string cellValue = null;

                    CellType cellType = scheduleData.GetCellType(i, j);

                    if (cellType == CellType.Text)
                    {
                        cellValue = scheduleData.GetCellText(i, j);
                    }
                    else
                    {
                        String elementId = scheduleData.GetCellText(i, j);

                    }
                    currentRow.Add(cellValue);
                }
                rows.Add(currentRow);
            }
            return rows;
        }
    }

}

