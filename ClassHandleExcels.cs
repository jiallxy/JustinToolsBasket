using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Forms=System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using OfficeOpenXml;
using Microsoft.Office.Interop.Word;
using System.Diagnostics.Eventing.Reader;

namespace JustinToolsBasket
{
    public class ClassHandleExcels : IDisposable
    {
        #region private field
        protected Boolean _disposed;
        protected Excel.Application _xl;
        protected Excel._Workbook _wb;
        protected System.Globalization.CultureInfo _oldCI;
        protected string _excelFilePathName;
        #endregion private field

        #region public property
        public Excel.Application XL
        {
            set { _xl = value; }
            get { return _xl; }
        }
        public Excel._Workbook WB
        {
            set { _wb = value; }
            get { return _wb; }
        }//because only handle one excel file, so in class, workbook can be used as public proterty, and used by all function

        #endregion public property 

        #region Basic Excel function
        /// <summary>
        /// construction function of this class
        /// </summary>
        /// <param name="excelFilePathName"> the path name of the excel file which i want to handle</param>
        public ClassHandleExcels(string excelFilePathName, bool bvisible)
        {
            _oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            // Create a new instance of Excel from scratch
            XL = new Excel.Application();
            _xl.Visible = bvisible;
            _xl.UserControl = false;
            _excelFilePathName= excelFilePathName;
            WB = OpenWorkbookfrom(excelFilePathName);
            _disposed = false;// it is like a lock showing that the instant is not disposed
        }

        public void PrintExcel(string printer)
        {
            //Excel._Workbook wb = null;
            Excel._Worksheet sheet1;
            try
            {
                //wb = OpenWorkbookfrom (ShearListPathName);

                //sheet = (Excel._Worksheet)wb.ActiveSheet;
                sheet1 = (Excel._Worksheet)(_wb.Sheets[1]);
                //MyXL.Parent.Windows("ShearList-VBnet.xlsx").Visible = true;
                sheet1.Activate();
                //Do manipulations of your  file here.
                sheet1._PrintOut(Type.Missing, Type.Missing, 1, false, printer, false, false);

                sheet1 = null;
                //_wb.Close(false, null, null);
                //_wb = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                             + System.Environment.NewLine + "For help, please email the information of this error"
                             + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
            }
        }

        protected Excel._Workbook OpenWorkbookfrom(String fileName)
        {
            if (File.Exists(fileName))
            {
                return _xl.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);
            }
            else
            {
                MessageBox.Show("Please make sure that the Excel file, "
                                   + System.Environment.NewLine + "\"" + fileName + "\""
                                   + System.Environment.NewLine + ", exists, or find another Excel file");


                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = Path.GetDirectoryName(fileName),
                    Filter = "Excel Files (*.xls; *.xlsx)|*.xls; *.xlsx|All Files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                { fileName = openFileDialog1.FileName; }



                if (File.Exists(fileName))
                {
                    return _xl.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);
                }
                else
                {
                    //this is the create new configurefile in old vb program
                    Excel._Workbook wbook = (Excel._Workbook)(_xl.Workbooks.Add(Missing.Value));
                    wbook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookDefault,
                     null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
                     false, false, null, null, null);
                    return wbook;
                }
            }
        }
        public void CheckFileListInExcelWBIfExistInFolders(string[] CNCFolders, string sheetName, string CNCProgramFileAffix
                                                        , int FirstFileNameCellRow, int FirstFileNameCellColumn, int ResultOffset
                                                         , string CncProjectFolder, bool bClean, ref Forms.Label labelStatus
                                                         , ref Forms.Label labelPercentage, ref Forms.ProgressBar progressBar1)
        {
            //bool isMaterialSeperator = true;

            int n = FirstFileNameCellRow;
            int m = FirstFileNameCellColumn;
            int k ;
            int j;
            int i=0 ;
            string result;
            string PartName;
            try
            {
                string NoCncFileDXFFolder = CncProjectFolder + "DXF";
                if(!Directory.Exists(NoCncFileDXFFolder) )
                {
                    Directory.CreateDirectory(NoCncFileDXFFolder);
                }

                string BOMFileFolder=Path.GetDirectoryName(_excelFilePathName);

                Excel._Worksheet Sheet1 = GetWorkSheetFromSheetname(sheetName);
                if (n <= 1)
                { k = 1; }
                else
                { k = n - 1; }

                Sheet1.Cells[k, m + ResultOffset] = "In both job folder and Library";

                while (((Excel.Range)(Sheet1.Cells[k, m])).Text.ToString() != "" || ((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                {
                    k = n;
                    n++;
                }


                progressBar1.Value = 1;
                progressBar1.Maximum = n-1- FirstFileNameCellRow;
                labelPercentage.Text = "0%";
                labelStatus.Text = "Searching CNC files for drawings in BOM from libray to " + CncProjectFolder;
                n = FirstFileNameCellRow;
                if (n <= 1)
                { k = 1; }
                else
                { k = n - 1; }

                while (((Excel.Range)(Sheet1.Cells[k, m])).Text.ToString() != "" || ((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                {
                    if (((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                    {
                        if (((Excel.Range)(Sheet1.Cells[n, m])).Value2 == null)
                        {
                            PartName = "";
                        }
                        else
                        {
                            PartName = (((Excel.Range)(Sheet1.Cells[n, m])).Value2).ToString();
                        }

                        if (PartName != "")
                        {
                            result = Program.SearchOneFileIfExistInFolders(CNCFolders, PartName, CNCProgramFileAffix, CncProjectFolder, bClean);
                            //Sheet1.Cells[n, m + ResultOffset] = result;
                            if (result.ToLower() != "new")
                            {
                                labelStatus.Text = "Find and copy " + PartName + CNCProgramFileAffix + " to " + CncProjectFolder;
                                string targetResult = CncProjectFolder + PartName + CNCProgramFileAffix;
                                Sheet1.Hyperlinks.Add(Sheet1.Cells[n, m + ResultOffset], targetResult, Type.Missing, targetResult, PartName + CNCProgramFileAffix);
                                ((Excel.Range)(Sheet1.Cells[n, m + ResultOffset])).Style = "Normal";
                                i++;
                            }
                            else
                            { 
                                labelStatus.Text = "Cannot find " + PartName + CNCProgramFileAffix + " in library";
                                Sheet1.Hyperlinks.Add(Sheet1.Cells[n, m + ResultOffset], PartName + ".DXF", Type.Missing, PartName + ".DXF", result);

                                if (File.Exists(Path.Combine(BOMFileFolder, PartName + ".DXF")))
                                {
                                    // Copy the file to the destination folder
                                    File.Copy(Path.Combine(BOMFileFolder, PartName + ".DXF"), Path.Combine(NoCncFileDXFFolder, PartName + ".DXF"), true);
                                }
                                ((Excel.Range)(Sheet1.Cells[n, m + ResultOffset])).Style = "Input";
                            }
                            labelStatus.Refresh();

                            progressBar1.Value = n + 1 - FirstFileNameCellRow;
                            j = (n + 1 - FirstFileNameCellRow) * 100 / progressBar1.Maximum;
                            labelPercentage.Text = j + "%";
                            labelPercentage.Refresh();
                        }
                    }

                    k = n;
                    n++;
                }
                labelStatus.Text = "Find and copy " + i + " CNC files " + "from libray to "+ CncProjectFolder;
                labelStatus.Refresh();
                Sheet1.get_Range((Object)Sheet1.Cells[FirstFileNameCellRow - 1, m + ResultOffset-1], (Object)Sheet1.Cells[n, m + ResultOffset ]).Columns.AutoFit();

                _wb.Save();
                _xl.Workbooks.Close();
                // Close the Excel application.
                // Excel will stick around after Quit if it is not under user 
                // control and there are outstanding references. When Excel is 
                // started or attached programmatically and 
                // Application.Visible = false, Application.UserControl is false. 
                // The UserControl property can be explicitly set to True which 
                // should force the application to terminate when Quit is called, 
                // regardless of outstanding references.
                _xl.UserControl = true;
                _xl.Quit();
                _wb = null;
                _xl = null;



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");

            }
        }

        public void CheckFileListInExcelWBIfExistInJobFolder( string sheetName, string CNCProgramFileAffix
                                                , int FirstFileNameCellRow, int FirstFileNameCellColumn, int ResultOffset
                                                 , string CncProjectFolder, ref Forms.Label labelStatus
                                                 , ref Forms.Label labelPercentage, ref Forms.ProgressBar progressBar1)
        {
            int n = FirstFileNameCellRow;
            int m = FirstFileNameCellColumn;
            int k;
            int j;
            int i = 0;
            //string result;
            string PartName;
            string fileProjectFullPathName;
            try
            {
                Excel._Worksheet Sheet1 = GetWorkSheetFromSheetname(sheetName);
                if (n <= 1)
                { k = 1; }
                else
                { k = n - 1; }

                Sheet1.Cells[k, m + ResultOffset] = "InJobFolder";

                while (((Excel.Range)(Sheet1.Cells[k, m])).Text.ToString() != "" || ((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                {
                    k = n;
                    n++;
                }

                progressBar1.Value = 1;
                progressBar1.Maximum = n - 1 - FirstFileNameCellRow;
                labelPercentage.Text = "0%";
                labelStatus.Text = "Searching CNC files for drawings only in job folder: " + CncProjectFolder;
                n = FirstFileNameCellRow;
                if (n <= 1)
                { k = 1; }
                else
                { k = n - 1; }


                while (((Excel.Range)(Sheet1.Cells[k, m])).Text.ToString() != "" || ((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                {
                    if (((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                    {
                        if (((Excel.Range)(Sheet1.Cells[n, m])).Value2 == null)
                        {
                            PartName = "";
                        }
                        else
                        {
                            PartName = (((Excel.Range)(Sheet1.Cells[n, m])).Value2).ToString();
                        }

                        if (PartName != "")
                        {
                            //result = Program.SearchOneFileIfExistInFolders(CNCFolders, PartName, CNCProgramFileAffix, CncProjectFolder, bClean);

                            fileProjectFullPathName = CncProjectFolder + PartName + CNCProgramFileAffix;
                            if (File.Exists(fileProjectFullPathName))
                            {
                                Sheet1.Cells[n, m + ResultOffset] = "Exist";
                                ((Excel.Range)(Sheet1.Cells[n, m + ResultOffset])).Style = "Normal";
                                labelStatus.Text = "Find " + PartName + CNCProgramFileAffix + " in " + CncProjectFolder;
                                
                            }
                            else
                            {
                                Sheet1.Cells[n, m + ResultOffset] = "Not found";
                                ((Excel.Range)(Sheet1.Cells[n, m + ResultOffset])).Style = "Bad";
                                labelStatus.Text = "Cannot find " + PartName + CNCProgramFileAffix + " in " + CncProjectFolder;
                                i++;
                            }
                            labelStatus.Refresh();

                            progressBar1.Value = n + 1 - FirstFileNameCellRow;
                            j = (n + 1 - FirstFileNameCellRow) * 100 / progressBar1.Maximum;
                            labelPercentage.Text = j + "%";
                            labelPercentage.Refresh();
                        }
                    }

                    k = n;
                    n++;
                }
                if(i>1)
                { labelStatus.Text = i + " CNC files are not found for parts in BOM in job folder: " + CncProjectFolder; }
                else if(i==1)
                { labelStatus.Text = i + " CNC file is not found for parts in BOM in job folder: " + CncProjectFolder; }
                else
                { labelStatus.Text = "All CNC files are found for parts in BOM in job folder: " + CncProjectFolder; }
                Sheet1.get_Range((Object)Sheet1.Cells[FirstFileNameCellRow-1, m + ResultOffset], (Object)Sheet1.Cells[n, m + ResultOffset+1]).Columns.AutoFit();

                labelStatus.Refresh();


                _wb.Save();
                _xl.Workbooks.Close();
                // Close the Excel application.
                // Excel will stick around after Quit if it is not under user 
                // control and there are outstanding references. When Excel is 
                // started or attached programmatically and 
                // Application.Visible = false, Application.UserControl is false. 
                // The UserControl property can be explicitly set to True which 
                // should force the application to terminate when Quit is called, 
                // regardless of outstanding references.
                _xl.UserControl = true;
                _xl.Quit();
                _wb = null;
                _xl = null;



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");

            }
        }

        public void AddHyperlinkOfFileinExcelColumn(string SheetName, string FileExt, 
                             int iRow1, int iColumn, int iRowLast, 
                             string FileFolder
                            , ref Forms.Label labelStatus
                            , ref Forms.Label labelPercentage
                            , ref Forms.ProgressBar progressBar1)
        {
            int n = iRow1;
            int m = iColumn;
            int k ;
            string PartName;
            string sLink;
            int i = 0;
            int j;
            try
            {
                Excel._Worksheet Sheet1 = GetWorkSheetFromSheetname(SheetName);

                if (n <= 1)
                { k = 1; }
                else
                { k = n - 1; }
                progressBar1.Value = 1;
                progressBar1.Maximum = iRowLast +1- n;
                labelPercentage.Text = "0%";
                labelStatus.Text = "Add hyperlink to files in "+ SheetName;

                //while (((Excel.Range)(Sheet1.Cells[k, m])).Text.ToString() != "" || ((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                while (k < iRowLast + 1)
                {
                    if (((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                    {
                        if (((Excel.Range)(Sheet1.Cells[n, m])).Value2 == null)
                        {
                            PartName = "";
                        }
                        else
                        {
                            PartName = (((Excel.Range)(Sheet1.Cells[n, m])).Value2).ToString();
                        }

                        if (PartName != "")
                        {
                            sLink = FileFolder + PartName + FileExt;
                            if (File.Exists(sLink))
                            {
                                labelStatus.Text = "Add hyperlink to " + sLink;
                                labelStatus.Refresh();
                                Sheet1.Hyperlinks.Add(Sheet1.Cells[n, m], sLink, Type.Missing, PartName + FileExt, PartName);
                                i++;
                            }
                        }
                        else
                        {
                            labelStatus.Text = "Cannot find file for row" + n + " in: " + FileFolder;
                            labelStatus.Refresh();
                        }
                    }

                    progressBar1.Value = k + 1 - iRow1;
                    j = (k + 1 - iRow1) * 100 / progressBar1.Maximum;
                    labelPercentage.Text = j + "%";
                    labelPercentage.Refresh();
                    k = n;
                    n++;
                }

                labelStatus.Text = "Add hyperlink for " + i + " files listed in " + SheetName + " of " + this.WB.Name +" from: " + FileFolder;
                labelStatus.Refresh();

                _wb.Save();
                _xl.Workbooks.Close();
                // Close the Excel application.
                // Excel will stick around after Quit if it is not under user 
                // control and there are outstanding references. When Excel is 
                // started or attached programmatically and 
                // Application.Visible = false, Application.UserControl is false. 
                // The UserControl property can be explicitly set to True which 
                // should force the application to terminate when Quit is called, 
                // regardless of outstanding references.
                _xl.UserControl = true;
                _xl.Quit();
                _wb = null;
                _xl = null;



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");

            }
        }

        public void CreateRadanProjectCSV(string SheetNameProject, string FileExt,
                                          int iRow1, int iRowLast, int iColumnPartName, int iColumnQty, int iCoulumnMaterialRadan,
                                          string FileFolder, string CNCFileFolder,
                                          ref Forms.Label LabelStatus, ref Forms.Label LabelPercentage, ref Forms.ProgressBar ProgressBar1)
        {
            try
            {
                // Set the license context for EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // Load the Excel file
                FileInfo fileInfo = new FileInfo(_excelFilePathName);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ProgressBar1.Value = 1;
                    ProgressBar1.Maximum = iRowLast + 1 - iRow1;
                    LabelPercentage.Text = "0%";
                    LabelStatus.Text = "Create Radan Project CSV ";

                    // Get the worksheets by name
                    ExcelWorksheet worksheetProjectInfo = package.Workbook.Worksheets[SheetNameProject];

                    // Determine the output CSV file path
                    string csvFilePath = Path.ChangeExtension(_excelFilePathName, ".csv");

                    // Create a StringBuilder to store CSV content
                    StringBuilder csvContent = new StringBuilder();
                    int blankPartNameEmergeTimes=0;
                    int row;
                    int PartQty = 0;
                    // Loop through the specified rows and read data from the specified columns
                    for ( row = iRow1; row <= iRowLast; row++)
                    {
                        string drawingName = worksheetProjectInfo.Cells[row, iColumnPartName].Text; // Column H (index 8)
                        string partQty = worksheetProjectInfo.Cells[row, iColumnQty].Text; // Column Qty
                        string materialRadan = worksheetProjectInfo.Cells[row, iCoulumnMaterialRadan].Text; // Column MaterialRadan
                        string materialThickness = worksheetProjectInfo.Cells[row, iCoulumnMaterialRadan + 1].Text; // Column MaterialThickness


                        if (!string.IsNullOrEmpty(drawingName))
                        {
                            blankPartNameEmergeTimes = 0;
                            // Create the CNC file path
                            string CNCfilePath = $"{CNCFileFolder}{drawingName}{FileExt}";

                            // Append the information to the CSV content
                            csvContent.AppendLine($"{CNCfilePath},{partQty},{materialRadan},{materialThickness},mm, ,0");
                            PartQty++;
                        }
                        else
                        {
                            if (blankPartNameEmergeTimes > 1) { break; }
                            blankPartNameEmergeTimes++; 
                        }
                        
                    }
                    ProgressBar1.Value  = iRowLast + 1 - iRow1;
                    LabelPercentage.Text = "100%";
                    LabelPercentage.Refresh();
                    // Update status label
                    LabelStatus.Text = "Add " +PartQty +" *"+ FileExt+ " file Path listed in " + SheetNameProject + " of " + Path.GetFileNameWithoutExtension(_excelFilePathName) + " to the CSV file";
                    LabelStatus.Refresh();

                    // Write the CSV content to a file
                    File.WriteAllText(csvFilePath, csvContent.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }


        public void CreateFileListInaExcelFile(string SheetName, string FileExt
                     , int iRow1, int iColumn
                     , string FileFolder
                     , ref Forms.Label labelStatus
                     , ref Forms.Label labelPercentage
                     , ref Forms.ProgressBar progressBar1)
        {
            int n = iRow1;
            int m = iColumn;
            string PartName;
            int i ;
            string[] files;
            try
            {

                files = Directory.GetFiles(FileFolder, "*" + FileExt, SearchOption.AllDirectories);
                if (files.Length > 0)
                {
                    Excel._Worksheet Sheet1 = GetWorkSheetFromSheetname(SheetName);

                    progressBar1.Value = 1;
                    progressBar1.Maximum = files.Length;
                    labelPercentage.Text = "0%";
                    labelStatus.Text = "Create a " + "*." + FileExt + " file list in " + SheetName + " for all files in: " + FileFolder;

                    for (i = 0; i < files.Length; i++)
                    {
                        if (File.Exists(files[i]))
                        {
                            PartName = Path.GetFileNameWithoutExtension(files[i]);
                            labelStatus.Text = "Add " + PartName + " to the list";
                            labelStatus.Refresh();

                            Sheet1.Cells[n + i, m] = PartName;
                            Sheet1.Hyperlinks.Add(Sheet1.Cells[n + i, m], files[i], Type.Missing, PartName + FileExt, PartName);

                        }
                        progressBar1.Value = i + 1;
                        labelPercentage.Text = (i + 1) * 100 / progressBar1.Maximum + "%";
                        labelPercentage.Refresh();
                    }

                    labelStatus.Text = "Added " + progressBar1.Maximum + " *" + FileExt + " files to a new list in " + SheetName +" of "+this.WB.Name ;
                    labelStatus.Refresh();

                    _wb.Save();
                    _xl.Workbooks.Close();
                    // Close the Excel application.
                    // Excel will stick around after Quit if it is not under user 
                    // control and there are outstanding references. When Excel is 
                    // started or attached programmatically and 
                    // Application.Visible = false, Application.UserControl is false. 
                    // The UserControl property can be explicitly set to True which 
                    // should force the application to terminate when Quit is called, 
                    // regardless of outstanding references.
                    _xl.UserControl = true;
                    _xl.Quit();
                    _wb = null;
                    _xl = null;
                }
                else
                {
                    labelStatus.Text = "There is no *" + FileExt + " files in: " + FileFolder;
                    labelStatus.Refresh();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");

            }
        }

        public List<string> GetFilePathFromCellsInExcelColumn(string SheetName, string FileExt, int iRow1, int iColumn, int iRowLast, string FileFolder)
        {
            int n = iRow1;
            int m = iColumn;
            int k ;
            string PartName;
            string sLink;
            List<string> files = new List<string>();

            try
            {
                Excel._Worksheet Sheet1 = GetWorkSheetFromSheetname(SheetName);

                if (n <= 1)
                { k = 1; }
                else
                { k = n - 1; }
                //while (((Excel.Range)(Sheet1.Cells[k, m])).Text.ToString() != "" || ((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                while (k < iRowLast + 1)
                {
                    if (((Excel.Range)(Sheet1.Cells[n, m])).Text.ToString() != "")
                    {
                        if (((Excel.Range)(Sheet1.Cells[n, m])).Value2 == null)
                        {
                            PartName = "";
                        }
                        else
                        {
                            PartName = (((Excel.Range)(Sheet1.Cells[n, m])).Value2).ToString();
                        }

                        if (PartName != "")
                        {
                            sLink = FileFolder + PartName + FileExt;

                            if (File.Exists(sLink))
                            {
                                files.Add(sLink);
                            }
                        }
                    }
                    k = n;
                    n++;
                }


                _wb.Save();
                _xl.Workbooks.Close();
                // Close the Excel application.
                // Excel will stick around after Quit if it is not under user 
                // control and there are outstanding references. When Excel is 
                // started or attached programmatically and 
                // Application.Visible = false, Application.UserControl is false. 
                // The UserControl property can be explicitly set to True which 
                // should force the application to terminate when Quit is called, 
                // regardless of outstanding references.
                _xl.UserControl = true;
                _xl.Quit();
                _wb = null;
                _xl = null;

                return files;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
                return null;
            }
        }

        protected object GetValue(String range)
        {

            return _xl.Application.get_Range(range, Type.Missing).Value2;
        }



        protected Excel._Worksheet GetWorkSheetFromSheetname(string sheetName)
        {
            try
            {
                Excel._Worksheet sheet1 = null;
                //find the sheet by the sheetname
                foreach (Excel._Worksheet sheet in _wb.Worksheets)
                {
                    if (sheet.Name == sheetName)
                    {
                        sheet1 = sheet;
                        break;
                    }
                    else if (sheet.Name == "Poject Info")
                    {
                        sheet1 = sheet;
                        sheet1.Name = sheetName;
                        break;
                    }
                }

                if (sheet1 == null)
                {
                    sheet1 = (Excel._Worksheet)_wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    sheet1.Name = sheetName;
                }

                return sheet1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ");
                return null;
            }


        }
        /// <summary>
        /// obsolete. I use the first row of DataArray to save ColumnNames, So don't need this functions
        /// </summary>
        /// <param name="sheet1"></param>
        /// <param name="ColumnNames"></param>
        protected void SetWorksheetFirst2RowsTitle(Excel._Worksheet sheet1, string[] ColumnNames)
        {
            if (ColumnNames != null || sheet1 != null)
            {
                sheet1.Cells[1, 1] = sheet1.Name;

                try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                   //so I use this stucture to ignore the error.
                {
                    ((Excel.Range)(sheet1.Cells[1, 1])).Style = "Title";
                }
                catch { }

                for (int i = 0; i < ColumnNames.Length; i++)
                {
                    sheet1.Cells[2, i + 1] = ColumnNames[i];
                }

                Excel.Range header = sheet1.get_Range((Object)sheet1.Cells[2, 1], (Object)sheet1.Cells[2, ColumnNames.Length]);

                try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                   //so I use this stucture to ignore the error.
                {
                    header.Style = "Note";
                }
                catch { }

                //header = null;
            }
        }
        /// <summary>
        /// M is the first cell row number, n is the first cell's column number
        /// </summary>
        /// <param name="DataArray"></param>
        /// <param name="sheetName"></param>
        /// <param name="m"></param>
        /// <param name="n"></param>
        public void Save1DArrayToColnumnOfWorkSheetByStartCell(string[] DataArray, string sheetName, int m, int n)
        {
            try
            {
                int k;
                Excel._Worksheet sheet1 = GetWorkSheetFromSheetname(sheetName);
                if (DataArray != null && sheet1 != null)
                {
                    //clear all the old records
                    //set the format of title rows
                    ////sheet1.Cells[1, 1] = sheet1.Name;

                    ////try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                    //////so I use this stucture to ignore the error.
                    ////{
                    ////    ((Excel.Range)(sheet1.Cells[1, 1])).Style = "Title";
                    ////}
                    ////catch { }


                    ////Excel.Range header = sheet1.get_Range((Object)sheet1.Cells[2, 1], (Object)sheet1.Cells[2, 5]);

                    ////try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                    //////so I use this stucture to ignore the error.
                    ////{
                    ////    header.Style = "Input";
                    ////}
                    ////catch { }

                    ////header = null;


                    for (k = 0; k < DataArray.GetLength(0); k++)
                    {
                        sheet1.Cells[k + m, n] = DataArray[k];

                        try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                           //so I use this stucture to ignore the error.
                        {
                            switch (DataArray[k])
                            {

                                case "OK":
                                    ((Excel.Range)sheet1.Cells[k + m, n]).Style = "Good";

                                    break;
                                case "failed":
                                    ((Excel.Range)sheet1.Cells[k + m, n]).Style = "Bad";
                                    break;
                                case "-":
                                    ((Excel.Range)sheet1.Cells[k + m, n]).Style = "Neutral";
                                    break;

                            }

                        }
                        catch { }

                    }

                    sheet1.get_Range((Object)sheet1.Cells[m, n], (Object)sheet1.Cells[DataArray.GetLength(0) + m, n]).Columns.AutoFit();
                }
                _wb.Save();
                _xl.Workbooks.Close();
                // Close the Excel application.
                // Excel will stick around after Quit if it is not under user 
                // control and there are outstanding references. When Excel is 
                // started or attached programmatically and 
                // Application.Visible = false, Application.UserControl is false. 
                // The UserControl property can be explicitly set to True which 
                // should force the application to terminate when Quit is called, 
                // regardless of outstanding references.
                _xl.UserControl = true;
                _xl.Quit();
                _wb = null;
                _xl = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
            }
        }

        #endregion  Basic Excel function

        #region Function related Configuration file for models
        /// <summary>
        /// I'd better transfer array from excel file to other class
        /// it can get me more freedom to communicate
        /// before I put the ocntrols directly into excel, this restrict the using area for this function
        /// </summary>
        /// <param name="sheetName">this sheet's first 2 rows are title, the data field start from A3</param>
        /// <returns></returns>
        public string[,] ReadDataFromExcelFileToDataArrayBySheetname(string sheetName)
        {
            string[,] DataArray = null;
            Excel._Worksheet sheet1 = GetWorkSheetFromSheetname(sheetName);

            //the second sheet must be change k factors
            if (sheet1 != null)
            {
                DataArray = GetDataArrayFromWorkSheet(sheet1);
            }
            else
            {
                MessageBox.Show("There is no " + sheetName + " worksheet in workbook " + _wb.Name
                                + System.Environment.NewLine + "Please make sure the Excel file is in correct version", "Warning!!!!!!!");
            }
            return DataArray;

        }

        public void WriteDataToExcelFileFromDataArrayBySheetName(string[,] DataArray, string sheetName)
        {
            Excel._Worksheet sheet1 = GetWorkSheetFromSheetname(sheetName);
            if (sheet1 != null)
            {
                SaveDataArrayToWorkSheet(DataArray, sheet1);
            }
            else
            {
                MessageBox.Show("There is no " + sheetName + " worksheet in workbook " + _wb.Name
                                + System.Environment.NewLine + "Please make sure the shearlist file is in correct version", "Warning!!!!!!!");
            }
        }

        private void ReadDataToListControlsNameTextFromDataArrayNameText(string[,] ControlNameTextArray, IList<Forms.Control> listToPopulate)
        {
            try
            {
                if (ControlNameTextArray != null)
                {
                    for (int i = 0; i < listToPopulate.Count; i++)
                    {
                        for (int j = 0; j < ControlNameTextArray.GetLength(0); j++)
                        {
                            if (listToPopulate[i].Name == ControlNameTextArray[j, 0])
                            {

                                if ((listToPopulate[i] is System.Windows.Forms.TextBox) || (listToPopulate[i] is System.Windows.Forms.ComboBox))
                                {

                                    try
                                    {
                                        listToPopulate[i].Text = ControlNameTextArray[j, 1];

                                        if (listToPopulate[i].Enabled == false)
                                        {
                                            listToPopulate[i].BackColor = System.Drawing.Color.LightGray;

                                        }
                                    }
                                    catch { }
                                }

                                else if (listToPopulate[i] is Forms.CheckBox box)
                                {
                                    try// the reason why use
                                    {

                                        box.Checked = Program.ConvertStringToBool(ControlNameTextArray[j, 1]);
                                    }
                                    catch { }
                                }
                                break;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                             + System.Environment.NewLine + "For help, please email the information of this error"
                             + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
            }
        }


        /// <summary>
        ///in this sheet, the first 2 row are for title
        /// the data fields are from A3 to some other corner
        /// it depend on how much column and rows in the data
        /// if there is no data in the sheet, the dataarray will be null
        /// </summary>
        /// <param name="sheet1"></param>
        /// <returns></returns>
        private string[,] GetDataArrayFromWorkSheet(Excel._Worksheet sheet1)
        {
            try
            {
                string[,] DataArray = null;
                int n;//row number
                int m;//column number
                int i;
                int k;

                n = 2;
                while (((Excel.Range)(sheet1.Cells[n, 1])).Text.ToString() != "")
                { n++; }
                n -= 2;

                m = 1;
                while (((Excel.Range)(sheet1.Cells[2, m])).Text.ToString() != "")
                { m++; }
                m --;

                if (n > 0 && m > 0)
                {
                    DataArray = new string[n, m];

                    for (i = 0; i < n; i++)
                    {
                        for (k = 0; k < m; k++)
                        {
                            //DataArray[i, k] = ((Excel.Range)(sheet1.Cells[i + 3, k + 1])).Text.ToString();
                            if (((Excel.Range)(sheet1.Cells[i + 2, k + 1])).Value2 == null)
                            {
                                DataArray[i, k] = "";
                            }
                            else
                            {
                                DataArray[i, k] = (((Excel.Range)(sheet1.Cells[i + 2, k + 1])).Value2).ToString();
                            }
                        }
                    }
                }
                return DataArray;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
                return null;
            }
        }

        public string[] GetDataArrayByWorkSheetNameAndColumnNO(string sheetName, int ColumnNo, int row1)
        {
            try
            {
                Excel._Worksheet sheet1 = GetWorkSheetFromSheetname(sheetName);

                //the second sheet must be change k factors
                if (sheet1 != null)
                {

                    string[] DataArray = null;
                    int n = row1;//row number
                    int m = ColumnNo;//column number
                    int i;
                    int j = 0;

                    while (j<2)
                    {
                        n++;
                        if (((Excel.Range)(sheet1.Cells[n, m])).Text.ToString() != "")
                        {
                            
                            j = 0;
                        }
                        else
                        { j++; }                        
                    }
                    


                    n -= row1+1;
                    if (n > 0)
                    {
                        DataArray = new string[n];
                        for (i = 0; i < n; i++)
                        {
                            if (((Excel.Range)(sheet1.Cells[i + row1, m])).Value2 == null)
                            {
                                DataArray[i] = "";
                            }
                            else
                            {
                                DataArray[i] = (((Excel.Range)(sheet1.Cells[i + row1, m])).Value2).ToString();
                            }
                        }
                    }
                    return DataArray;
                }
                else
                {
                    MessageBox.Show("There is no " + sheetName + " worksheet in workbook " + _wb.Name
                                    + System.Environment.NewLine + "Please make sure the Excel file is in correct version", "Warning!!!!!!!");
                    return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
                return null;
            }
        }
        /// <summary>
        /// first row of dataArray is for title "name and text"
        /// </summary>
        /// <param name="DataArray"></param>
        /// <param name="sheet1"></param>
        private void SaveDataArrayToWorkSheet(string[,] DataArray, Excel._Worksheet sheet1)
        {
            try
            {
                int i;
                int k;
                if (DataArray != null && sheet1 != null)
                {
                    //clear all the old records
                    ((Excel.Range)(sheet1.Rows["1:200", Type.Missing])).EntireRow.Delete(Excel.XlDirection.xlUp);


                    sheet1.Cells[1, 1] = sheet1.Name;

                    try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                       //so I use this stucture to ignore the error.
                    {
                        ((Excel.Range)(sheet1.Cells[1, 1])).Style = "Title";
                    }
                    catch { }


                    for (i = 0; i < DataArray.GetLength(1); i++)
                    {
                        sheet1.Cells[2, i + 1] = DataArray[0, i];
                    }

                    Excel.Range header = sheet1.get_Range((Object)sheet1.Cells[2, 1], (Object)sheet1.Cells[2, DataArray.GetLength(1)]);

                    try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                       //so I use this stucture to ignore the error.
                    {
                        header.Style = "Note";
                    }
                    catch { }

                    header = null;

                    for (i = 1; i < DataArray.GetLength(0); i++)
                    {
                        for (k = 0; k < DataArray.GetLength(1); k++)
                        {
                            sheet1.Cells[i + 2, k + 1] = DataArray[i, k];

                            try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                               //so I use this stucture to ignore the error.
                            {
                                ((Excel.Range)(sheet1.Cells[i + 2, k + 1])).Style = "Output";
                            }
                            catch { }

                        }
                    }
                    sheet1.get_Range((Object)sheet1.Cells[2, 1], (Object)sheet1.Cells[DataArray.GetLength(0) + 2, DataArray.GetLength(1)]).Columns.AutoFit();
                }
                _wb.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                           + System.Environment.NewLine + "For help, please email the information of this error"
                                           + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
            }
        }
        /// <summary>
        /// the way to write the data is starting from 3rd row, textbox name on A column,and value on B column
        /// </summary>
        /// <param name="topControl"></param>
        public void WriteDataFromControlsOnInterfaceToConfigExcelFile(Forms.Control topControl)
        {
            IList<Forms.Control> InfoControls = new List<Forms.Control>();

            //string[,] ControlNameText = null;//first column is name, second column is text for a control

            TraverseFindallInfoControlsInParentControl(topControl, ref InfoControls);
            string[,] ControlNameTextArray ;
            //first column is name, second column is text for a control
            ControlNameTextArray = GetNameTextArrayFromInfoControlList(InfoControls);
            WriteDataToExcelFileFromDataArrayBySheetName(ControlNameTextArray, "Model Data");
        }

        public void ReadDataFromConfigExcelFileToControlsOnInterface(Forms.Control topControl)
        {
            IList<Forms.Control> InfoControls = new List<Forms.Control>();
            //string[,] ControlNameText = null;//first column is name, second column is text for a control

            TraverseFindallInfoControlsInParentControl(topControl, ref InfoControls);
            //string[,] ControlNameTextArray ;
            //first column is name, second column is text for a control
            _ = ReadDataFromExcelFileToDataArrayBySheetname("Model Data");
            ReadDataToListControlsNameTextFromDataArrayNameText(new string[InfoControls.Count, 2], InfoControls);
        }

        public string ReadValuefromSheetCell(string sheetName, int row, int column)
        {
            Excel._Worksheet Sheet1 = GetWorkSheetFromSheetname(sheetName);
            if (((Excel.Range)(Sheet1.Cells[row, column])).Value2 == null)
            {
                return "";
            }
            else
            {
                return (((Excel.Range)(Sheet1.Cells[row, column])).Value2).ToString();
            }
            
        }

        #endregion Function related Configuration file for models

        #region transfer data from interface controls to array

        /// <summary>
        /// this is for shearlist where put info by excel range name and rang value.
        /// The config type is starting from 3rd row, textbox name on A column,and value on B column
        /// </summary>
        /// <param name="topControl"></param>
        public void WriteDataFromControlsOnInterfaceToExcelFileNameRange(Forms.Control topControl)
        {
            IList<Forms.Control> InfoControls = new List<Forms.Control>();
            //string[,] ControlNameText = null;//first column is name, second column is text for a control

            TraverseFindallInfoControlsInParentControl(topControl, ref InfoControls);

            //ControlNameText = getNameTextArrayFromInfoControlList(InfoControls);
            WriteFromListControlsNameTextToExcelFileNamedRange(InfoControls);

        }

        public void ReadDataFromExcelFileNameRangeToControlsOnInterface(Forms.Control topControl)
        {
            IList<Forms.Control> InfoControls = new List<Forms. Control>();
            //string[,] ControlNameText = null;//first column is name, second column is text for a control

            TraverseFindallInfoControlsInParentControl(topControl, ref InfoControls);

            ReadFromExcelFileNamedRangeToListControlsNameText(InfoControls);

        }

        /// <summary>
        /// from classJustinSwModel to use this function
        /// </summary>
        /// <param name="bomData">bom data array from assembly drawing bom</param>
        /// <param name="BomDrwUnitSystem">what kind of unit system kg or lbs</param>
        /// <param name="QtyAssy">how many assembly</param>
        public void WriteBomDataToShearList(string[,] bomData, string BomDrwUnitSystem, int QtyAssy)
        {
            int j;
            int k;

            int m; // first dimension ubound
            int n; // second dimension ubound
            int i;
            //int r; //to remember last subtoal row number
            double DWeightSub; // sub weight lbs
            Excel._Worksheet sheet1;
            //int[] SubTotalfields ;
            //int[] SubTotalfields = new int[2] { 8, 9 };

            //m = UBound(bomData, 1)
            //n = UBound(bomData, 2) - 4 
            m = bomData.GetUpperBound(0);
            //n = bomData.GetUpperBound(1)- 7 ;
            n = 7;//in the future, if I add new columns in the bom, I need not to change this number
            Excel.Range BomSumaryRange;
            Excel.Range BomBodayRange;
            Excel.Range BOMBeforeSubTotal;
            Excel.Range BOMBeforeSort;
            Excel.Range rng;
            int[] fields = new int[] { 8, 9 };
            //because i don’t want to show code seperately, and don’t want to shown ShearX ShearY weight with 
            //either IPS or MMGS
            // i show it with the configName partId


            //Dim _xl As Application = Nothing  use the private property

            //Dim _wb As Workbook = Nothing  the private property



            try
            {

                sheet1 = (Excel._Worksheet)(_wb.Sheets[1]);
                sheet1.Activate();

                //ungroup the old group,because i need new group
                //sheet1.Rows("15:100").Select()


                ((Excel.Range)(sheet1.Rows["15:100", Type.Missing])).EntireRow.Delete(Excel.XlDirection.xlUp);




                //clear the old information
                //sheet1.Rows("15:100").Select()

                rng = (Excel.Range)sheet1.Rows.get_Range("15:100", Type.Missing);
                rng.ClearContents();
                rng.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                try//different langugauge version of excel, this funtion is easy to make a mistake. Maybe the style name is different.
                   //so I use this stucture to ignore the error.
                {
                    rng.Style = "Normal";
                }
                catch { }




                //sheet1.Rows("15:100").ClearContents();

                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                //sheet1.Rows("15:100").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;


                //rng.Font.Name = "Arial";
                //rng.Font.FontStyle = "Regular";
                //rng.Font.Size = 10;
                //rng.Font.Strikethrough = false;
                //rng.Font.Superscript = false;
                //rng.Font.Subscript = false;
                //rng.Font.OutlineFont = false;
                //rng.Font.Shadow = false;
                //rng.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;
                //rng.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;


                //write information into the shearlist

                for (j = 0; j < m + 1; j++)
                {

                    switch (BomDrwUnitSystem)
                    {
                        // The following switch section causes an error.
                        case "IPS":
                            //for BOM show inches and lbs
                            sheet1.Cells[13, 3] = "SHEAR SIZE (inch)";
                            sheet1.Cells[14, 4] = "Y";
                            sheet1.Cells[14, 9] = "Weight Lbs";
                            sheet1.Cells[14, 10] = "SQ. FT";
                            sheet1.Cells[14, 11] = "Weight Lbs";

                            for (k = 1; k <= n; k++)
                            {   //input bom information into an array for shearlist information
                                if (k == 5) // qty=qtyassy*bom(j,5)
                                { sheet1.Cells[j + 15, k + 1] = Convert.ToInt16(bomData[j, k]) * QtyAssy; }
                                else if (k == 6) //partid=name+code
                                { sheet1.Cells[j + 15, k + 2] = bomData[j, k] + bomData[j, k + 2]; }
                                else if (k == 7) //weight lbs
                                {
                                    //sheet1.Cells[j + 15, k + 4] = Convert.ToDouble(bomData[j, k]);
                                    sheet1.Cells[j + 15, k + 4] = bomData[j, k];
                                }
                                else
                                { sheet1.Cells[j + 15, k + 1] = bomData[j, k]; }
                            }
                            //sqrt ft.
                            ((Excel.Range)(sheet1.Cells[j + 15, 10])).FormulaR1C1 = "=RC[-7]*RC[-6]*RC[-4]/144";
                            //total Lbs
                            ((Excel.Range)(sheet1.Cells[j + 15, 9])).FormulaR1C1 = "=RC[2]*RC[-3]";

                            break;

                        case "MMGS":
                            // for BOM shown mm and KG
                            sheet1.Cells[13, 3] = "SHEAR SIZE (mm)";
                            sheet1.Cells[14, 4] = "Y";
                            sheet1.Cells[14, 9] = "Weight Kg";
                            sheet1.Cells[14, 10] = "SQ. M";
                            sheet1.Cells[14, 11] = "Weight Kg";

                            for (k = 1; k <= n; k++)
                            {
                                //input bom information into an array for shearlist information
                                if (k == 2) // ShearXmm
                                { sheet1.Cells[j + 15, k + 1] = bomData[j, k + 7]; }
                                else if (k == 3) // ShearYmm
                                { sheet1.Cells[j + 15, k + 1] = bomData[j, k + 7]; }
                                else if (k == 5) // qty=qtyassy*bom(j,5)
                                { sheet1.Cells[j + 15, k + 1] = Convert.ToInt16(bomData[j, k]) * QtyAssy; }
                                else if (k == 6) //partid=name+code
                                { sheet1.Cells[j + 15, k + 2] = bomData[j, k] + bomData[j, k + 2]; }
                                else if (k == 7) //weight Kg
                                { //sheet1.Cells(j + 15, k + 2) = CDbl(bomData(j, k)) * CDbl(bomData(j, k - 2)) * QtyAssy
                                  //sheet1.Cells[j + 15, k + 4] = Convert.ToDouble(bomData[j, k + 4]);
                                    sheet1.Cells[j + 15, k + 4] = bomData[j, k + 4];
                                }
                                else
                                { sheet1.Cells[j + 15, k + 1] = bomData[j, k]; }
                            }
                            //sqrt ft.
                            ((Excel.Range)(sheet1.Cells[j + 15, 10])).FormulaR1C1 = "=RC[-7]*RC[-6]*RC[-4]/1000000";
                            //total Lbs
                            ((Excel.Range)(sheet1.Cells[j + 15, 9])).FormulaR1C1 = "=RC[2]*RC[-3]";

                            break;
                    }

                }

                //arrange the the shearlist
                //sort the data
                //sheet1.Range(sheet1.Cells(15, 2), sheet1.Cells(15 + m, 10)).Select()

                BOMBeforeSort = sheet1.get_Range((Object)sheet1.Cells[15, 2], (Object)sheet1.Cells[15 + m, 11]);
                BOMBeforeSort.Sort(
                BOMBeforeSort.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending,
                Type.Missing, Type.Missing, Excel.XlSortOrder.xlAscending,
                Type.Missing, Excel.XlSortOrder.xlAscending,
                Excel.XlYesNoGuess.xlNo, Type.Missing, Type.Missing,
                Excel.XlSortOrientation.xlSortColumns,
                Excel.XlSortMethod.xlPinYin,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal);
                BOMBeforeSort = null;


                //subtotal the data
                //sheet1.Range(sheet1.Cells(14, 2), sheet1.Cells(14 + m + 1, 10)).Select()
                BOMBeforeSubTotal = sheet1.get_Range((Object)sheet1.Cells[14, 2], (Object)sheet1.Cells[14 + m + 1, 10]);

                BOMBeforeSubTotal.Subtotal(1, Excel.XlConsolidationFunction.xlSum, fields, true, false, Excel.XlSummaryRow.xlSummaryBelow);
                BOMBeforeSubTotal = null;

                //BOMBeforeSubTotal.Subtotal(GroupBy:=1, Function:=XlConsolidationFunction.xlSum, TotalList:=SubTotalfields, _
                //Replace:=True, PageBreaks:=False, SummaryBelowData:=True)



                i = 14; //for row number
                        ////Excel.Range cell = (Excel.Range)sheet1.Cells[1, 2];
                        //////get the text
                        ////string text = (string)cell.Text ;

                while (((Excel.Range)(sheet1.Cells[i, 2])).Text.ToString() != "")
                { i++; }


                i--; // i =items qty+ material qty+1
                           //i-m-14-1=material qty

                //sheet1.Range(sheet1.Cells(i, 2), sheet1.Cells(i, 10)).Select()
                ((sheet1.get_Range((Object)sheet1.Cells[i, 2], (Object)sheet1.Cells[i, 11])).EntireRow).Delete(Excel.XlDirection.xlUp);
                i--;// i =items qty+ material qty
                          // i-m-14=material qty

                //copy the subtotal material to the sumary form

                ((Excel.Range)(sheet1.Cells[i + 2, 2])).Formula = "=B14";


                //copy the subtotal wight lbs to the sumary form
                ((Excel.Range)(sheet1.Cells[i + 2, 3])).Formula = "=I14";

                //copy the subtotal Sqrt Ft.to the sumary form
                ((Excel.Range)(sheet1.Cells[i + 2, 4])).Formula = "=J14";


                k = 0; // how many subtotal items
                       //r = 16; //subtotal item row number
                DWeightSub = 0.0;
                for (j = 15; j <= i; j++)
                {

                    //string cellText = (string)(((Excel.Range )(sheet1.Cells[j, 2])).Value2) ;//when name is number, this will get an error

                    string cellText = Convert.ToString(((Excel.Range)(sheet1.Cells[j, 2])).Value2);

                    if (cellText.Substring(cellText.Length - 5).ToUpper() == "TOTAL")
                    {
                        ((Excel.Range)(sheet1.Cells[i + 3 + k, 2])).Value2 = cellText.Substring(0, cellText.Length - 6);
                        ((Excel.Range)(sheet1.Cells[i + 3 + k, 3])).Formula = ((Excel.Range)(sheet1.Cells[j, 9])).Formula;
                        ((Excel.Range)(sheet1.Cells[i + 3 + k, 4])).Formula = ((Excel.Range)(sheet1.Cells[j, 10])).Formula;
                        k++;
                        DWeightSub = 0;
                        ((Excel.Range)(sheet1.get_Range((Object)sheet1.Cells[j, 2], (Object)sheet1.Cells[j, 10]))).EntireRow.Clear();
                    }
                    else
                    { DWeightSub += Convert.ToDouble(((Excel.Range)(sheet1.Cells[j, 9])).Value2); }
                }


                //add shearlist frame border
                //i-1 because of the last subtotal should minus
                //sheet1.Range(sheet1.Cells(14, 2), sheet1.Cells(i - 1, 10)).Select()
                BomBodayRange = sheet1.get_Range((Object)sheet1.Cells[14, 2], (Object)sheet1.Cells[i - 1, 11]);
                BomBodayRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                //BomBodayRange.Borders.Weight = Excel.XlBorderWeight.xlThin; 
                //BomBodayRange.ColorIndex  = Excel.XlColorIndex.xlColorIndexAutomatic; 


                //sumary frame border
                //sheet1.Range(sheet1.Cells(i + 2, 2), sheet1.Cells(i + k + 2, 4)).Select()
                BomSumaryRange = sheet1.get_Range((Object)sheet1.Cells[i + 2, 2], (Object)sheet1.Cells[i + k + 2, 4]);

                BomSumaryRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                //autofit the data area column

                sheet1.get_Range((Object)sheet1.Cells[13, 2], (Object)sheet1.Cells[i + k + 2, 11]).Columns.AutoFit();
                //set print area
                sheet1.get_Range((Object)sheet1.Cells[4, 2], (Object)sheet1.Cells[i + k + 2, 8]).Select();
                sheet1.PageSetup.PrintArea = sheet1.get_Range((Object)sheet1.Cells[4, 2], (Object)sheet1.Cells[i + k + 2, 8])
                              .get_Address(false, false, Excel.XlReferenceStyle.xlA1, false, Type.Missing);


                _wb.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                             + System.Environment.NewLine + "For help, please email the information of this error"
                             + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
            }



        }

        /// <summary>
        /// put all the combo text check control in the top cntrol into a list
        /// </summary>
        /// <param name="ctl">top level of control</param>
        /// <returns></returns>
        private void TraverseFindallInfoControlsInParentControl(Forms.Control ctl, ref IList<Forms.Control> ControlsCombTextCheck)
        {


            foreach (Forms.Control childCtl in ctl.Controls)
            {
                if ((childCtl is System.Windows.Forms.TextBox) || (childCtl is System.Windows.Forms.ComboBox) || (childCtl is System.Windows.Forms.CheckBox))
                {
                    ControlsCombTextCheck.Add(childCtl);
                }

                else if (childCtl.Controls.Count > 0)
                {
                    TraverseFindallInfoControlsInParentControl(childCtl, ref ControlsCombTextCheck);
                }
            }
        }
        /// <summary>
        /// From all the controls, get its name and text to form a array
        /// The first row is "name,value" for title
        /// </summary>
        /// <param name="ControlsCombTextCheck"></param>
        /// <returns></returns>
        private string[,] GetNameTextArrayFromInfoControlList(IList<Forms.Control> listToPopulate)
        {
            string[,] ControlNameTextArray = new string[listToPopulate.Count + 1, 2];//first column is name, second column is text for a control
            ControlNameTextArray[0, 0] = "Control Name";
            ControlNameTextArray[0, 1] = "Control Value";
            for (int i = 0; i < listToPopulate.Count; i++)
            {
                if ((listToPopulate[i] is System.Windows.Forms.TextBox) || (listToPopulate[i] is System.Windows.Forms.ComboBox))
                {
                    ControlNameTextArray[i + 1, 0] = listToPopulate[i].Name;
                    ControlNameTextArray[i + 1, 1] = listToPopulate[i].Text.Trim();
                }

                else if (listToPopulate[i] is Forms.CheckBox box)
                {
                    ControlNameTextArray[i + 1, 0] = listToPopulate[i].Name;
                    ControlNameTextArray[i + 1, 1] = box.Checked.ToString();
                }
            }

            return ControlNameTextArray;
        }
        /// <summary>
        /// Get list of controls from interface, and find the named ranged in excel file
        /// then put the control.text into the cell with the same name of the control.name
        /// </summary>
        /// <param name="listToPopulate"></param>
        private void WriteFromListControlsNameTextToExcelFileNamedRange(IList<Forms.Control> listToPopulate)
        {
            try
            {
                Excel._Worksheet sheet1 = null;
                sheet1 = (Excel._Worksheet)(_wb.Sheets[1]);
                sheet1.Activate();

                for (int i = 0; i < listToPopulate.Count; i++)
                {
                    if ((listToPopulate[i] is System.Windows.Forms.TextBox) || (listToPopulate[i] is System.Windows.Forms.ComboBox))
                    {

                        try
                        {
                            Excel.Range rnc = sheet1.get_Range(listToPopulate[i].Name, Type.Missing);
                            rnc.Value2 = listToPopulate[i].Text.Trim();
                        }
                        catch { }
                    }

                    else if (listToPopulate[i] is Forms.CheckBox box)
                    {
                        try// the reason why use
                        {
                            Excel.Range rnc = sheet1.get_Range(listToPopulate[i].Name, Type.Missing);
                            rnc.Value2 = box.Checked.ToString();
                        }
                        catch { }

                    }
                }

                _wb.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                             + System.Environment.NewLine + "For help, please email the information of this error"
                             + System.Environment.NewLine + "to : Justin.luo@halton.com ");
            }

        }
        private void ReadFromExcelFileNamedRangeToListControlsNameText(IList<Forms.Control> listToPopulate)
        {
            try
            {
                Excel._Worksheet sheet1 = null;
                sheet1 = (Excel._Worksheet)(_wb.Sheets[1]);
                sheet1.Activate();

                for (int i = 0; i < listToPopulate.Count; i++)
                {
                    if ((listToPopulate[i] is System.Windows.Forms.TextBox) || (listToPopulate[i] is System.Windows.Forms.ComboBox))
                    {

                        try
                        {
                            Excel.Range rnc = sheet1.get_Range(listToPopulate[i].Name, Type.Missing);
                            listToPopulate[i].Text = (string)rnc.Value2;
                        }
                        catch { }
                    }

                    else if (listToPopulate[i] is Forms.CheckBox box)
                    {
                        try// the reason why use
                        {
                            Excel.Range rnc = sheet1.get_Range(listToPopulate[i].Name, Type.Missing);

                            //bool ss=(bool)rnc.Value2;

                            box.Checked = (bool)rnc.Value2;
                        }
                        catch { }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                             + System.Environment.NewLine + "For help, please email the information of this error"
                             + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
            }

        }






        #endregion transfer data from interface controls to array

        #region Idispose interface functions

        //////// Demonstrates using the resource.  
        //////// It must not be already disposed. 
        //////public void DoSomethingWithResource() {
        //////    if (_disposed)
        //////        throw new ObjectDisposedException("Resource was disposed.");

        //////    // Show the number of bytes. 
        //////    int numBytes = (int) _resource.Length;
        //////    Console.WriteLine("Number of bytes: {0}", numBytes.ToString());
        //////}


        public void Dispose()
        {
            Dispose(true);

            // Use SupressFinalize in case a subclass 
            // of this type implements a finalizer.
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {


            // If you need thread safety, use a lock around these  
            // operations, as well as in your methods that use the resource. 

            if (!_disposed)
            {
                if (disposing)
                {
                    try
                    {
                        if (_xl != null)
                        {
                            _xl.Visible = false;
                            _xl.UserControl = false;
                            _wb?.Close(false, null, null);
                            _xl.Workbooks.Close();
                            // Close the Excel application.
                            // Excel will stick around after Quit if it is not under user 
                            // control and there are outstanding references. When Excel is 
                            // started or attached programmatically and 
                            // Application.Visible = false, Application.UserControl is false. 
                            // The UserControl property can be explicitly set to True which 
                            // should force the application to terminate when Quit is called, 
                            // regardless of outstanding references.
                            _xl.UserControl = true;
                            _xl.Quit();
                        }
                    }
                    catch { }
                    // Gracefully exit out and destroy all COM objects to avoid hanging instances
                    // of Excel.exe whether our method failed or not.



                    //if (module != null) { Marshal.ReleaseComObject(module); }
                    //if (_sheet1 != null) { Marshal.ReleaseComObject(sheet1); }
                    //if (_sheet2 != null) { Marshal.ReleaseComObject(sheet2); }
                    if (_wb != null) { Marshal.ReleaseComObject(_wb); }
                    if (_xl != null) { Marshal.ReleaseComObject(_xl); }

                    //module = null;
                    //_sheet1 = null;
                    //_sheet2 = null;
                    _wb = null;
                    _xl = null;
                    GC.Collect();


                }

                System.Threading.Thread.CurrentThread.CurrentCulture = _oldCI;

                // Indicate that the instance has been disposed.
                _disposed = true;   //this lock mean that the instant is disposed, so it can not invoke dispose method
            }
        }

        #endregion Idispose interface functions

    }

}

