using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Printing;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Interop.Excel;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using System.Globalization;
using System.Data.Common;
using System.Text.RegularExpressions;


namespace JustinToolsBasket
{
    public partial class FormJustinTools : Form
    {
        #region private field
        private string _BomFileName;
        private string[] _CNCFolders;
        private string _BomFolderFullPath;
        private string _BomLastFolderName;
        private string _BomCSVFileName;
        private string _CombinedCSVFileName;

        private string _SheetName;
        private string _CNCProgramFileExt;
        private string _CncProjectFolder;


        private int _iRowNo1Cell;
        private int _iColumnNo1Cell;
        private int _iOffsetNoResult;
        #endregion private field

        #region public field
        public string BomFileName 
        {
            set { _BomFileName = value; }
            get { return TextExcelPathName.Text.Trim(); }
        }
        public string BomFolderFullPath
        {
            set { _BomFolderFullPath = value; }
            get { return Path.GetDirectoryName(TextExcelPathName.Text); }
        }
        public string BomLastFolderName
        {
            set { _BomLastFolderName = value; }
            get { return Path.GetFileName(BomFolderFullPath); }
        }
        public string CombinedCSVFileName
        {
            set { _CombinedCSVFileName = value; }
            get { return Path.Combine(BomFolderFullPath, BomLastFolderName + ".csv"); }
        }
        public string BomCSVFileName
        {
            set { _BomCSVFileName = value; }
            get { return Path.ChangeExtension(BomFileName, ".csv"); }
        }


        #endregion public field

        #region common functions
        private bool InitializePrivateField(object sender, EventArgs e)
        {
            bool bresult = true;
            try
            {
                _CNCFolders = GetCNCDatabaseFolders();
                if (!(_CNCFolders[0] == "" || Directory.Exists(_CNCFolders[0])))
                {
                    bresult = false;
                    MessageBox.Show(_CNCFolders[0] + " does not exist!", "Warning!!!!!!");
                }
                if (!(_CNCFolders[1] == "" || Directory.Exists(_CNCFolders[1])))
                {
                    bresult = false;
                    MessageBox.Show(_CNCFolders[1] + " does not exist!", "Warning!!!!!!");
                }

                _SheetName = TextSheetName.Text.Trim();
                if (_SheetName == "")
                { _SheetName = "Project Info"; }

                _CNCProgramFileExt = TextCNCFileExtension.Text.Trim();
                if (_CNCProgramFileExt != "")
                {
                    if (_CNCProgramFileExt.StartsWith("*"))
                    { _CNCProgramFileExt = _CNCProgramFileExt.Substring(1, _CNCProgramFileExt.Length - 1); }
                    else if (!_CNCProgramFileExt.StartsWith("."))
                    { _CNCProgramFileExt = "." + _CNCProgramFileExt; }
                }
                else
                {
                    bresult = false;
                    MessageBox.Show("No CNC program file extension is specified!", "Warning!!!!!!");
                }

                _CncProjectFolder = TextCNCProjectFolder.Text.Trim();
                if (_CncProjectFolder.Trim() != "")
                {
                    if (_CncProjectFolder.Substring(_CncProjectFolder.Length - 1) != "\\")
                    { _CncProjectFolder += "\\"; }
                }


                if (!Directory.Exists(_CncProjectFolder))
                {
                    bresult = false;
                    MessageBox.Show(_CncProjectFolder + " does not exist!", "Warning!!!!!!");
                }

                _iRowNo1Cell = (int)NumericUpDownRow1.Value;
                _iColumnNo1Cell = (int)NumericUpDownPartNoColumn.Value;
                _iOffsetNoResult = (int)NumericOffsetNo1Result.Value;

                _BomFileName = TextExcelPathName.Text.Trim();
                if (!System.IO.File.Exists(_BomFileName))
                {
                    bresult = false;
                    MessageBox.Show(_BomFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                }
                return bresult;
            }
            catch
            {
                return false;
            }

        }

        private bool InitializePrivateFieldForDelteButton(object sender, EventArgs e)
        {
            bool bresult = true;
            try
            {
                _CNCFolders = new string[2] { "", TextCNCDataBasePath1.Text.Trim() };
                _CNCFolders[0] = "";
                if (!(_CNCFolders[1] == "" || Directory.Exists(_CNCFolders[1])))
                {
                    bresult = false;
                    MessageBox.Show(_CNCFolders[1] + " does not exist!", "Warning!!!!!!");
                }

                _CNCProgramFileExt = TextCNCFileExtension.Text.Trim();
                if (_CNCProgramFileExt != "")
                {
                    if (_CNCProgramFileExt.StartsWith("*"))
                    { _CNCProgramFileExt = _CNCProgramFileExt.Substring(1, _CNCProgramFileExt.Length - 1); }
                    else if (!_CNCProgramFileExt.StartsWith("."))
                    { _CNCProgramFileExt = "." + _CNCProgramFileExt; }
                }
                else
                {
                    bresult = false;
                    MessageBox.Show("No CNC program file extension is specified!", "Warning!!!!!!");
                }
                return bresult;
            }
            catch
            {
                return false;
            }

        }

        private void DeletFilesInList(IList<string> listDeleteFiles)
        {
            string message = "";

            foreach (string DeleteFileName in listDeleteFiles)
            {
                message = message + System.Environment.NewLine + DeleteFileName;
            }

            string caption = "Delete the files and add revison numbers to them?";
            string messageContent = "Do you want to delete the files Listed in the text file (DeletFilesList.txt), and add revison numbers to them?" + System.Environment.NewLine +
                                     "" + System.Environment.NewLine +
                                     "If you want to delete, please check the files list in the text file, and remove the files from the list that you do not want to delete. Save and close the text file. Then Click Yes Button." + System.Environment.NewLine +
                                     "" + System.Environment.NewLine +
                                     "If you  do not want to delete, please close the text file and then click No Button.";

            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            string TextFilePathName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DeletFilesList.txt";

            System.IO.File.WriteAllText(TextFilePathName, message);
            //File.OpenRead (TextFilePathName );
            System.Diagnostics.Process.Start(TextFilePathName);

            DialogResult result;

            result = MessageBox.Show(messageContent, caption, buttons);


            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                listDeleteFiles.Clear();
                // delete all the files, and add revision for them.
                string[] readText = System.IO.File.ReadAllLines(TextFilePathName);

                foreach (string s in readText)
                {

                    if (s != "")
                    {
                        if (System.IO.File.Exists(s))
                        { listDeleteFiles.Add(s); }
                    }
                }

                ProgressBar1.Value = 0;
                ProgressBar1.Maximum = listDeleteFiles.Count ;
                LabelPercentage.Text = "0%";
                int i = 1;
                int j;

                foreach (string DeleteFileName in listDeleteFiles)
                {
                    Program.CreateRevsionVersionOfaFileThenDeleteIt(DeleteFileName);

                    ProgressBar1.Value = i;
                    j = i * 100 / ProgressBar1.Maximum;
                    LabelPercentage.Text = j + "%";
                    i++;
                    LabelPercentage.Refresh();
                }
                LabelStatus.Text = "Delete " + ProgressBar1.Maximum + " files " + "in Library";

                LabelStatus.Refresh();


            }
        }
        private IList<string> FindFileNameWithCertianTxtInDirectorie(string CertainTxt, string fileExtension, string directory)
        {

            bool isItARevisionFile;
            string fileName;
            IList<string> listDeleteFiles = new List<string>();
            if (directory != "")
            {

                if (directory.Substring(directory.Length - 1) != "\\")
                {
                    directory += "\\";
                }

                string[] files = Directory.GetFiles(directory, "*" + CertainTxt + "*" + fileExtension, SearchOption.AllDirectories);
                ProgressBar1.Value=0;
                ProgressBar1.Maximum = files.Length;
                LabelPercentage.Text = "0%";
                LabelStatus.Text = "Search "+ CertainTxt+ " in directory";
                int i = 1;
                int j ;
                foreach (string file in files)
                {
                    if ((file.ToUpper()).Contains(CertainTxt.ToUpper()))
                    {
                        isItARevisionFile = false;
                        fileName = Path.GetFileNameWithoutExtension(file);
                        //check if it is a revsion file. if it is, just ignore this file ,don't do anything
                        if (fileName.Length > 6)
                        {
                            if (fileName.Substring(fileName.Length - 6, 3).ToUpper() == "REV")
                            {
                                isItARevisionFile = true;
                            }
                        }
                        if (!isItARevisionFile)
                        {
                            listDeleteFiles.Add(file);
                        }
                    }
                    ProgressBar1.Value = i;
                    j = i * 100 / ProgressBar1.Maximum;
                    LabelPercentage.Text = j + "%";
                    i++;
                    LabelPercentage.Refresh();
                }
                LabelStatus.Text = "Find " + listDeleteFiles.Count +" files "+"in "+directory; 
                LabelStatus.Refresh();

            }
            return listDeleteFiles;
        }

        private bool InitializePrivateFieldForToTargetFolder(object sender, EventArgs e)
        {
            bool bresult = true;
            try
            {
                _CNCFolders = new string[2] { "", TextCNCDataBasePath1.Text.Trim() };
                _CNCFolders[0] = "";
                if (!(_CNCFolders[1] == "" || Directory.Exists(_CNCFolders[1])))
                {
                    bresult = false;
                    MessageBox.Show(_CNCFolders[1] + " does not exist!", "Warning!!!!!!");
                }

                _CNCProgramFileExt = TextCNCFileExtension.Text.Trim();
                if (_CNCProgramFileExt != "")
                {
                    if (_CNCProgramFileExt.StartsWith("*"))
                    { _CNCProgramFileExt = _CNCProgramFileExt.Substring(1, _CNCProgramFileExt.Length - 1); }
                    else if (!_CNCProgramFileExt.StartsWith("."))
                    { _CNCProgramFileExt = "." + _CNCProgramFileExt; }
                }
                else
                {
                    bresult = false;
                    MessageBox.Show("No CNC program file extension is specified!", "Warning!!!!!!");
                }

                _CncProjectFolder = TextCNCProjectFolder.Text.Trim();
                if (_CncProjectFolder.Trim() != "")
                {
                    if (_CncProjectFolder.Substring(_CncProjectFolder.Length - 1) != "\\")
                    { _CncProjectFolder += "\\"; }
                }


                if (!Directory.Exists(_CncProjectFolder))
                {
                    bresult = false;
                    MessageBox.Show(_CncProjectFolder + " does not exist!", "Warning!!!!!!");
                }


                return bresult;
            }
            catch
            {
                return false;
            }

        }

        private void MoveFilesInListToTargetFolder(IList<string> listDeleteFiles)
        {
            ProgressBar1.Value = 0;
            ProgressBar1.Maximum = listDeleteFiles.Count;
            LabelPercentage.Text = "0%";
            int i = 1;
            int j;
            string act;
            foreach (string DeleteFileName in listDeleteFiles)
            {
                Program.CopyFileToTargetFolderAndDeleteOrNot(DeleteFileName, CheckBoxDelProgram.Checked, _CncProjectFolder);

                ProgressBar1.Value = i;
                j = i * 100 / ProgressBar1.Maximum;
                LabelPercentage.Text = j + "%";
                i++;
                LabelPercentage.Refresh();
            }
            if (CheckBoxDelProgram.Checked)
            { act = "Move"; }
            else
            { act = "Copy"; }
            LabelStatus.Text = act + " " + ProgressBar1.Maximum + " files " + "from Library to " + _CncProjectFolder;

            LabelStatus.Refresh();
        }

        private string[] GetCNCDatabaseFolders()
        {
            if (TextCNCDataBasePath1.Text.Trim() != "" && TextCNCProjectFolder.Text.Trim() != "")
            {
                if ((TextCNCDataBasePath1.Text.Trim().ToUpper()).Contains(TextCNCProjectFolder.Text.Trim().ToUpper()))
                { TextCNCDataBasePath1.Text = ""; }
                else if ((TextCNCProjectFolder.Text.Trim().ToUpper()).Contains(TextCNCDataBasePath1.Text.Trim().ToUpper()))
                { TextCNCDataBasePath1.Text = ""; }
            }

            string[] CNCFolders = new string[2] { TextCNCProjectFolder.Text.Trim(),TextCNCDataBasePath1.Text.Trim() };


            if (CNCFolders[0].Trim() != "")
            {
                if (CNCFolders[0].Substring(CNCFolders[0].Length - 1) != "\\")
                { CNCFolders[0] = CNCFolders[0] + "\\"; }
            }

            if (CNCFolders[1].Trim() != "")
            {
                if (CNCFolders[1].Substring(CNCFolders[1].Length - 1) != "\\")
                { CNCFolders[1] = CNCFolders[1] + "\\"; }
            }

            return CNCFolders;

        }
        private void ReadDataFromFile(string fileName)
        {
            if (System.IO.File.Exists(fileName))
            {
                using (StreamReader reader = new StreamReader(fileName))
                {
                    while (!reader.EndOfStream)
                    {
                        // Read the control's index or name
                        string[] indexAndName = reader.ReadLine().Split(':');
                        int index = int.Parse(indexAndName[0]);
                        string name = indexAndName[1];

                        // Read the control's value or text
                        string valueOrText = reader.ReadLine();

                        // Get the control by index or name
                        Control control;
                        if (index >= 0 && index < Controls.Count)
                        {
                            control = Controls[index];
                        }
                        else
                        {
                            control = Controls.Find(name, true).FirstOrDefault();
                        }

                        // Assign the value or text to the control
                        if (control != null)
                        {
                            if (control is System.Windows.Forms.TextBox box)
                            {
                                box.Text = valueOrText;
                            }
                            else if (control is System.Windows.Forms.ComboBox box1)
                            {
                                int selectedIndex = int.Parse(valueOrText);
                                if (selectedIndex >= 0 && selectedIndex < box1.Items.Count)
                                {
                                    box1.SelectedIndex = selectedIndex;
                                }
                            }
                            else if (control is NumericUpDown down)
                            {
                                down.Value = decimal.Parse(valueOrText);
                            }
                        }
                    }
                }
            }

        }
        private void WriteDataToFile(string fileName)
        {
            using (StreamWriter writer = new StreamWriter(fileName))
            {
                SaveControlData(this, writer);
            }
        }
        private void SaveControlData(Control parent, StreamWriter writer)
        {
            foreach (Control control in parent.Controls)
            {
                if (control is System.Windows.Forms.TextBox box)
                {
                    writer.WriteLine(Controls.IndexOf(control).ToString() + ":" + control.Name);
                    writer.WriteLine(box.Text);
                }
                else if (control is System.Windows.Forms.ComboBox box1)
                {
                    writer.WriteLine(Controls.IndexOf(control).ToString() + ":" + control.Name);
                    writer.WriteLine(box1.SelectedIndex);
                }
                else if (control is NumericUpDown down)
                {
                    writer.WriteLine(Controls.IndexOf(control).ToString() + ":" + control.Name);
                    writer.WriteLine(down.Value);
                }
                else if (control is SplitContainer || control is SplitterPanel
                      || control is TableLayoutPanel 
                      || control is TabControl || control is TabPage)
                {
                    // Recursively save data for child controls
                    SaveControlData(control, writer);
                }
                // Add additional control types as needed
            }
        }


        #endregion common functions
        public FormJustinTools()
        {
            InitializeComponent();

            // Add printers to combobox
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                ComboBoxPrinter.Items.Add(printer);
            }
            // Set default printer as selected
            ComboBoxPrinter.SelectedItem = new PrinterSettings().PrinterName;
            // Read data from file
            ReadDataFromFile("data.txt");

        }



        private void ButtonOpenBom_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(TextExcelPathName.Text))
            {
                System.Diagnostics.Process.Start(TextExcelPathName.Text);
            }
            else
            {
                MessageBox.Show(TextExcelPathName.Text+" is not found.");
            }
        }

        private void ButtonSearchCopy_Click(object sender, EventArgs e)
        {
            WriteDataToFile("data.txt");
            if (InitializePrivateField(sender, e))
            {
                ClassHandleExcels FindCNCPathName = new ClassHandleExcels(_BomFileName, true);
                FindCNCPathName.CheckFileListInExcelWBIfExistInFolders(_CNCFolders, _SheetName, _CNCProgramFileExt, _iRowNo1Cell, _iColumnNo1Cell, _iOffsetNoResult+1, _CncProjectFolder, false,ref LabelStatus,ref LabelPercentage ,ref ProgressBar1 );
                FindCNCPathName.Dispose();
            }
        }

        private void ButtonFindBOM_Click(object sender, EventArgs e)
        {
            string fileName;
            OpenFileDialog openFileDialog1;
            //DialogResult result;

            openFileDialog1 = new OpenFileDialog
            {
                Title =
                "Select the BOM Excel File."
            };

            fileName = TextExcelPathName.Text.Trim();

            if (fileName != "")
            {
                if (fileName.ToUpper().Contains(".XLS"))
                {
                    openFileDialog1.InitialDirectory = Path.GetDirectoryName(fileName);
                }
                else if (fileName.EndsWith("\\"))
                {
                    if (Directory.Exists(fileName))
                    { openFileDialog1.InitialDirectory = fileName; }
                }
                else
                {
                    fileName += "\\";
                    if (Directory.Exists(fileName))
                    { openFileDialog1.InitialDirectory = fileName; }
                    else
                    {
                        fileName = fileName.Substring(fileName.Length - 2);
                        fileName = fileName.Substring(fileName.LastIndexOf("//"));
                        if (Directory.Exists(fileName))
                        { openFileDialog1.InitialDirectory = fileName; }
                    }
                }
            }
            openFileDialog1.Filter = "Excel Files (*.xls; *.xlsx)|*.xls; *.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                TextExcelPathName.Text = fileName;
            }
            else
            {
                MessageBox.Show("You haven't chosen an EXCEL file", "Warning!!!!!!!");
            }
        }

        private void ButtonFindCNCProjectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog
            {
                Description = "Select the directory that you save your CNC program files for this type of product.",

                ShowNewFolderButton = true
            };
            // Default to the My Documents folder.
            //folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyComputer;

            if (TextCNCProjectFolder.Text.Trim() == "")
            {
                folderBrowserDialog1.SelectedPath = System.Environment.SpecialFolder.MyDocuments.ToString();
                //folderBrowserDialog1.SelectedPath = "\\FSMIS520\\misfsvol1\\Operations\\ProdEng\\PTC_data\\wo";
            }
            else
            { folderBrowserDialog1.SelectedPath = TextCNCProjectFolder.Text.Trim(); }

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                TextCNCProjectFolder.Text = folderBrowserDialog1.SelectedPath;
            }
            TextCNCDataBasePath1_Leave(sender, e);
        }

        private void ButtonFindCNCPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1;

            folderBrowserDialog1 = new FolderBrowserDialog
            {
                Description = "Select the directory that you save your CNC program files for this type of product.",

                //Do not allow the user to create New files via the FolderBrowserDialog.
                //
                ShowNewFolderButton = true
            };

            // Default to the My Documents folder.
            //folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.MyComputer;

            if (TextCNCDataBasePath1.Text.Trim() == "")

            {
                folderBrowserDialog1.SelectedPath = System.Environment.SpecialFolder.MyDocuments.ToString();
                //folderBrowserDialog1.SelectedPath = "\\\\FSMIS520\\misfsvol1\\Operations\\ProdEng\\PTC_data\\";
            }
            else
            { folderBrowserDialog1.SelectedPath = TextCNCDataBasePath1.Text.Trim(); }

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {

                TextCNCDataBasePath1.Text = folderBrowserDialog1.SelectedPath;

            }
            TextCNCDataBasePath1_Leave(sender, e);
        }

        private void FormJustinTools_Load(object sender, EventArgs e)
        {
            TextCNCDataBasePath1_Leave(sender, e);
        }

        private void FormJustinTools_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            string filePath = Path.Combine(System.Windows.Forms.Application.StartupPath, "Justin Tools Basket Installation and User Guide.pdf");
            if (System.IO.File.Exists(filePath))
            {
                System.Diagnostics.Process.Start(filePath);
            }
        }

        private void ButtonDeletProgramFiles_Click(object sender, EventArgs e)
        {
            if (CheckBoxDelProgram.Checked)
            {
                if (InitializePrivateFieldForDelteButton(sender, e))
                {
                    string fileName = TextCncProgramTobeDel.Text.Trim();

                    if (TextCncProgramTobeDel.Text.Trim() != "")
                    {

                        IList<string> listDeleteFiles = FindFileNameWithCertianTxtInDirectorie(fileName, _CNCProgramFileExt, _CNCFolders[1]);

                        DeletFilesInList(listDeleteFiles);

                        //Program.DeletFilesWithTextInFolders(_CNCFolders, fileName, _CNCProgramFileExt, ref textCncProgramTobeDel);
                    }
                }
            }
        }

        private void ButtonToJobFolder_Click(object sender, EventArgs e)
        {
            if (InitializePrivateFieldForToTargetFolder(sender, e))
            {

                string fileName = TextCncProgramTobeDel.Text.Trim();
                if (TextCncProgramTobeDel.Text.Trim() != "")
                {
                    IList<string> listDeleteFiles = FindFileNameWithCertianTxtInDirectorie(fileName, _CNCProgramFileExt, _CNCFolders[1]);
                    MoveFilesInListToTargetFolder(listDeleteFiles);
                    //Program.MoveFilesWithTextInFoldersToTargetFolder(_CNCFolders, fileName, _CNCProgramFileExt, ref textCncProgramTobeDel, checkBoxDelProgram.Checked, _CncProjectFolder);
                }
            }
        }

        private void TabPageDelete_Enter(object sender, EventArgs e)
        {
            if (TextCNCDataBasePath1.Text.Trim() == "")
            { LabelStatus.Text = "Please input library path first"; }
            else
            { LabelStatus.Text = "Delete specific CNC files in Library: " + TextCNCDataBasePath1.Text; }
        }

        private void TextCncProgramTobeDel_Leave(object sender, EventArgs e)
        {
            if (TextCncProgramTobeDel.Text.Trim()=="")
            { LabelStatus.Text = "Please provide file name information for the process"; }
        }

        private void TextCNCDataBasePath1_Leave(object sender, EventArgs e)
        {
            if (TextCNCDataBasePath1.Text.Trim() == "")
            { LabelStatus.Text = "Only search CNC files in job folder: " + TextCNCProjectFolder.Text; }
            else
            { LabelStatus.Text = "Search CNC files in library: "+TextCNCDataBasePath1.Text + " and job folder: " + TextCNCProjectFolder.Text; }
        }

        private void TextCNCProjectFolder_Leave(object sender, EventArgs e)
        {
            if (TabPageSearch.Focused)
            { TextCNCDataBasePath1_Leave(sender, e); }
            else if (TabPageAddLabel.Focused)
            { TabPageAddLabel_Enter(sender, e); }
        }

        private void TabPageSearch_Enter(object sender, EventArgs e)
        {
            TextCNCDataBasePath1_Leave(sender, e);
        }

        private void TabPageAddLabel_Enter(object sender, EventArgs e)
        {
            LabelStatus.Text = "Add a lable ("+ TextProjectInfo.Text+") to all CNC files in folder: "+TextCNCProjectFolder.Text ;
        }

        private void TextProjectInfo_Leave(object sender, EventArgs e)
        {
            TabPageAddLabel_Enter(sender, e);
        }

        private void TextCNCProjectFolder_TextChanged(object sender, EventArgs e)
        {
            if (TabPageSearch.Focused)
            { TextCNCDataBasePath1_Leave(sender, e); }
            else if (TabPageAddLabel.Focused)
            { TabPageAddLabel_Enter(sender, e); }
        }

        private void ButtonAddProjectInfoToFiles_Click(object sender, EventArgs e)
        {
            WriteDataToFile("data.txt");
            string sInfor = TextProjectInfo.Text.Trim();

            _CncProjectFolder = TextCNCProjectFolder.Text.Trim();
            if (_CncProjectFolder != "")
            {
                if (_CncProjectFolder.Substring(_CncProjectFolder.Length - 1) != "\\")
                { _CncProjectFolder += "\\"; }
            }

            if (!Directory.Exists(_CncProjectFolder))
            {
                MessageBox.Show(_CncProjectFolder + " does not exist!", "Warning!!!!!!");
            }
            else
            {
                _CNCProgramFileExt = TextCNCFileExtension.Text.Trim();
                if (_CNCProgramFileExt != "")
                {
                    if (_CNCProgramFileExt.StartsWith("*"))
                    { _CNCProgramFileExt = _CNCProgramFileExt.Substring(1, _CNCProgramFileExt.Length - 1); }
                    else if (!_CNCProgramFileExt.StartsWith("."))
                    { _CNCProgramFileExt = "." + _CNCProgramFileExt; }

                    string[] fs = Directory.GetFiles(_CncProjectFolder, "*" + _CNCProgramFileExt);

                    // Set Minimum to 1 to represent the first file being copied.
                    ProgressBar1.Minimum = 0;
                    // Set Maximum to the total number of files to copy.
                    ProgressBar1.Maximum = fs.Length - 1;
                    // Set the initial value of the ProgressBar.
                    ProgressBar1.Value = 1;
                    // Set the Step property to a value of 1 to represent each file being copied.
                    ProgressBar1.Step = 1;
                    
                    for (int i = 0; i < fs.Length; i++)
                    {
                        Program.AddPrintLabelWithFileNameAndInfoTextToAFile(fs[i], sInfor);

                        ProgressBar1.PerformStep();
                        LabelPercentage.Text=(i+1)*100/fs.Length+"%";
                        LabelPercentage.Refresh();
                        LabelStatus.Text = "Added label to file: " + fs[i];
                        LabelStatus.Refresh();
                    }
                    LabelStatus.Text = "Added printing label to " + fs.Length + " files in job folder: "+TextCNCProjectFolder.Text ;
                     
                }
                else
                {
                    MessageBox.Show("No CNC program file extension is specified!", "Warning!!!!!!");
                }

            }
        }

        private void ButtonAddHyperLink_Click(object sender, EventArgs e)
        {
            string SheetName;
            string FileExt;
            string ExcelFileName;
            string FileFolder;
            int iRow1;
            int iRowLast;
            int iColumn;

            try
            {

                SheetName = TextSheetName.Text.Trim();
                if (SheetName == "")
                { SheetName = "Project Info"; }

                FileExt = TextBoxAffix.Text.Trim();
                if (FileExt != "")
                {
                    if (FileExt.StartsWith("*"))
                    { FileExt = FileExt.Substring(1, FileExt.Length - 1); }
                    else if (!FileExt.StartsWith("."))
                    { FileExt = "." + FileExt; }
                }
                else
                {
                    FileExt = ".pdf";
                }

                iRow1 = (int)NumericUpDownRow1.Value;
                iRowLast = (int)NumericUpDownRowLast.Value;
                iColumn = (int)NumericUpDownPartNoColumn.Value;

                ExcelFileName = TextExcelPathName.Text.Trim();
                if (!System.IO.File.Exists(ExcelFileName))
                {
                    MessageBox.Show(ExcelFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                    return;
                }
                else
                {
                    FileFolder = Path.GetDirectoryName(ExcelFileName);
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }

                    if (!Directory.Exists(FileFolder))
                    {
                        MessageBox.Show(FileFolder + " does not exist!", "Warning!!!!!!");
                        return;
                    }
                    else
                    {
                        ClassHandleExcels AddHyperLinkToColumnCess = new ClassHandleExcels(ExcelFileName, true);
                        AddHyperLinkToColumnCess.AddHyperlinkOfFileinExcelColumn(SheetName, FileExt, iRow1, iColumn, iRowLast
                                                      , FileFolder
                                                      , ref LabelStatus, ref LabelPercentage, ref ProgressBar1);
                        AddHyperLinkToColumnCess.Dispose();
                        AddHyperLinkToColumnCess = null;
                    }
                }



            }
            catch
            {

            }
        }
        private void ButtonCreateListinExcel_Click(object sender, EventArgs e)
        {
            string SheetName;
            string FileExt;
            string ExcelFileName;
            string FileFolder;
            int iRow1;
            //int iRowLast;
            int iColumn;

            try
            {

                SheetName = TextSheetName.Text.Trim();
                if (SheetName == "")
                { SheetName = "Sheet1"; }

                FileExt = TextBoxAffix.Text.Trim();
                if (FileExt != "")
                {
                    if (FileExt.StartsWith("*"))
                    { FileExt = FileExt.Substring(1, FileExt.Length - 1); }
                    else if (!FileExt.StartsWith("."))
                    { FileExt = "." + FileExt; }
                }
                else
                {
                    FileExt = ".pdf";
                }

                iRow1 = (int)NumericUpDownRow1.Value;
                //iRowLast = (int)NumericUpDownRowLast.Value;
                iColumn = (int)NumericUpDownPartNoColumn.Value;

                ExcelFileName = TextExcelPathName.Text.Trim();
                if (!System.IO.File.Exists(ExcelFileName))
                {
                    MessageBox.Show(ExcelFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                    return;
                }
                else
                {
                    FileFolder = Path.GetDirectoryName(ExcelFileName);
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }

                    if (!Directory.Exists(FileFolder))
                    {
                        MessageBox.Show(FileFolder + " does not exist!", "Warning!!!!!!");
                        return;
                    }
                    else
                    {
                        ClassHandleExcels CreateFileListInExcel = new ClassHandleExcels(ExcelFileName, true);
                        CreateFileListInExcel.CreateFileListInaExcelFile(SheetName, FileExt, iRow1, iColumn
                                                      , FileFolder
                                                      , ref LabelStatus, ref LabelPercentage, ref ProgressBar1);
                        CreateFileListInExcel.Dispose();
                        CreateFileListInExcel = null;
                    }
                }



            }
            catch
            {

            }
        }
        private void ButtonSearchInJobFolder_Click(object sender, EventArgs e)
        {
            if (InitializePrivateField(sender, e))
            {
                ClassHandleExcels FindCNCPathName = new ClassHandleExcels(_BomFileName, true);
                FindCNCPathName.CheckFileListInExcelWBIfExistInJobFolder(_SheetName, _CNCProgramFileExt, _iRowNo1Cell, _iColumnNo1Cell, _iOffsetNoResult, _CncProjectFolder,  ref LabelStatus, ref LabelPercentage, ref ProgressBar1);
                FindCNCPathName.Dispose();
            }
        }

        private void ButtonPrintFiles_Click(object sender, EventArgs e)
        {
            string SheetName;
            string FileExt;
            string ExcelFileName;
            string FileFolder;
            int iRow1;
            int iRowLast;
            int iColumn;
            _ = new List<string>();
            int i, j = 0;

            bool rbet;

            try
            {

                SheetName = TextSheetName.Text.Trim();
                if (SheetName == "")
                { SheetName = "Project Info"; }

                FileExt = ".pdf";

                //iRow1 = (int)NumericUpDownRow1.Value;
                //for assembly file
                iRow1 = 4;
                iRowLast = (int)NumericUpDownRowLast.Value;
                iColumn = (int)NumericUpDownPartNoColumn.Value;

                ExcelFileName = TextExcelPathName.Text.Trim();
                if (!System.IO.File.Exists(ExcelFileName))
                {
                    MessageBox.Show(ExcelFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                    return;
                }
                else
                {
                    FileFolder = Path.GetDirectoryName(ExcelFileName);
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }


                    if (!Directory.Exists(FileFolder))
                    {
                        MessageBox.Show(FileFolder + " does not exist!", "Warning!!!!!!");
                        return;
                    }
                    else
                    {
                        System.Drawing.Printing.PrinterSettings settings = new System.Drawing.Printing.PrinterSettings();
                        string defaultPrinterName = settings.PrinterName;
                        //get the file path names
                        ClassHandleExcels GetFilePathInColumn = new ClassHandleExcels(ExcelFileName, true);
                        GetFilePathInColumn.PrintExcel(defaultPrinterName);
                        List<string> files = GetFilePathInColumn.GetFilePathFromCellsInExcelColumn(SheetName, FileExt, iRow1, iColumn, iRowLast, FileFolder);
                        //get the default printer

                        //print files one by one

                        // Set Minimum to 1 to represent the first file being copied.
                        ProgressBar1.Minimum = 0;
                        // Set Maximum to the total number of files to copy.
                        ProgressBar1.Maximum = files.Count;
                        // Set the initial value of the ProgressBar.
                        ProgressBar1.Value = 1;
                        // Set the Step property to a value of 1 to represent each file being copied.
                        ProgressBar1.Step = 1;

                        for (i = 0; i < files.Count; i++)
                        {

                            rbet = PrintFile(FileExt, files[i], defaultPrinterName,CheckBoxPrint1Page.Checked);

                            if (!rbet)
                            { MessageBox.Show(files[i] + " can not be printed! Only support PDF printing", "Warning!!!!!!"); }
                            else
                            {
                                ProgressBar1.PerformStep();
                                LabelPercentage.Text = ((i+1) * 100 / files.Count) + "%";
                                LabelPercentage.Refresh();
                                LabelStatus.Text = "Printed " + files[i];
                                LabelStatus.Refresh();
                                j++;
                            }
                        }

                        LabelStatus.Text = "Printed " + j +" of "+files.Count + " PDF files listed in " + Path.GetFileName(ExcelFileName);
                        LabelStatus.Refresh();
                        GetFilePathInColumn.Dispose();
                        GetFilePathInColumn = null;
                    }
                }
            }
            catch
            {
            }
        }

        public bool PrintFile(string FileExt, string file, string printer, bool onlyPrint1page)
        {
            try
            {
                if (FileExt.ToUpper() == ".PDF")
                {
                    if (onlyPrint1page) 
                    { PrintPDFWithoutDXFBOM(file, printer, onlyPrint1page); }
                    else
                    { PrintPDF(file, printer); }
                    
                }
                else
                {
                    return false;
                }
                return true;
            }
            catch
            { return false; }
        }

        public void PrintPDF(string file, string defaultPrinterName)
        {
            if (!string.IsNullOrEmpty(file))
            {
                try
                {
                    ProcessStartInfo info = new ProcessStartInfo
                    {
                        Verb = "print",
                        FileName = file,
                        Arguments = "\"" + defaultPrinterName + "\" /p ",
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    };
                    Process.Start(info);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error printing file: " + ex.Message);
                }
            }
        }

        public void PrintPDFWithoutDXFBOM(string file, string printer, bool onlyPrint1page)
        {
            //Create a PdfDocument object
            Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument();
            //Load a PDF file
            doc.LoadFromFile(file);
            //Specify printer name
            doc.PrintSettings.PrinterName = printer;
            //Select a page range to print
            if(onlyPrint1page)
            { doc.PrintSettings.SelectPageRange(1, 1); }
            else 
            { //doc.PrintSettings.SelectPageRange(1, 2);
            }    
            
            //Select discontinuous pages to print
            //doc.PrintSettings.SelectSomePages(new int[] { 1, 3, 5, 7 });
            //Print document
            doc.Print();
        }

        private void TabPageCNC_Enter(object sender, EventArgs e)
        {
            LabelStatus.Text = "Search CNC files listed in " + Path.GetFileName(TextExcelPathName.Text)+"; Copy to: "+TextCNCProjectFolder.Text; 
        }

        private void TabPagePDF_Enter(object sender, EventArgs e)
        {
            LabelStatus.Text = "Handle PDF files listed in " + Path.GetFileName(TextExcelPathName.Text);
        }

        private void TabPageHyperlink_Enter(object sender, EventArgs e)
        {
            LabelStatus.Text = "Add hyperlink for files listed in " +TextSheetName.Text+" of "+ Path.GetFileName(TextExcelPathName.Text)+", or create a new list in it.";
        }

        private void TextBoxAffix_TextChanged(object sender, EventArgs e)
        {
            //textBoxAffix.AutoCompleteMode = AutoCompleteMode.Suggest;
            //textBoxAffix.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //textBoxAffix.AutoCompleteCustomSource.AddRange(new string[] { "PDF", "DXF", "DWG" ,"EDRW","JPG","DOCX","PNG","XLS","XLSX"});
        }

        private void TextCNCFileExtension_TextChanged(object sender, EventArgs e)
        {

        }

        private void ButtonConvertToPDF_Click(object sender, EventArgs e)
        {
            string ExcelFileName;
            string FileFolder;

            try
            {

                ExcelFileName = TextExcelPathName.Text.Trim();
                if (System.IO.File.Exists(ExcelFileName))
                {
                    FileFolder = Path.GetDirectoryName(ExcelFileName);
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }
                }
                else if (System.IO.Directory.Exists(ExcelFileName))
                {
                    FileFolder = ExcelFileName;
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }
                }
                else
                {
                    MessageBox.Show(ExcelFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                    return;
                }

                if (!Directory.Exists(FileFolder))
                {
                    MessageBox.Show(FileFolder + " does not exist!", "Warning!!!!!!");
                    return;
                }
                else
                {
                    // Get the selected items from the ListBox
                    string[] selectedFileTypes = listBoxDocType.SelectedItems.Cast<string>().ToArray();

                    if (selectedFileTypes.Length > 0)
                    {
                        foreach (string selectedFileType in selectedFileTypes)
                        {
                            switch (selectedFileType)
                            {
                                case "Word Document":
                                    // Convert .docx files to .pdf
                                    ConvertWordToPdf(FileFolder);
                                    break;
                                case "Excel Document":
                                    // Convert .doc files to .pdf
                                    ConvertExcelToPdf(FileFolder);
                                    break;
                                case "AutoCAD Drawing":
                                    // Convert .dwg files to .pdf
                                    MessageBox.Show("DWG file is not supported now. I will work on it later :)");
                                    //ConvertDwgToPdf(FileFolder);
                                    break;
                                default:
                                    MessageBox.Show("Unsupported file type: " + selectedFileType);
                                    break;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select one or more file types.");
                    }
                }

            }
            catch
            {

            }
        }

        private void ConvertWordToPdf(string FileFolder)
        {
            // Create a Word application object
            WordApplication wordApp = new WordApplication();
            if (wordApp != null)
            {
                // Get all .docx and .doc files in the input folder
                string[] docxFiles = Directory.GetFiles(FileFolder, "*.doc");

                int i = 0;
                int Lg = docxFiles.Length;

                if (Lg > 0)
                {
                    ProgressBar1.Value = 1;
                    ProgressBar1.Maximum = Lg;
                    LabelPercentage.Text = "0%";
                    LabelStatus.Text = "Convert all Word files in: " + FileFolder + " to PDF files.";

                    foreach (string docxFile in docxFiles)
                    {
                        // Open the Word document
                        Document doc = wordApp.Documents.Open(docxFile);

                        // Construct the output PDF file path
                        string pdfFile = Path.Combine(FileFolder, Path.GetFileNameWithoutExtension(docxFile) + ".pdf");

                        // Save as PDF
                        doc.SaveAs(pdfFile, WdSaveFormat.wdFormatPDF);
                        LabelStatus.Text = "Converting " + docxFile + " to PDF";
                        LabelStatus.Refresh();
                        // Close the document
                        doc.Close();
                        i++;
                        ProgressBar1.Value = i;
                        LabelPercentage.Text = (i) * 100 / ProgressBar1.Maximum + "%";
                        LabelPercentage.Refresh();

                    }
                    LabelStatus.Text = "Converted " + Lg + " Word Files to PDF";
                    LabelStatus.Refresh();
                }

                // Quit Word application
                wordApp.Quit();
            }

        }

        private void ConvertExcelToPdf(string FileFolder)
        {
            // Create a Word application object
            ExcelApplication ExcelApp = new ExcelApplication();
            if (ExcelApp != null )
            {
                // Get all .xlsx and .xls files in the input folder
                string[] excelFiles = Directory.GetFiles(FileFolder, "*.xls");

                int i = 0;
                int totalFiles = excelFiles.Length;

                if (totalFiles > 0)
                {
                    ProgressBar1.Value = 1;
                    ProgressBar1.Maximum = totalFiles;
                    LabelPercentage.Text = "0%";
                    LabelStatus.Text = "Converting all Excel files in: " + FileFolder + " to PDF files.";

                    foreach (string excelFile in excelFiles)
                    {
                        // Load the Excel workbook
                        // Open the Excel workbook
                        Workbook workbook = ExcelApp.Workbooks.Open(excelFile);

                        // Construct the output PDF file path
                        string pdfFile = Path.Combine(FileFolder, Path.GetFileNameWithoutExtension(excelFile) + ".pdf");

                        // Save as PDF
                        workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfFile);
                        LabelStatus.Text = "Converting " + excelFile + " to PDF";
                        LabelStatus.Refresh();
                        i++;
                        ProgressBar1.Value = i;
                        LabelPercentage.Text = (i * 100 / ProgressBar1.Maximum) + "%";
                        LabelPercentage.Refresh();
                    }

                    LabelStatus.Text = "Converted " + totalFiles + " Excel Files to PDF";
                    LabelStatus.Refresh();
                }
                ExcelApp.Quit();
            }
        }


        private void LoadCsvFiles()
        {
            string folderPath = Path.GetDirectoryName(TextExcelPathName.Text);
            if (folderPath.Trim() != "")
            {
                if (folderPath.Substring(folderPath.Length - 1) != "\\")
                { folderPath += "\\"; }
            }

            if (Directory.Exists(folderPath))
            {
                string[] csvFiles = Directory.GetFiles(folderPath, "*.csv");
                ListBoxFiles.Items.Clear();

                foreach (string file in csvFiles)
                {
                    if ( file != CombinedCSVFileName)
                    {
                        ListBoxFiles.Items.Add(file);
                    }
                }
                //ListBoxFiles.Items.AddRange(csvFiles);
            }
        }

        private void TabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage == tabPageUKApp)
            {
                LoadCsvFiles();
            }
        }


        private void CombineCsvFiles()
        {
            Dictionary<string, (int quantity, string itemType, string unitName)> aggregatedData = new Dictionary<string, (int, string, string)>();
            ProgressBar1.Maximum = ListBoxFiles.Items.Count;
            int i = 0;
            foreach (string filePath in ListBoxFiles.Items)
            {
                if (filePath != CombinedCSVFileName)
                {
                    using (var reader = new StreamReader(filePath))
                    {
                        //var header = reader.ReadLine(); // Read header line

                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(',');

                            string itemName = values[0].Trim();
                            int quantity = int.Parse(values[1].Trim(), CultureInfo.InvariantCulture);
                            string itemType = values[2].Trim();
                            string unitName = values[4].Trim();

                            if (aggregatedData.ContainsKey(itemName))
                            {
                                aggregatedData[itemName] = (aggregatedData[itemName].quantity + quantity, itemType, unitName);
                            }
                            else
                            {
                                aggregatedData[itemName] = (quantity, itemType, unitName);
                            }
                        }
                    }
                }
                i++;
                ProgressBar1.Value = i;
            }
            
            //string folderPath = Path.GetDirectoryName(TextExcelPathName.Text);
            //string lastFolderName = Path.GetFileName(folderPath);
            //string outputFilePath = Path.Combine(folderPath, lastFolderName + ".csv");
            using (var writer = new StreamWriter(CombinedCSVFileName))
            {
                //writer.WriteLine("ItemName,Quantity,ItemType,,UnitName,");
                foreach (var kvp in aggregatedData)
                {
                    writer.WriteLine($"{kvp.Key},{kvp.Value.quantity},{kvp.Value.itemType},{kvp.Value.unitName},,0");
                }
            }
            LabelPercentage.Text = "100%";
            LabelStatus.Text = "Combine " + i + " CSV files to " + CombinedCSVFileName;
            //MessageBox.Show("Combined CSV file created successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ButtonCreateStickerWordDoc_Click(object sender, EventArgs e)
        {
            try
            {

                int i;
                int q;
                int j;
                // Get user inputs
                int rows = (int)NumericUpDownStickerRows.Value;
                int columns = (int)NumericUpDownStickerColumns.Value;
                string paperSize = ComboBoxPaperSize.SelectedItem.ToString();


                string SheetNameProject = TextSheetName.Text.Trim();
                if (SheetNameProject == "")
                { SheetNameProject = "Project Info"; }

                int iRow1 = (int)NumericUpDownRow1.Value;
                int iRowLast = (int)NumericUpDownRowLast.Value;


                string ExcelFileName = TextExcelPathName.Text.Trim();
                if (!System.IO.File.Exists(ExcelFileName))
                {
                    MessageBox.Show(ExcelFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                    return;
                }
                else
                {
                    string FileFolder = Path.GetDirectoryName(ExcelFileName);
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }

                    ClassHandleExcels CreateStickerWordDoc = new ClassHandleExcels(ExcelFileName, true);
                    string ProjectName = CreateStickerWordDoc.ReadValuefromSheetCell(SheetNameProject, 4, 3);
                    string SerialNo = CreateStickerWordDoc.ReadValuefromSheetCell(SheetNameProject, 6, 3);
                    string[] sticker=CreateStickerWordDoc.GetDataArrayByWorkSheetNameAndColumnNO(TextSheetName.Text.Trim(), (int)NumericUpDownPartNoColumn.Value, iRow1);
                    string[] qty=CreateStickerWordDoc.GetDataArrayByWorkSheetNameAndColumnNO(TextSheetName.Text.Trim(),(int)NumericUpDownPartQty.Value, iRow1);

                    List<string> stickers = new List<string>();
                    for (i = 0; i < qty.Length; i++)
                    {
                        if (qty[i] != "")
                        {
                            q = Convert.ToInt32(qty[i]);
                            if(q>0)
                            {
                                for (j = 0; j < q; j++)
                                { 
                                    if(RadioButtonPartInfo.Checked)
                                    {stickers.Add(sticker[i] + "\n" + SerialNo);}
                                    else
                                    {stickers.Add(ProjectName + "\n" + SerialNo);}
                                }
                            }
                        }
                    }

                    ProgressBar1.Value = 1;
                    ProgressBar1.Maximum =stickers.Count;
                    LabelPercentage.Text = "0%";
                    LabelStatus.Text = "Creating "+ Path.GetFileNameWithoutExtension(BomFileName) + "-sticker.docx" + " in " + FileFolder;

                    // Calculate the total number of stickers per page
                    int stickersPerPage = rows * columns;

                    // Calculate the number of pages needed
                    int totalItems = stickers.Count;
                    int totalPages = (int)Math.Ceiling((double)totalItems / stickersPerPage);

                    // Initialize Word application
                    WordApplication wordApp = new WordApplication();
                    Document wordDoc = wordApp.Documents.Add();
                    // Set the paper size based on the user selection
                    if (paperSize == "A4")
                    {
                        wordDoc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                    }
                    else if (paperSize == "Letter")
                    {
                        wordDoc.PageSetup.PaperSize = WdPaperSize.wdPaperLetter;
                    }
                    // Set margins (in points)
                    float marginTop = wordDoc.PageSetup.TopMargin;
                    float marginBottom = wordDoc.PageSetup.BottomMargin;
                    float marginLeft = wordDoc.PageSetup.LeftMargin;
                    float marginRight = wordDoc.PageSetup.RightMargin;
                    

                    // Calculate available printable width and height
                    float printableWidth = wordDoc.PageSetup.PageWidth - marginLeft - marginRight-1;
                    float printableHeight = wordDoc.PageSetup.PageHeight - marginTop - marginBottom-1;

                    // Calculate cell width and height
                    float cellWidth = printableWidth / columns;
                    float cellHeight = printableHeight / rows;


                    // Populate the document with the array contents
                    int currentIndex = 0;
                    for (int page = 1; page <= totalPages; page++)
                    {
                        if (currentIndex >= totalItems) break;

                        if (page > 1)
                        {
                            //wordDoc.Content.Collapse(WdCollapseDirection.wdCollapseEnd);
                            //wordDoc.Paragraphs.Add();

                            wordDoc.Content.Collapse(WdCollapseDirection.wdCollapseEnd);
                            // Insert a section break to create a new page
                            wordDoc.Sections.Add();
                            wordDoc.Content.Collapse(WdCollapseDirection.wdCollapseEnd);

                        }
                        // Add a table to the Word document
                        //Table table = wordDoc.Tables.Add(wordDoc.Range(wordDoc.Content.End - 1), rows, columns);
                        Microsoft.Office.Interop.Word.Range tableRange = wordDoc.Content;
                        tableRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        Table table = wordDoc.Tables.Add(tableRange, rows, columns);
                        table.Borders.Enable = 1;

                        // Set table and cell dimensions
                        table.Rows.SetHeight(cellHeight, WdRowHeightRule.wdRowHeightExactly);
                        foreach (Row row in table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                cell.Width = cellWidth;
                                // Set text alignment
                                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            }
                        }

                        // Populate the table with the string array contents
                        for ( i = 0; i < stickersPerPage && currentIndex < totalItems; i++)
                        {
                            int row = i / columns + 1;
                            int column = i % columns + 1;

                            table.Cell(row, column).Range.Text = stickers[currentIndex++];

                            ProgressBar1.Value = currentIndex;
                            LabelPercentage.Text = (currentIndex * 100 / ProgressBar1.Maximum) + "%";
                            LabelPercentage.Refresh();
                        }

                        ////Add a page break if not the last page
                        //if (page < totalPages)
                        //{
                        //    //wordDoc.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                        //    Microsoft.Office.Interop.Word.Range rng = wordDoc.Words.Last;
                        //    rng.Collapse(WdCollapseDirection.wdCollapseEnd);
                        //    rng.InsertBreak(WdBreakType.wdPageBreak);
                        //}

                    }



                    // Save the document
                    object filename = Path.Combine(BomFolderFullPath, Path.GetFileNameWithoutExtension(BomFileName) + "-sticker.docx");
                    //object filename = @"C:\path\to\your\document.docx";
                    wordDoc.SaveAs2(ref filename);
                    wordDoc.Close();
                    wordApp.Quit();

                    //MessageBox.Show("Word document for stickers "+ Path.GetFileNameWithoutExtension(BomFileName) + "-sticker.docx" + " was created successfully.");

                    LabelStatus.Text = "Created " + ProgressBar1.Maximum + " stickers in "+ Path.GetFileNameWithoutExtension(BomFileName) + "-sticker.docx";
                    LabelStatus.Refresh();

                    CreateStickerWordDoc.Dispose();
                    CreateStickerWordDoc = null;
                }

            }
            catch
            {

            }

        }

        private void ButtonCreateRadanProjectCSV_Click(object sender, EventArgs e)
        {
            string SheetNameProject;
            string FileExt;
            string ExcelFileName;
            string FileFolder;
            string CNCFileFolder;
            int iRow1;
            int iRowLast;
            int iColumnPartName;
            int iColumnQty = 6;
            int iCoulumnMaterialRadan = 16;

            try
            {

                SheetNameProject = TextSheetName.Text.Trim();
                if (SheetNameProject == "")
                { SheetNameProject = "Project Info"; }

                FileExt = TextCNCFileExtension.Text.Trim();
                if (FileExt != "")
                {
                    if (FileExt.StartsWith("*"))
                    { FileExt = FileExt.Substring(1, FileExt.Length - 1); }
                    else if (!FileExt.StartsWith("."))
                    { FileExt = "." + FileExt; }
                }
                else
                {
                    FileExt = ".sym";
                }

                iRow1 = (int)NumericUpDownRow1.Value;
                iRowLast = (int)NumericUpDownRowLast.Value;
                iColumnPartName = (int)NumericUpDownPartNoColumn.Value;
                iColumnQty = (int)NumericUpDownPartQty.Value;
                iCoulumnMaterialRadan = (int)NumericUpDownRadanMaterial.Value;

                ExcelFileName = TextExcelPathName.Text.Trim();
                if (!System.IO.File.Exists(ExcelFileName))
                {
                    MessageBox.Show(ExcelFileName + " does not exist!", "Warning!!!!!!");
                    ButtonFindBOM_Click(sender, e);
                    return;
                }
                else
                {
                    FileFolder = Path.GetDirectoryName(ExcelFileName);
                    if (FileFolder.Trim() != "")
                    {
                        if (FileFolder.Substring(FileFolder.Length - 1) != "\\")
                        { FileFolder += "\\"; }
                    }

                    //CNCFileFolder = FileFolder + "CNC\\";
                    CNCFileFolder = TextCNCProjectFolder.Text.Trim();
                    if (CNCFileFolder!= "")
                    {
                        if (CNCFileFolder.Substring(CNCFileFolder.Length - 1) != "\\")
                        { CNCFileFolder += "\\"; }
                    }

                    if (!Directory.Exists(CNCFileFolder))
                    {
                        CNCFileFolder = FileFolder + "CNC\\";
                        if (!Directory.Exists(CNCFileFolder))
                        {
                            Directory.CreateDirectory(CNCFileFolder);
                        }
                    }

                    ClassHandleExcels CreateRadanProjectCSV = new ClassHandleExcels(ExcelFileName, true);
                    CreateRadanProjectCSV.CreateRadanProjectCSV(SheetNameProject, FileExt
                                                  , iRow1, iRowLast, iColumnPartName, iColumnQty, iCoulumnMaterialRadan
                                                  , FileFolder, CNCFileFolder
                                                  , ref LabelStatus, ref LabelPercentage, ref ProgressBar1);
                    CreateRadanProjectCSV.Dispose();
                    CreateRadanProjectCSV = null;

                }

            }
            catch
            {

            }
        }

        private void BtnFind_Click(object sender, EventArgs e)
        {
            string folderPath = "";
            if (System.IO.File.Exists(TextExcelPathName.Text))
            { folderPath = Path.GetDirectoryName(TextExcelPathName.Text); }
            else
            { folderPath = TextExcelPathName.Text; }
            if (folderPath.Trim() != "")
            {
                if (folderPath.Substring(folderPath.Length - 1) != "\\")
                { folderPath += "\\"; }
            }
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show(folderPath+" is not a valid folder!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                InitialDirectory = folderPath
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    if (!ListBoxFiles.Items.Contains(file) && file != CombinedCSVFileName)
                    {
                        ListBoxFiles.Items.Add(file);
                    }
                }
            }
        }

        private void BtnCombine_Click(object sender, EventArgs e)
        {
            CombineCsvFiles();
        }

        private void ButtonChangeFlieName_Click(object sender, EventArgs e)
        {
            string folderPath="";
            if (System.IO.File.Exists(TextExcelPathName.Text))
            {folderPath = Path.GetDirectoryName(TextExcelPathName.Text);}
            else
            { folderPath =TextExcelPathName.Text; }

            if (folderPath.Trim() != "")
            {
                if (folderPath.Substring(folderPath.Length - 1) != "\\")
                { folderPath += "\\"; }
            }
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show("Please provide a folder where you want to change the file names!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                // Display a confirmation dialog with Yes and No buttons
                DialogResult result = MessageBox.Show("Is " + folderPath + " the right folder?", "Confirmation",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                // If the user clicks 'No', return from the function
                if (result == DialogResult.No)
                {
                    return;  // Stop the function execution
                }                         
            }
            string textSuffix = TextSuffix.Text.Trim();
            string textPair1 = TextPair1.Text.Trim();
            string textPair2 = TextPair2.Text.Trim();
            string textExt = TextExt.Text.Trim();

            // Check if extension is provided
            if (string.IsNullOrEmpty(textExt))
            {
                MessageBox.Show("Please provide a file extension in TextExt!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Ensure the extension starts with a period (.) when filtering files
            if (!textExt.StartsWith("."))
            {
                textExt = "." + textExt;
            }

            // Get all files in the folder with the specified extension
            string[] files = Directory.GetFiles(folderPath, "*" + textExt);

            string prefix = "";
            string suffix = "";

            // Check if TextSuffix contains A;B format and split it
            if (!string.IsNullOrEmpty(textSuffix))
            {
                string[] suffixParts = textSuffix.Split(';');
                if (suffixParts.Length == 2)
                {
                    prefix = suffixParts[0].Trim(); // Assign the first part as prefix
                    suffix = suffixParts[1].Trim(); // Assign the second part as suffix
                }
                else
                {
                    suffix = textSuffix; // If no `;` is present, use the whole value as suffix
                }
            }

            int i = 1;
            foreach (var file in files)
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(file); // File name without extension
                string originalExtension = Path.GetExtension(file);  // Original extension

                string writerName = ExtractBetween(fileNameWithoutExt, textPair1); // Extract writer name
                string nameString = ExtractBetween(fileNameWithoutExt, textPair2);  // Extract name string

                string newFileName = "";
                if (!string.IsNullOrEmpty(prefix))
                {
                    newFileName += $"{prefix}-";
                }

                // Handle case when no data is extracted
                if (!string.IsNullOrEmpty(writerName) && !string.IsNullOrEmpty(nameString))
                {
                    // Generate the new file name
                    newFileName += $"{writerName} {nameString}";
                }
                else// 
                {
                    if (string.IsNullOrEmpty(writerName))
                    {
                        newFileName += $"{nameString}";
                    }
                    else if (string.IsNullOrEmpty(nameString))
                    {
                        newFileName += $"{writerName}";
                    }
                    newFileName += i.ToString();
                }

                if (!string.IsNullOrEmpty(suffix))
                {
                    newFileName += $"-{suffix}";
                }
                newFileName += originalExtension;  // Retain the original extension

                // Rename the file
                string newFilePath = Path.Combine(folderPath, newFileName);
                File.Move(file, newFilePath);
                i++;
            }

            MessageBox.Show(i-1+" Files renamed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        private string ExtractBetween(string fileName, string pair)
        {
            if (string.IsNullOrEmpty(pair))
            {
                return null;
            }

            string[] delimiters = pair.Split(';');
            if (delimiters.Length == 2)
            {
                string startDelimiter = delimiters[0];
                string endDelimiter = delimiters[1];

                // Handle the case where startDelimiter is empty (means from the beginning)
                if (string.IsNullOrEmpty(startDelimiter))
                {
                    startDelimiter = "^";  // Start of the string
                }
                else
                {
                    startDelimiter = Regex.Escape(startDelimiter);
                }

                // Handle the case where endDelimiter is empty (means to the end of the string)
                if (string.IsNullOrEmpty(endDelimiter))
                {
                    endDelimiter = "$";  // End of the string
                }
                else
                {
                    endDelimiter = Regex.Escape(endDelimiter);
                }

                // Regex to extract the part between start and end delimiters
                string pattern = startDelimiter + "(.*?)" + endDelimiter;

                var match = Regex.Match(fileName, pattern);
                if (match.Success)
                {
                        return match.Groups[1].Value.Trim();  // Extract the matched content between delimiters
                }
            }
            return null;
        }
    }
}
