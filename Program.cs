using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JustinToolsBasket
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormJustinTools());
        }
        public static Boolean TheConfigNameIsDefault(string swConfigName)
        {
            if (swConfigName.Length > 1)
            {
                if ((swConfigName.Substring(0, 2).ToUpper() == "DE") || (swConfigName.Substring(0, 2) == "默认"))
                { return true; }
                else
                { return false; }
            }
            else
            { return false; }

        }

        public static double ConvertStringToDouble(string value)
        {

            NumberFormatInfo numberFormatInfo = new System.Globalization.NumberFormatInfo();
            NumberFormatInfo provider = numberFormatInfo;

            provider.NumberDecimalSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            provider.NumberGroupSeparator = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyGroupSeparator; ;
            provider.NumberGroupSizes = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyGroupSizes;

            NumberFormatInfo numberFormatInfo1 = new System.Globalization.NumberFormatInfo();
            NumberFormatInfo providerE = numberFormatInfo1;
            providerE.NumberDecimalSeparator = ".";
            providerE.NumberGroupSeparator = ",";
            providerE.NumberGroupSizes = new int[] { 3 };

            double rd = 1;

            try
            {
                rd = Convert.ToDouble(value);
            }
            catch
            {
                try
                {
                    rd = Convert.ToDouble(value, provider);
                }
                catch
                {

                    try
                    {
                        rd = Convert.ToDouble(value, providerE);
                    }
                    catch
                    {


                    }
                }

            }

            return rd;

        }

        public static String ConvertFromDoubleToStringInEnglishFormat(double value)
        {

            System.Globalization.CultureInfo _oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //string format = "+#.##0.00;-#.##0.00;(0)";

            //string sValue = value.ToString("#,##0.00;(-#,##0.00);Zero");
            string sValue = value.ToString("F3", CultureInfo.InvariantCulture);
            System.Threading.Thread.CurrentThread.CurrentCulture = _oldCI;

            return sValue;
        }


        public static String ConvertFromDoubleToStringInEnglishFormatRoundNumber(double value)
        {

            System.Globalization.CultureInfo _oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            value = Math.Round(value);

            string sValue = value.ToString();

            //string sValue = value.ToString("#,##0;($#,##0);Zero");

            System.Threading.Thread.CurrentThread.CurrentCulture = _oldCI;

            return sValue;
        }

        public static string ConvertDoubleInMMToControlText(double ControlText, double unitFactor)
        {
            try
            {

                return ConvertFromDoubleToStringInEnglishFormat(ControlText / unitFactor);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Control Text " + ControlText + " can not been converted to double"
                                     + ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                     + System.Environment.NewLine + "For help, please email the information of this error"
                                     + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
                return "0";
            }
        }

        public static double ConvertControlTextToDoubleInMM(string ControlText, double unitFactor)
        {
            try
            {

                if (ControlText != null)//
                {
                    return ConvertStringToDouble(ControlText) * unitFactor;
                }
                else
                {
                    MessageBox.Show("Control Text " + ControlText + " can not been converted to double"
                                    + System.Environment.NewLine + "For help, please email the information of this error"
                                    + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Control Text " + ControlText + " can not been converted to double"
                                     + ex.Message + System.Environment.NewLine + "Stack Trace:" + ex.StackTrace
                                     + System.Environment.NewLine + "For help, please email the information of this error"
                                     + System.Environment.NewLine + "to : Justin.luo@halton.com ", "Error!!!!!!");
                return 1;
            }
        }

        /// <summary>
        /// get the value in string from excel cells
        /// the checkbox control only accept the bool value, so i create this function
        /// </summary>
        /// <param name="Boolvalue"></param>
        /// <returns></returns>
        public static bool ConvertStringToBool(string Boolvalue)
        {
            try
            {
                //return (Convert.ToBoolean(Boolvalue));
                if (Boolvalue != null)//there is no this checkbox value in the excel
                {
                    if (Boolvalue.ToUpper() == "TRUE")
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (FormatException)
            { return false; }

        }

        public static string SearchOneFileIfExistInFolders(string[] CNCFolders, string FileName, string CNCProgramFileAffix, string CncProjectFolder, bool bClean)
        {
            string result = "new";
            foreach (string CNCFolder in CNCFolders)
            {
                if (result == "new")
                {
                    if (CNCFolder != "")
                    {
                        string CNCFolderSearch;
                        if (CNCFolder.Substring(CNCFolder.Length - 1) != "\\")
                        {
                            CNCFolderSearch = CNCFolder + "\\";
                        }
                        else
                        { CNCFolderSearch = CNCFolder; }
                        //FileExistInAFolderAndAllItsChildrenOrNot(FileName, CNCFolderSearch, CNCProgramFileAffix, ref result, CncProjectFolder, bClean);
                        DirSearchrecursivelyForFile(FileName, CNCFolderSearch, CNCProgramFileAffix, ref result, CncProjectFolder, bClean);
                    }
                }
                else
                { return result; }
            }
            return result;
        }

        public static void DirSearchrecursivelyForFile(string FileName, string CNCFolder, string CNCProgramFileAffix, ref string result, string CncProjectFolder, bool bClean)
        {
            try
            {
                FileExistInAFolderOrNot(FileName, CNCFolder, CNCProgramFileAffix, ref result, CncProjectFolder, bClean);

                if (result == "new")
                {
                    foreach (string d in Directory.GetDirectories(CNCFolder))
                    {
                        DirSearchrecursivelyForFile(FileName, d, CNCProgramFileAffix, ref result, CncProjectFolder, bClean);
                    }
                }
                else
                { return; }
            }
            catch (System.Exception excpt)
            {
                //Console.WriteLine(excpt.Message);
                MessageBox.Show(excpt.Message + System.Environment.NewLine + "Stack Trace:" + excpt.StackTrace
                                      + System.Environment.NewLine + "For help, please email the information of this error"
                                      + System.Environment.NewLine + "to : Justin.luo@halton.com ");
            }
        }

        public static void FileExistInAFolderOrNot(string FileName, string CNCFolder, string CNCProgramFileAffix, ref string result, string CncProjectFolder, bool bClean)
        {
            if (result == "new")
            {
                try
                {
                    if (CncProjectFolder.Substring(CncProjectFolder.Length - 1) != "\\" && CncProjectFolder != "")
                    { CncProjectFolder += "\\"; }

                    if (CNCFolder.Substring(CNCFolder.Length - 1) != "\\" && CNCFolder != "")
                    { CNCFolder += "\\"; }

                    string fileFullPathName = CNCFolder + FileName + CNCProgramFileAffix;
                    string fileProjectFullPathName = CncProjectFolder + FileName + CNCProgramFileAffix;

                    if (File.Exists(fileFullPathName))
                    {
                        result = fileFullPathName;

                        if (bClean)
                        {
                            if (File.Exists(fileProjectFullPathName))
                            { File.Delete(fileProjectFullPathName); }
                        }
                        else
                        {
                            if (!File.Exists(fileProjectFullPathName))
                            { File.Copy(fileFullPathName, fileProjectFullPathName, true); }
                        }

                        return;

                    }

                    //MessageBox.Show("Find the cp file: "result );


                }
                catch (System.Exception excpt)
                {
                    Console.WriteLine(excpt.Message);
                    //MessageBox.Show(excpt.Message + System.Environment.NewLine + "Stack Trace:" + excpt.StackTrace
                    //                      + System.Environment.NewLine + "For help, please email the information of this error"
                    //                      + System.Environment.NewLine + "to : Justin.luo@halton.com ");
                }

            }
        }

        public static void FileExistInAFolderAndAllItsChildrenOrNot(string FileName, string CNCFolder, string CNCProgramFileAffix, ref string result, string CncProjectFolder, bool bClean)
        {
            result = "new";
            try
            {
                if (CncProjectFolder.Substring(CncProjectFolder.Length - 1) != "\\" && CncProjectFolder != "")
                { CncProjectFolder += "\\"; }

                if (CNCFolder.Substring(CNCFolder.Length - 1) != "\\" && CNCFolder != "")
                { CNCFolder += "\\"; }

                string fileFullPathName = CNCFolder + FileName + CNCProgramFileAffix;
                string fileProjectFullPathName = CncProjectFolder + FileName + CNCProgramFileAffix;

                string[] files = Directory.GetFiles(CNCFolder, FileName + CNCProgramFileAffix, SearchOption.AllDirectories);

                if (files.Length > 0)
                {
                    result = files[0];

                    if (bClean)
                    {
                        if (File.Exists(fileProjectFullPathName))
                        { File.Delete(fileProjectFullPathName); }
                        //File.Copy(files[0], fileProjectFullPathName, true);
                    }
                    else
                    {
                        if (!File.Exists(fileProjectFullPathName))
                        { File.Copy(files[0], fileProjectFullPathName, true); }
                    }
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        public static void DeletFilesWithTextInFolders(string[] CNCFolders, string FileName, string CNCProgramFileAffix, ref TextBox textCncProgramTobeDel)
        {
            bool bGotit = false;
            string CNCFolderSearch;
            foreach (string CNCFolder in CNCFolders)
            {
                if (CNCFolder != "")
                {

                    if (CNCFolder.Substring(CNCFolder.Length - 1) != "\\")
                    {
                        CNCFolderSearch = CNCFolder + "\\";

                    }
                    else
                    { CNCFolderSearch = CNCFolder; }

                    IList<string> listDeleteFiles = new List<string>();
                    DirSearchrecursivelyForDeletFilesListWithText(FileName, CNCFolderSearch, CNCProgramFileAffix, ref bGotit, ref textCncProgramTobeDel, ref listDeleteFiles);

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

                    File.WriteAllText(TextFilePathName, message);
                    //File.OpenRead (TextFilePathName );
                    System.Diagnostics.Process.Start(TextFilePathName);

                    DialogResult result = System.Windows.Forms.DialogResult.No;
                    result = MessageBox.Show(messageContent, caption, buttons);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {

                        listDeleteFiles.Clear();
                        // delete all the files, and add revision for them.
                        string[] readText = File.ReadAllLines(TextFilePathName);
                        foreach (string s in readText)
                        {

                            if (s != "")
                            {
                                if (File.Exists(s))
                                { listDeleteFiles.Add(s); }
                            }
                        }

                        foreach (string DeleteFileName in listDeleteFiles)
                        {
                            CreateRevsionVersionOfaFileThenDeleteIt(DeleteFileName);
                        }

                    }

                    //DirSearchrecursivelyForDeletFilesWithText(FileName, CNCFolderSearch, CNCProgramFileAffix, ref bGotit, ref textCncProgramTobeDel);

                }

            }
            textCncProgramTobeDel.Text = FileName;
            textCncProgramTobeDel.Refresh();
        }
        public static void MoveFilesWithTextInFoldersToTargetFolder(string[] CNCFolders, string FileName, string CNCProgramFileAffix, ref TextBox textCncProgramTobeDel, bool bDel, string sTargetFolder)
        {
            bool bGotit = false;
            string CNCFolderSearch;
            foreach (string CNCFolder in CNCFolders)
            {
                if (CNCFolder != "")
                {


                    if (CNCFolder.Substring(CNCFolder.Length - 1) != "\\")
                    {
                        CNCFolderSearch = CNCFolder + "\\";

                    }
                    else
                    { CNCFolderSearch = CNCFolder; }

                    DirSearchrecursivelyForMoveFilesWithText(FileName, CNCFolderSearch, CNCProgramFileAffix, ref bGotit, ref textCncProgramTobeDel, bDel, sTargetFolder);
                }

            }
            textCncProgramTobeDel.Text = FileName;
            textCncProgramTobeDel.Refresh();
        }
        /// <summary>
        /// if file exist, it will check whether there is old revision version exist.
        /// </summary>
        /// <param name="fileName"></param>
        public static void CreateRevsionVersionOfaFileThenDeleteIt(string fileName)
        {
            bool isItARevisionFile = false;
            string f = Path.GetFileNameWithoutExtension(fileName);
            //check if it is a revsion file. if it is, just ignore this file ,don't do anything
            if (f.Length > 6)
            {
                if (f.Substring(f.Length - 6, 3).ToUpper() == "REV")
                {
                    isItARevisionFile = true;
                }
            }
            if (!isItARevisionFile)
            {

                File.Copy(fileName, GetFlieNewRevisonName(fileName), true);
                File.Delete(fileName);

            }

        }

        /// <summary>
        /// check the topest revision number for this file in the same folder
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetFlieNewRevisonName(string fileName)
        {
            int i = 1;
            while (File.Exists(GetFileRevisionNameWithRevisionNumber(fileName, i)))
            {
                i++;
            }
            return GetFileRevisionNameWithRevisionNumber(fileName, i);

        }

        /// <summary>
        /// affix a revsionnumber to a file, use the format of -Rev001
        /// keep 3 digit number for revision number
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        public static string GetFileRevisionNameWithRevisionNumber(string fileName, int i)
        {
            string sExt = Path.GetExtension(fileName);
            string sFileFolder = Path.GetDirectoryName(fileName) + @"\";
            string sfilename = Path.GetFileNameWithoutExtension(fileName);
            string sRevNum;

            if (i < 10)
            {
                sRevNum = "00" + i.ToString();
            }
            else if (i < 100)
            {
                sRevNum = "0" + i.ToString();
            }
            else
            {
                sRevNum = i.ToString();
            }

            return sFileFolder + sfilename + "-Rev" + sRevNum + sExt;
        }


        public static void DirSearchrecursivelyForDeletFilesWithText(string FileName, string CNCFolder, string CNCProgramFileAffix, ref bool bGotit, ref TextBox textCncProgramTobeDel)
        {
            try
            {
                if (!bGotit)
                {
                    textCncProgramTobeDel.Text = "It is searching for " + FileName + " in the folder: " + CNCFolder + ".";
                    textCncProgramTobeDel.Refresh();
                    foreach (string f in Directory.GetFiles(CNCFolder, "*" + CNCProgramFileAffix))
                    {

                        if ((f.ToUpper()).Contains(FileName.ToUpper()))
                        {
                            CreateRevsionVersionOfaFileThenDeleteIt(f);
                            bGotit = false;
                            //bGotit = true;
                        }
                    }

                    if (!bGotit)
                    {
                        foreach (string d in Directory.GetDirectories(CNCFolder))
                        {

                            textCncProgramTobeDel.Text = "It is searching for " + FileName + " in the folder: " + d + ".";
                            textCncProgramTobeDel.Refresh();
                            //foreach (string f in Directory.GetFiles(d, "*" + CNCProgramFileAffix))
                            //{
                            //    if (f.Contains(FileName))
                            //    {
                            //        File.Delete(f);
                            //        bGotit = true;
                            //    }
                            //}

                            DirSearchrecursivelyForDeletFilesWithText(FileName, d, CNCProgramFileAffix, ref bGotit, ref textCncProgramTobeDel);
                        }
                    }
                }
                else
                { return; }
            }
            catch (System.Exception excpt)
            {
                //Console.WriteLine(excpt.Message);
                MessageBox.Show(excpt.Message + System.Environment.NewLine + "Stack Trace:" + excpt.StackTrace
                                      + System.Environment.NewLine + "For help, please email the information of this error"
                                      + System.Environment.NewLine + "to : Justin.luo@halton.com ");
            }
        }
        public static void CopyFileToTargetFolderAndDeleteOrNot(string fileName, bool bDel, string CncProjectFolder)
        {
            string sExt = Path.GetExtension(fileName);
            string sfilename = Path.GetFileNameWithoutExtension(fileName);

            if (CncProjectFolder.Substring(CncProjectFolder.Length - 1) != "\\" && CncProjectFolder != "")
            { CncProjectFolder += "\\"; }

            string sToTargetFileName = CncProjectFolder + sfilename + sExt;
            if (bDel)//delete the one in database
            {
                File.Copy(fileName, sToTargetFileName, true);
                File.Delete(fileName);
            }
            else
            {
                File.Copy(fileName, sToTargetFileName, true);
            }
        }

        /// <summary>
        /// this function is to find all the files that need to be deleted, and put them into a list variable listDeleteFiles
        /// </summary>
        /// <param name="FileName"></param>
        /// <param name="CNCFolder"></param>
        /// <param name="CNCProgramFileAffix"></param>
        /// <param name="bGotit"></param>
        /// <param name="textCncProgramTobeDel"></param>
        /// <param name="listDeleteFiles"></param>
        public static void DirSearchrecursivelyForDeletFilesListWithText(string FileName, string CNCFolder, string CNCProgramFileAffix, ref bool bGotit, ref TextBox textCncProgramTobeDel, ref IList<string> listDeleteFiles)
        {
            try
            {
                bool isItARevisionFile;
                string fileName;
                if (!bGotit)
                {
                    textCncProgramTobeDel.Text = "Searching for " + FileName + " in the folder: " + CNCFolder + ".";
                    textCncProgramTobeDel.Refresh();
                    foreach (string f in Directory.GetFiles(CNCFolder, "*" + CNCProgramFileAffix))
                    {

                        if ((f.ToUpper()).Contains(FileName.ToUpper()))
                        {
                            isItARevisionFile = false;
                            fileName = Path.GetFileNameWithoutExtension(f);
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
                                listDeleteFiles.Add(f);
                            }

                            bGotit = false;
                            //bGotit = true;
                        }
                    }

                    if (!bGotit)
                    {
                        foreach (string d in Directory.GetDirectories(CNCFolder))
                        {

                            textCncProgramTobeDel.Text = "It is searching for " + FileName + " in the folder: " + d + ".";
                            textCncProgramTobeDel.Refresh();


                            DirSearchrecursivelyForDeletFilesListWithText(FileName, d, CNCProgramFileAffix, ref bGotit, ref textCncProgramTobeDel, ref listDeleteFiles);
                        }
                    }
                }
                else
                { return; }
            }

            catch (System.Exception excpt)
            {
                //Console.WriteLine(excpt.Message);
                MessageBox.Show(excpt.Message + System.Environment.NewLine + "Stack Trace:" + excpt.StackTrace
                                      + System.Environment.NewLine + "For help, please email the information of this error"
                                      + System.Environment.NewLine + "to : Justin.luo@halton.com ");
            }
        }


        public static void DirSearchrecursivelyForMoveFilesWithText(string FileName, string CNCFolder, string CNCProgramFileAffix, ref bool bGotit, ref TextBox textCncProgramTobeDel, bool bDel, string sTargetFolder)
        {
            try
            {
                if (!bGotit)
                {
                    textCncProgramTobeDel.Text = "It is searching for " + FileName + " in the folder: " + CNCFolder + ".";
                    textCncProgramTobeDel.Refresh();
                    foreach (string f in Directory.GetFiles(CNCFolder, "*" + CNCProgramFileAffix))
                    {

                        if ((f.ToUpper()).Contains(FileName.ToUpper()))
                        {
                            CopyFileToTargetFolderAndDeleteOrNot(f, bDel, sTargetFolder);
                            bGotit = false;
                            //bGotit = true;;
                        }
                    }

                    if (!bGotit)
                    {
                        foreach (string d in Directory.GetDirectories(CNCFolder))
                        {

                            textCncProgramTobeDel.Text = "It is searching for " + FileName + " in the folder: " + d + ".";
                            textCncProgramTobeDel.Refresh();

                            DirSearchrecursivelyForMoveFilesWithText(FileName, d, CNCProgramFileAffix, ref bGotit, ref textCncProgramTobeDel, bDel, sTargetFolder);
                        }
                    }
                }
                else
                { return; }
            }
            catch (System.Exception excpt)
            {
                //Console.WriteLine(excpt.Message);
                MessageBox.Show(excpt.Message + System.Environment.NewLine + "Stack Trace:" + excpt.StackTrace
                                      + System.Environment.NewLine + "For help, please email the information of this error"
                                      + System.Environment.NewLine + "to : Justin.luo@halton.com ");
            }


        }

        public static void AddPrintLabelWithFileNameAndInfoTextToAFile(string filePathName, string sInfor)
        {
            //string sExt = Path.GetExtension(filePathName);
            string sfilename = Path.GetFileNameWithoutExtension(filePathName);
            System.Text.StringBuilder newFile = new System.Text.StringBuilder();

            string temp;

            string[] file = File.ReadAllLines(filePathName);

            foreach (string line in file)

            {
                if (line.Contains("PRINT_LABEL -1 "))//if other CNC program use other type of label, add a else if

                {

                    temp = "PRINT_LABEL -1 32128 (" + sfilename + ")(" + sInfor + ");";

                    newFile.Append(temp + "\r\n");

                }
                else if (line.Contains("INKJET -1 "))//if other CNC program use other type of label, add a else if

                {

                    temp = "INKJET -1 32128 (" + sfilename + ")(" + sInfor + ")()();";

                    newFile.Append(temp + "\r\n");

                }
                else
                {

                    newFile.Append(line + "\r\n");
                }

            }

            File.WriteAllText(filePathName, newFile.ToString());


        }




    }
}
