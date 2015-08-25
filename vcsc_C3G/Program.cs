using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Diagnostics;
using System.Reflection;
using System.Timers;
using System.Resources;
using Excel = Microsoft.Office.Interop.Excel;


 // test 
namespace vcsc_C3G
{
static class Buffer
    {
        static List<string> _Logbuffer; // Static List instance
        static Buffer() { _Logbuffer = new List<string>(); }
        public static void Record(string value) { _Logbuffer.Add(value); }
        public static void Delete(string value) { _Logbuffer.Remove(value); }
        public static Int32 Count() { return _Logbuffer.Count(); }
        public static List<string> getbuffer() { return _Logbuffer; }
        public static void Display() { foreach (var value in _Logbuffer) { Console.WriteLine(value); } }
        public static bool Contains(string file){if (_Logbuffer.Contains(file)) { return true; }else { return false; } }
    }

static class Debug
{
    public static void Init()
    {
        Trace.Listeners.Add(new TextWriterTraceListener("C3g_Debug.log"));
        Trace.AutoFlush = true;
        Trace.Indent();
        Trace.Unindent();
        Trace.Flush();
    }
    public static void Restart()
    {
        Console.WriteLine("System will restart in 10 seconds");
        System.Threading.Thread.Sleep(10000);
        var fileName = Assembly.GetExecutingAssembly().Location;
        System.Diagnostics.Process.Start(fileName);
        Environment.Exit(0);
    }
    public static void Message(string ls_part, string ls_message)
    {
        Trace.WriteLine("DT: " + System.DateTime.Now + " P: " + ls_part + " M: " + ls_message);
        Console.WriteLine("DT: " + System.DateTime.Now + " P: " + ls_part + " M: " + ls_message);
        using(EventLog eventlog = new EventLog("Application"))
        {
            eventlog.Source = "Application";
            eventlog.WriteEntry(ls_message, EventLogEntryType.Information, 101, 1);
        }
    }
}

public class ConsoleSpiner
{
    int counter;
    public ConsoleSpiner()
    {
        counter = 0;
    }
    public void Turn()
    {
        counter++;
        switch (counter % 4)
        {
            case 0: Console.Write("/"); break;
            case 1: Console.Write("-"); break;
            case 2: Console.Write("\\"); break;
            case 3: Console.Write("-"); break;
        }
        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
    }
} 

class Program
    {
         //Main
        static void Main(string[] args)
        {
            Console.Title = "VOLVO Comau C3G Reads Const in Ltool files | Build by SDEBEUL version: 0.01";
            Console.BufferHeight = 100;
            Debug.Init();
            Debug.Message("INFO", "System restarted");
            Console.WriteLine(" call of Ltool FILES scan Take smoke break can take up to 10 min");
            ToolfileScan();
            Debug.Message("INFO", "Found " + Buffer.Count());
            //*****************************************************************************************************************************************
            //big buffer table 
            DataTable BigBuffer = MakeConstBufferTable();
            BigBuffer.AcceptChanges();
            
            ConsoleSpiner spin = new ConsoleSpiner();
            while (true)
            {
            System.Threading.Thread.Sleep(500);
            try
            {

                if (Buffer.Count() == 0) 
                {
                    Console.Write("\r System ready (buffer empty)   Rows in table: {0}                        ", BigBuffer.Rows.Count); 
                    spin.Turn();
                    BigBuffer.ExportToExcel(AppDomain.CurrentDomain.BaseDirectory + "ToolFile.xlsx");
                    Console.WriteLine("Done....");
                    Console.ReadKey();

                }
                else
                {
                    List<string> localbuffer = Buffer.getbuffer();
                    Int32 Cfilecount = 1;
                    foreach (string file in localbuffer.ToList())
                    {
                        try
                        {
                            Console.Write("\r System ready | Filebuffer status: {0:D3} | Localbuffer:  {1:D3} ", Buffer.Count(), Cfilecount);
                            spin.Turn();
                            Cfilecount++;
                            if (IsFileReady(file) && Buffer.Contains(file)) 
                            {
                                TranslateC3G(file);
                                string currentpdl = Regex.Replace(file, ".cod", ".pdl", RegexOptions.IgnoreCase);
                                if (IsFileReady(currentpdl)) 
                                {

                                    foreach (DataRow dr in ReadC3GGunConstants(currentpdl).Rows) 
                                            {
                                                BigBuffer.Rows.Add(dr.ItemArray);
                                            }
                                    Buffer.Delete(file);
                                    File.Delete(currentpdl);
                                    
                                } 
                                else 
                                { Buffer.Delete(file); } 
                            }
                            else if (!File.Exists(file)) { Buffer.Delete(file); Debug.Message("FileNotExistWhileInBuffer", file.Substring(Math.Max(0, file.Length - 40))); }
                        }
                        catch (Exception ex) { Debug.Message("Buffersweep", file.Substring(Math.Max(0, file.Length - 40)) + " msg: " + ex.Message); }
                    }
                }
            }
            catch (Exception ex) {Debug.Message("GeneralCatch", " msg: " + ex.Message); } 
          }
        }

        //scan for var files

        //scan for var files
        private static void ToolfileScan()
        {
            List<string> VARSearchpaths = new List<String>() {@"\\gnl9011101\6308-APP-NASROBOTBCK0001\Robot_ga\ROBLAB\",
                @"\\gnl9011101\6308-APP-NASROBOTBCK0001\Robot_ga\SIBO\", 
                @"\\gnl9011101\6308-APP-NASROBOTBCK0001\Robot_ga\FLOOR\",
                @"\\gnl9011101\6308-APP-NASROBOTBCK0001\Robot_ga\P1X_SIBO\",
                @"\\gnl9011101\6308-APP-NASROBOTBCK0001\Robot_ga\P1X_FLOOR\"};
            List<string> VARExeptedfiles = new List<string>() { "LTOOL_1","LTOOL_2","LTOOL_19","LTOOL_20" };
            List<string> VARExeptedFolders = new List<string>() { @"\transfert\" };
            List<string> VARResultList = ReqSearchDir(VARSearchpaths, "*.cod", VARExeptedfiles, VARExeptedFolders);
            foreach (string file in VARResultList) { Buffer.Record(file); }
        }
 
        //*****************************************************************************************************************************************
        //File reading
        //*****************************************************************************************************************************************  
        //Read the logfile
        private static DataTable ReadC3GGunConstants(string fullFilePath)
        {
            try
            {
                string[] lines = System.IO.File.ReadAllLines(fullFilePath);
                // buffer table
                DataTable Buffer = MakeConstBufferTable();
                DataRow row = Buffer.NewRow();
                Buffer.AcceptChanges();
                foreach (string line in lines)
                {
                    if (line.Contains("CONST") && line.Contains("="))
                    {
                        String Const = ExtractString(line,"CONST","=");
                        String Comment = "NA";
                        String Value = "";
                        if (line.Contains("--"))
                        {
                            Value = ExtractString(line, "=", "--");
                            Comment = line.Substring((line.IndexOf("--") + 2)).Trim();
                        }
                        else
                        {
                            Value = line.Substring((line.IndexOf("=") + 1)).Trim();
                        }

                        //Console.WriteLine("Const: {0} |Value:  {1} |Comment: {2}", Const, Value,Comment);
                        row = Buffer.NewRow();
                        row["controller_name"] = GetRobotName(fullFilePath);
                        row["Tool_file"] = Path.GetFileName(fullFilePath);
                        row["Const"] = Const;
                        row["Value"] = Value;
                        row["Comment"] = Comment;
                        Buffer.Rows.Add(row);
                    
                    }
                    if (line.Contains("_sbcu_check(") && line.Contains("--"))
                    {
                        row = Buffer.NewRow();
                        row["controller_name"] = GetRobotName(fullFilePath);
                        row["Tool_file"] = Path.GetFileName(fullFilePath);
                        row["Const"] = "SBCU comment out detected";
                        row["Value"] = "OUT OF USE";
                        row["Comment"] = line;
                        Buffer.Rows.Add(row);
                    }

                }
                return Buffer;
            }
            catch (Exception e)
            {
                Debug.Message("constReading", fullFilePath.Substring(Math.Max(0, fullFilePath.Length - 40)) + " Msg: " + e.Message);
                DataTable Buffer = MakeConstBufferTable();
                return Buffer;
            }
        }    
    //function to get data(string) between 2 other string regex matches
       static string ExtractString(string s, string start, string end)
        {
            int startIndex = s.IndexOf(start) + start.Length;
            int endIndex = s.IndexOf(end, startIndex);
            return s.Substring(startIndex, endIndex - startIndex).Trim();
        }
     //Make datatable templates
        private static DataTable MakeConstBufferTable()
        {
            DataTable Buffer = new DataTable("Constant");

            DataColumn controller_id = new DataColumn();
            controller_id.DataType = System.Type.GetType("System.String");
            controller_id.ColumnName = "controller_name";
            Buffer.Columns.Add(controller_id);

            DataColumn Tool_file = new DataColumn();
            Tool_file.DataType = System.Type.GetType("System.String");
            Tool_file.ColumnName = "Tool_file";
            Buffer.Columns.Add(Tool_file);

            DataColumn Module = new DataColumn();
            Module.DataType = System.Type.GetType("System.String");
            Module.ColumnName = "Const";
            Buffer.Columns.Add(Module);

            DataColumn Version = new DataColumn();
            Version.DataType = System.Type.GetType("System.String");
            Version.ColumnName = "Value";
            Buffer.Columns.Add(Version);

            DataColumn Comment = new DataColumn();
            Comment.DataType = System.Type.GetType("System.String");
            Comment.ColumnName = "Comment";
            Buffer.Columns.Add(Comment);

            return Buffer;

        }
        //function to convert a comau date string to datetime (handles all possible data types for comau)
        static DateTime ConvertComauDate(String ad_date)
        {
            DateTime parsedDate;
            // first posible comau date type
            // ad_date = "26-09-14 17:40:45"; datepatern that would be provided 
            string pattern1 = "dd-MM-yy HH:mm:ss";
            //
            // second possible comau date type
            // ad_date = "1-OCT-14 17:40:45"; //datepatern that would be provided 
            string pattern2 = "d-MMM-yy HH:mm:ss";
             CultureInfo ci = CultureInfo.CreateSpecificCulture("en-GB");
             DateTimeFormatInfo dtfi = ci.DateTimeFormat;
             dtfi.AbbreviatedMonthNames = new string[] { "JAN", "FEB", "MAR", 
                                                   "APR", "MAY", "JUN", 
                                                   "JUL", "AUG", "SEP", 
                                                   "OCT", "NOV", "DEC", "" };
             dtfi.AbbreviatedMonthGenitiveNames = dtfi.AbbreviatedMonthNames;
            // thrd posibile coma data type
            //ad_date = "Sun Aug 31 01:27:00 2014";
             string pattern3 = "ddd MMM dd HH:mm:ss yyyy";
            if (DateTime.TryParseExact(ad_date.Trim(), pattern1, null, DateTimeStyles.AdjustToUniversal, out parsedDate))
            {
                //Console.WriteLine("Converted '{0}' to {1}.", ad_date, parsedDate);
                return parsedDate;
            }
            else if (DateTime.TryParseExact(ad_date, pattern2, new CultureInfo("en-GB"), DateTimeStyles.AllowWhiteSpaces, out parsedDate))
            {
                //Console.WriteLine("Converted '{0}' to {1}.", ad_date, parsedDate);
                return parsedDate;
            }
            else if (DateTime.TryParseExact(ad_date, pattern3, new CultureInfo("en-GB"), DateTimeStyles.AllowWhiteSpaces, out parsedDate))
            {
                //Console.WriteLine("Converted '{0}' to {1}.", ad_date, parsedDate);
                return parsedDate;
            }
            else if (ad_date.Contains("00-")) //y2k bug 
            {
              Debug.Message("dt convert", "y2kbug trow:  " + ad_date);
              return Convert.ToDateTime("2000-01-01 00:00:00.00");
            }

            else
            {
                Debug.Message("Dt converter" ,"Unable to convert: " + ad_date + "to a date and time.");
                return parsedDate;
            }
        }
        //function to extract Robotname from a string 
        static String GetRobotName(String As_inString)
        {
            String Result = "";
            var regexR = new Regex(@"(\d\d\d\d\d)R(\d\d)");
            var matchR = regexR.Match(As_inString);
             if (matchR.Success) { Result = matchR.Groups[1].Value + 'R' + matchR.Groups[2].Value; }
             var regexP = new Regex(@"(\d\d\d\d\d)P(\d\d)");
             var matchP = regexP.Match(As_inString);
             if (matchP.Success) {Result = matchP.Groups[1].Value + 'P' + matchP.Groups[2].Value; }
            return Result;
        }
        //function to extract tool id from a string 
        static Int16 GetToolId(String As_instring)
        {
            Int16 toolid = Convert.ToInt16(As_instring.Substring((As_instring.IndexOf("Tool_", StringComparison.OrdinalIgnoreCase) + 5), 2));
            return toolid;
        }
        //translate comau file 
        static void TranslateC3G(String as_FullFilepath)
        {
            //extract the C3G decomplir from the resource into the executionpath
            byte[] exeBytes = Properties.Resources.pdl2_v561;
            string exeToRun = new Uri(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + @"\pdl2_v561.exe").LocalPath;
            if (!File.Exists(exeToRun)) {  using (FileStream exeFile = new FileStream(exeToRun, FileMode.CreateNew)) { exeFile.Write(exeBytes, 0, exeBytes.Length); } }
            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = as_FullFilepath.Replace(Path.GetFileName(as_FullFilepath), "").Trim();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = exeToRun;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.RedirectStandardOutput = true;
            startInfo.Arguments = @"/B " + as_FullFilepath;
            try { using (Process exeProcess = Process.Start(startInfo)){ exeProcess.WaitForExit(); }}
            catch { Debug.Message("TranslationErr", " For: " + GetRobotName(as_FullFilepath)); }
        }
        //*****************************************************************************************************************************************
        //FILE HANDELING
        //*****************************************************************************************************************************************  
        //search for files
        static List<string> ReqSearchDir(List<string> als_filepaths, string as_mask, List<String> als_exeptedFiles, List<String> als_exeptedFolders)
        {
            List<string> List = new List<string>();
            try
            {
               foreach (string filepath in als_filepaths)
               {
                   Console.WriteLine("\r Searching: {1} Found: {0:D3}", List.Count, filepath.Substring(Math.Max(0, filepath.Length - 40)));
                   var allFiles = Directory.GetFiles(filepath, as_mask, SearchOption.AllDirectories);
                foreach (string f in allFiles)
                {
                    foreach (string exeptedFolder in als_exeptedFolders)
                    {
                        if (f.Contains(exeptedFolder))
                        {
                            foreach (string exeptedfile in als_exeptedFiles) { if (f.Contains(exeptedfile)) { List.Add(f);}}
                        }
                    }
                }
            }
           }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return List;
        }
        //removes folders / subfolders if empty
        static private void RemoveEmptyFolders(string folderPath)
        {
            var allFolders = Directory.GetDirectories(folderPath);
            if (allFolders.Length > 0)
            {
                foreach (string folder in allFolders)
                {
                    RemoveEmptyFolders(folder);
                }
            }

            if (Directory.GetDirectories(folderPath).Length == 0 && Directory.GetFiles(folderPath).Length == 0)
            {
                Directory.Delete(folderPath);
            }
        }
        //check if file is accesable
        public static bool IsFileReady(String sFilename)
        {
            // If the file can be opened for exclusive access it means that the file
            // is no longer locked by another process.
            try
            {
                using (FileStream inputStream = File.Open(sFilename, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    if (inputStream.Length > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        //Deletes the file. (with savegard to only delete files in log location
        public static void SafeDelete(String fullPath)
        {
            if (fullPath.IndexOf(@"\\gnl9011101\6308-APP-NASROBOTBCK0001\logs\Comau\3\", 0, StringComparison.CurrentCultureIgnoreCase) != -1) 
            { File.Delete(fullPath); }
            else { Debug.Message("IllegalFiledelete", fullPath.Substring(Math.Max(0, fullPath.Length - 40))); }
        }
        //*****************************************************************************************************************************************
        //UTIL
        //*****************************************************************************************************************************************      
        static void CallexcelMacro(string Excelfile, string macro)
        {

            try
            {
                //~~> Define your Excel Objects
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook;
                //~~> Start Excel and open the workbook.
                xlWorkBook = xlApp.Workbooks.Open(Excelfile);
                //~~> Run the macros by supplying the necessary arguments
                xlApp.Run(macro);
                //~~> Clean-up: Close the workbook
                xlWorkBook.Close(false);
                //~~> Quit the Excel Application
                xlApp.Quit();
                //~~> Clean Up
                releaseObject(xlApp);
                releaseObject(xlWorkBook);
            }
            catch(Exception ex)
            {
                Debug.Message("Excel caller",ex.Message);
            }
        }
        static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }


    } // end of progam

          // datatablto csv 
public static class DataTableExtensions {
        public static void WriteToCsvFile(this DataTable dataTable, string filePath) {
            StringBuilder fileContent = new StringBuilder();

            foreach (var col in dataTable.Columns) {
                fileContent.Append(col.ToString() + ",");
            }

            fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);



            foreach (DataRow dr in dataTable.Rows) {

                foreach (var column in dr.ItemArray) {
                    fileContent.Append("\"" + column.ToString() + "\",");
                }

                fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);
            }

           // System.IO.File.WriteAllText(filePath, fileContent.ToString());
            System.IO.File.AppendAllText (filePath, fileContent.ToString());
        }
    }

public static class My_DataTable_Extensions
{

    // Export DataTable into an excel file with field names in the header line
    // - Save excel file without ever making it visible if filepath is given
    // - Don't save excel file, just make it visible if no filepath is given
    public static void ExportToExcel(this DataTable Tbl, string ExcelFilePath = null)
    {
        try
        {
            if (Tbl == null || Tbl.Columns.Count == 0)
                throw new Exception("ExportToExcel: Null or empty input table!\n");

            // load excel, and create a new workbook
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            // single worksheet
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            // column headings
            for (int i = 0; i < Tbl.Columns.Count; i++)
            {
                workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
            }

            // rows
            for (int i = 0; i < Tbl.Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (int j = 0; j < Tbl.Columns.Count; j++)
                {
                    workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                }
            }

            // check fielpath
            if (ExcelFilePath != null && ExcelFilePath != "")
            {
                try
                {
                    workSheet.SaveAs(ExcelFilePath);
                    excelApp.Quit();
                   // MessageBox.Show("Excel file saved!");
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                        + ex.Message);
                }
            }
            else    // no filepath is given
            {
                excelApp.Visible = true;
            }
        }
        catch (Exception ex)
        {
            throw new Exception("ExportToExcel: \n" + ex.Message);
        }
    }
}
} // end of namespace
