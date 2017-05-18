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
            Console.Title = "VOLVO Comau C3G vcsc Build by SDEBEUL version: 17W20D04";
            Console.BufferHeight = 100;
            Debug.Init();
            Debug.Message("INFO", "System restarted");
            //*****************************************************************************************************************************************
            //build file sytem watch  
            try { 
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\logs\Comau\3\";
            watcher.InternalBufferSize = (watcher.InternalBufferSize * 2); //2 times default buffer size 
            watcher.Error += new ErrorEventHandler(OnError);
            watcher.Filter = "*.LOG";
            watcher.IncludeSubdirectories = true;
            watcher.Error += new ErrorEventHandler(OnError);
            watcher.Created += new FileSystemEventHandler(OnCreate);
            watcher.EnableRaisingEvents = true;
            }
            catch (Exception ex) { Debug.Message("Wachter", ex.Message); Debug.Restart(); }
            //*****************************************************************************************************************************************
            Task.Run(() => VarfileScan());
            //*****************************************************************************************************************************************
            Task.Run(() => C3GLogFilescan());
            //*****************************************************************************************************************************************
            Timer TriggerTimer = new System.Timers.Timer(24 * 60 * 60 * 1000); //run every day 
            TriggerTimer.Start();
            TriggerTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            //*****************************************************************************************************************************************
            
            ConsoleSpiner spin = new ConsoleSpiner();
            while (true)
            {
            System.Threading.Thread.Sleep(500);
            try
            {
                if (Buffer.Count() == 0) { Console.Write("\r System ready (buffer empty)                               "); spin.Turn(); }
                else
                {
                    List<string> localbuffer = Buffer.getbuffer();
                    Int32 Cfilecount = 1;
                    foreach (string file in localbuffer.ToList())
                    {
                        try
                        {
                            Console.Write("\r System ready | Filebuffer Processing: | {1:D3} / {0:D3} |  ", Buffer.Count(), Cfilecount);
                            spin.Turn();
                            Cfilecount++;
                            if (IsFileReady(file) && Buffer.Contains(file)) { HandelVarfile(file); }
                            if (IsFileReady(file) && Buffer.Contains(file)) { HandelLogfile(file); }
                            else if (!File.Exists(file)) { Buffer.Delete(file); Debug.Message("FileNotExistWhileInBuffer", file.Substring(Math.Max(0, file.Length - 40))); }
                        }
                        catch (Exception ex) { Debug.Message("Buffersweep", file.Substring(Math.Max(0, file.Length - 60)) + " msg: " + ex.Message); }
                    }
                }
            }
            catch (Exception ex) {Debug.Message("GeneralCatch", " msg: " + ex.Message); } 
          }
        }
        //scan for Log files
        private static void C3GLogFilescan()
        {
            Debug.Message("INFO", "Logfilescan"); 
            List<string> LOGSearchpaths = new List<String>() { @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\logs\Comau\3\" };
            List<string> LOGExeptedfiles = new List<string>() { "TOOL_01.LOG", "TOOL_02.LOG", "TOOL_03.LOG", "TOOL_04.LOG", "ERROR.LOG" };
            List<string> LOGExeptedFolders = new List<string>() { @"\Comau\3\" };
            List<string> LOGResultList = ReqSearchDir(LOGSearchpaths, "*.LOG", LOGExeptedfiles, LOGExeptedFolders);
            foreach (string file in LOGResultList) { Buffer.Record(file);}
            Debug.Message("INFO", "Logfilescan DONE");
        }
        //scan for var files
        private static void VarfileScan()
        {
            Debug.Message("INFO", "Varfilescan"); 
            List<string> VARSearchpaths = new List<String>() {
                @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\Robot_ga\ROBLAB\"};/*,
                @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\Robot_ga\SIBO\", 
                @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\Robot_ga\FLOOR\",
                @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\Robot_ga\P1X_SIBO\",
                @"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\Robot_ga\P1X_FLOOR\"};*/
            List<string> VARExeptedfiles = new List<string>() { "LY413", "LY283", "LY55X", "LA440","LA441","LA442", "LTOOL_", "TT_TOOL1.VAR", "TUVFRAME.VAR", "LArc", "LGripp", "LStatGun", "LGun", "Lstud" };
            List<string> VARExeptedFolders = new List<string>() { @"\transfert\" };
            List<string> VARResultList = ReqSearchDir(VARSearchpaths, "*.VAR", VARExeptedfiles, VARExeptedFolders);
            foreach (string file in VARResultList) { Buffer.Record(file); }
            Debug.Message("INFO", "Varfilescan DONE");
        }

        // Event handeler for priodic scan ecent 
        private static void OnTimedEvent(object source, ElapsedEventArgs e) 
        { 
            Task.Run(() => VarfileScan()); 
        }    
        // Event handeler for robot puts file on server
        private static void OnCreate(object source, FileSystemEventArgs e){Buffer.Record(e.FullPath); }
        // Event handeler for error event in wacther (auto restart)
        private static void OnError(object source, ErrorEventArgs e) { Debug.Message("FileWachter", e.GetException().Message); Debug.Restart(); }
        //*****************************************************************************************************************************************
        //File reading
        //*****************************************************************************************************************************************  
        //handle log file
        public static void HandelLogfile(string fullFilepath)
        {
                DataTable buffertable = new DataTable();
                switch (IsC3GLog(fullFilepath))
                {
                    case "Errorlog":
                        /*
                        buffertable = ReadC3GErrlog(fullFilepath);
                        buffertable = CheckDataConsistensyC3G(buffertable);
                        BulkCopyToGadata("ROBOTGA",buffertable, "rt_alarm");
                        */
                        SafeDelete(fullFilepath);
                        Buffer.Delete(fullFilepath);
                        RemoveEmptyFolders(@"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\logs\Comau\3\" + GetRobotName(fullFilepath));
                        break;
                    case "Toollog":
                        buffertable = ReadC3GToollog(fullFilepath);
                        BulkCopyToGadata("C3G", buffertable, "rt_toollog");
                        SafeDelete(fullFilepath);
                        Buffer.Delete(fullFilepath);
                        RemoveEmptyFolders(@"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\logs\Comau\3\" + GetRobotName(fullFilepath));
                        break;
                    default:
                        Debug.Message("Unknow filetype", fullFilepath.Substring(Math.Max(0, fullFilepath.Length - 40)));
                        File.Delete(fullFilepath);
                        Buffer.Delete(fullFilepath);
                        break;
                }
                buffertable.Dispose();

        }
        //handle Var file
        public static void HandelVarfile(string fullFilePath)
        {
            if (fullFilePath.IndexOf("var", 0, StringComparison.CurrentCultureIgnoreCase) != -1)
            {
                if (!fullFilePath.EndsWith("var",StringComparison.CurrentCultureIgnoreCase)) 
                {
                    Debug.Message("VarfileInvalid: ", GetRobotName(fullFilePath) + " File: " + fullFilePath.Substring(Math.Max(0, fullFilePath.Length - 40)));
                    Buffer.Delete(fullFilePath); return;  
                }

                Buffer.Delete(fullFilePath);
                Int32 C3GRobotID = 0;
                Int32 C4GRobotID = 0;
                //check if the robot is C3G and translate
                if (GetC3GRobotID(GetRobotName(fullFilePath)) != 0)
                { TranslateC3G(fullFilePath); C3GRobotID = GetC3GRobotID(GetRobotName(fullFilePath)); } 
                //check if the robot is C4G and translate
                else if (GetC4GRobotID(GetRobotName(fullFilePath)) != 0)
                { TranslateC4G(fullFilePath); C4GRobotID = GetC4GRobotID(GetRobotName(fullFilePath)); }
                else
                { Debug.Message("Robot not in DB: ", GetRobotName(fullFilePath) + "| ID returns 0"); return; }

                if (File.Exists(Regex.Replace(fullFilePath, ".var", ".lsv", RegexOptions.IgnoreCase)))
                {
                    DataTable buffer = new DataTable();
                    if (C4GRobotID != 0)
                    {
                       buffer = ReadPosVarFile(Regex.Replace(fullFilePath, ".var", ".lsv", RegexOptions.IgnoreCase), C4GRobotID);
                       BulkCopyToGadata("C4G", buffer, "L_robotpositions"); 
                    }
                    if (C3GRobotID != 0) 
                    {
                        buffer = ReadPosVarFile(Regex.Replace(fullFilePath, ".var", ".lsv", RegexOptions.IgnoreCase), C3GRobotID);
                        BulkCopyToGadata("C3G", buffer, "L_robotpositions"); 
                    }
                    File.Delete(Regex.Replace(fullFilePath, ".var", ".lsv", RegexOptions.IgnoreCase));
                }
                else 
                {
                    Debug.Message("TranslateERR: ", GetRobotName(fullFilePath) + " File: " + fullFilePath.Substring(Math.Max(0, fullFilePath.Length - 40)));
                }

            }
        }
        //check if tile is errorlog
        public static string IsC3GLog(string fullFilePath)
        {
        Stream stream = File.Open(fullFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using (var reader = new StreamReader(stream))
        {
            var hasComau = false;
            var hasCorrectType = false;
            var hasDmeas = false;
            var hasDsetup = false;

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (!hasComau)
                {
                    if (line.StartsWith("*  C O M A U  :  Robotics              *"))  { hasComau = true; }
                }
                if (!hasCorrectType)
                {
                    if (line.StartsWith("* ERROR FORMAT RELEASE     : 1.0       *")) { hasCorrectType = true; }
                }
                if (!hasDmeas)
                {
                    if (line.Contains("dmeas=")) { hasDmeas = true; }
                }
                if (!hasDsetup)
                {
                    if (line.Contains("dsetup=")) { hasDsetup = true; }
                }
                //saves me from reading whole file
                if ((hasCorrectType && hasComau) | (hasDmeas && hasDsetup)){break;}
            }
            if (hasCorrectType && hasCorrectType) { return "Errorlog"; }
            else if (hasDmeas && hasDsetup) {return "Toollog"; }
            else { return "Unknown"; }        
          }
    }
        //Read the logfilemm
        private static DataTable ReadC3GErrlog(string fullFilePath)
        {
            try
            {
                string[] lines = System.IO.File.ReadAllLines(fullFilePath);
                // file reading
                int index = 0;
                string sPattern = "<...>";
                string dateString = "";
                string FullErrorCodeString = "";
                string LogtextString = "";
                Int32 Logcode = 0;
                Int32 LogSeverity = 0;
                //get robot id 
                Int32 RobotId = GetC3GRobotID(GetRobotName(fullFilePath));
                // buffer table
                DataTable Buffer = MakeErrorlogBufferTable();
                DataRow row = Buffer.NewRow();
                Buffer.AcceptChanges();
                foreach (string line in lines)
                {
                    //finds begin of datetime line
                    if (System.Text.RegularExpressions.Regex.IsMatch(line, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        //extract datetime format from current line
                        dateString = Regex.Replace(line, sPattern, "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        // read errorlog from next line AND Split logcode and logtext  // (28694-10 ):( Safety gate or Emergency STOP)   
                        FullErrorCodeString = lines[index + 1].Split(':')[0];
                        //LogtextString = lines[index + 1].Split(':')[1];
                        LogtextString = lines[index + 1].Substring(FullErrorCodeString.Length + 1);
                        // Split logcode // (28694)-(10 ) AND Convert to int 
                        Logcode = Convert.ToInt32(FullErrorCodeString.Split('-')[0]);
                        LogSeverity = Convert.ToInt32(FullErrorCodeString.Split('-')[1]);
                        //Console.WriteLine("Date: '{0}'  Err: '{1}' serv: '{2}' Text: '{3}", ConvertComauDate(dateString.Trim()).ToString(), Logcode.ToString(), LogSeverity.ToString(), LogtextString.Trim());
                        row = Buffer.NewRow();
                        row["controller_id"] = RobotId;
                        row["error_timestamp"] = ConvertComauDate(dateString.Trim());
                        row["error_number"] = Logcode;
                        row["error_severity"] = LogSeverity;
                        row["error_text"] = LogtextString.Trim();
                        row["error_text_id"] = DBNull.Value;
                        Buffer.Rows.Add(row);
                    }
                    index++;
                }
                return Buffer;
            }
            catch (Exception e)
            {
                Debug.Message("LogReading", fullFilePath.Substring(Math.Max(0, fullFilePath.Length - 40)) + " Msg: " + e.Message);
                DataTable Buffer = MakeErrorlogBufferTable();
                return Buffer;
            }
        }
        private static DataTable ReadC3GToollog(string fullFilePath)
        {
            string[] lines = System.IO.File.ReadAllLines(fullFilePath);
            // file reading
            int index = 0;
            string sPattern = ".-...-.....:..:..";
            //get robot id 
            Int32 RobotId = GetC3GRobotID(GetRobotName(fullFilePath));
            // buffer table
            DataTable Buffer = MakeToollogBufferTable();
            DataRow row = Buffer.NewRow();
            Buffer.AcceptChanges();
            //
            foreach (string line in lines)
            {
                //check if the SBCU is in simulation mode 
                if (line.Contains("MinUpdate:"))
                {
                    //MinUpdate:  1.000, MaxUpdate:  10.000, MinWearing:  0.000
                    float iMinUpdate = float.Parse(line.Split(':')[1].Split(',')[0].Trim(), CultureInfo.InvariantCulture);
                    float iMaxUpdate = float.Parse(line.Split(':')[2].Split(',')[0].Trim(), CultureInfo.InvariantCulture);
                    if (iMaxUpdate > 10 || iMinUpdate > 1)
                    {
                        SendrErrorC3G(RobotId, 99907, 2, string.Format("SBCU warning: {0}",line));
                    }
                }

                //finds begin of datetime line and next line has tcp
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase) && lines[index + 1].Contains("T <"))
                {
                    //extract datetime format from current line
                   string dateString = line.Substring(0, line.IndexOf("dmeas"));
                    //extract dmeas and dsetup from current line
                   //if (line.ToString().Contains("Attr")) //selection for type of logfile (new version with attrbute)
                   //{
                   float Dmeas = float.Parse(line.ToString().Split('=')[1].Replace("dsetup", "").Trim(), CultureInfo.InvariantCulture);
                   float Dsetup = float.Parse(line.ToString().Split('=')[2].Replace("Attr", "").Trim(), CultureInfo.InvariantCulture);
                   Boolean Longcheck = false;
                   if (line.Contains("Attr") && line.ToString().Split('=')[3].Contains('L')) { Longcheck = true; }
                   Boolean Update = false;
                   if (line.Contains("Attr") && line.ToString().Split('=')[3].Contains('U')) { Update = true; }
                   //get toolvalues from next logline 
                       // T < 338.492, 262.238, 1060.765, -142.690, 155.900, 34.730,>
                       string TcpString = lines[index + 1].Replace("T <", "").Replace(",>", "").Replace("\0", "").Trim();
                       // 338.492, 262.238, 1060.765, -142.690, 155.900, 34.730
                       float x = float.Parse(TcpString.Split(',')[0], CultureInfo.InstalledUICulture);
                       float y = float.Parse(TcpString.Split(',')[1], CultureInfo.InstalledUICulture);
                       float z = float.Parse(TcpString.Split(',')[2], CultureInfo.InstalledUICulture);
                       float a = float.Parse(TcpString.Split(',')[3], CultureInfo.InstalledUICulture);
                       float e = float.Parse(TcpString.Split(',')[4], CultureInfo.InstalledUICulture);
                       float r = float.Parse(TcpString.Split(',')[5], CultureInfo.InstalledUICulture);
                       //Console.WriteLine("x: '{0}'  y: '{1}' z: '{2}' a: '{3}'  e: '{4}' r: '{5}'", x, y, z, a, e, r);
                  //add to buffer 
                   row = Buffer.NewRow();
                   if (line.Contains("Attr"))
                   {
                     row["Longcheck"] = Longcheck;
                     row["TcpUpdate"] = Update;
                   }
                   else
                   {
                     row["Longcheck"] = DBNull.Value;
                     row["TcpUpdate"] = DBNull.Value;
                   }
                   row["controller_id"] = RobotId;
                   row["tool_timestamp"] = ConvertComauDate(dateString);
                   row["tool_id"] = GetToolId(fullFilePath);
                   row["Dmeas"] = Dmeas;
                   row["Dsetup"] = Dsetup;
                   row["ToolX"] = x;
                   row["ToolY"] = y;
                   row["ToolZ"] = z;
                   row["ToolA"] = a;
                   row["ToolE"] = e;
                   row["ToolR"] = r;
                   Buffer.Rows.Add(row); 
                }
                index++;
            }
            return Buffer;
        }
        private static DataTable ReadPosVarFile(string fullFilePath, Int32 Robotid)
        {
            string[] lines = System.IO.File.ReadAllLines(fullFilePath);
            //to extract date from file 
            string datestring = lines[1].Substring((lines[1].IndexOf(".VAR", StringComparison.OrdinalIgnoreCase) + 4 + 2));
            // file reading
            int index = 0;
            int TFnum = 1;
            Int32 numtools = 0;
            string sPatternPOS = "POS  Priv";
            string sPatternXTND1 = "XTND Arm: 1 Ax: 1";
            string sPatternXTND2 = "XTND Arm: 1 Ax: 2";
            string sPatternTool = "vp_tools";  
            string sPatternFrame = "vp_frames";
            string sPatternC4GFrame = "REC tuye_frame_table";
            string sPatternC4GTool = "REC ttye_tool_table";
            // buffer table
            DataTable Buffer = MakePosBufferTable();
            DataRow row = Buffer.NewRow();
            Buffer.AcceptChanges();
            //
            foreach (string line in lines)
            {
              //  Console.WriteLine(line);
                Boolean bPOSmode = false;
                Boolean bXTND1mode = false;
                Boolean bXTND2mode = false;
                Boolean Toolmode = false;
                Boolean Framemode = false;
                Boolean C4gFramemode = false;
                Boolean C4gToolmode = false;
                string posname = "";
                string cnfg = "";
                float x = 0.0f;
                float y = 0.0f;
                float z = 0.0f;
                float a = 0.0f;
                float e = 0.0f;
                float r = 0.0f;
                float ax7 = 0.0f;
                float ax8 = 0.0f;
                Int16 ratio = 1; //ratio is the relation beween the posname line and the line where we read the position. (implemented for c4g tool / frame)
                //finds position lines
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternPOS, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) {bPOSmode = true;}  
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternXTND1, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) {bXTND1mode = true;}
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternXTND2, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) { bXTND2mode = true; }
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternTool, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) { Toolmode = true;
                 numtools = Int32.Parse(line.Substring((line.IndexOf("APOS[") + 5), 2).Trim());
                }
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternFrame, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) { Framemode = true;
                 numtools = Int32.Parse(line.Substring((line.IndexOf("APOS[") + 5), 2).Trim());
                }
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternC4GFrame, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) { C4gFramemode = true; ratio = 2; }
                if (System.Text.RegularExpressions.Regex.IsMatch(line, sPatternC4GTool, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) { C4gToolmode = true; ratio = 2; }

                // extract pos name (pos name should be before the MATCH on same line)
                if (bPOSmode) { posname = line.Substring(0, line.IndexOf(sPatternPOS)).Trim(); };
                if (bXTND1mode) { posname = line.Substring(0, line.IndexOf(sPatternXTND1)).Trim(); };
                if (bXTND2mode) { posname = line.Substring(0, line.IndexOf(sPatternXTND2)).Trim(); };
                if (C4gFramemode) { posname = line.Substring(0, line.IndexOf(sPatternC4GFrame)).Trim(); };
                if (C4gToolmode) { posname = line.Substring(0, line.IndexOf(sPatternC4GTool)).Trim(); };
            NextTF:
                if (Toolmode && TFnum < numtools)
                {
                    posname = "Tool_" + TFnum; 
                    TFnum++; 
                    index++;
                    
                }
                else {Toolmode = false;}

            if (Framemode && TFnum < numtools)
            {
                posname = "Frame_" + TFnum;
                TFnum++;
                index++;

            }
            else { Framemode = false; }



            if (bPOSmode | bXTND1mode | bXTND2mode | Toolmode | Framemode | C4gFramemode | C4gToolmode)
                {
                    //get position from next line  Line ex:  X:4606.30 Y:-366.59 Z:1373.78 A: -15.32 E:  37.21 R:-154.37
                    if (!lines[index + ratio].Contains("*******")) //handels uninit positions
                    {
                       // string currentline1 = line;
                       // string currentline = lines[index + 1];
                       // Console.WriteLine(currentline);
                        x = float.Parse(lines[index + ratio].Split(':')[1].TrimEnd(new char[] { 'Y' }).Trim(), CultureInfo.InstalledUICulture);
                        y = float.Parse(lines[index + ratio].Split(':')[2].TrimEnd(new char[] { 'Z' }).Trim(), CultureInfo.InstalledUICulture);
                        z = float.Parse(lines[index + ratio].Split(':')[3].TrimEnd(new char[] { 'A' }).Trim(), CultureInfo.InstalledUICulture);
                        a = float.Parse(lines[index + ratio].Split(':')[4].TrimEnd(new char[] { 'E' }).Trim(), CultureInfo.InstalledUICulture);
                        e = float.Parse(lines[index + ratio].Split(':')[5].TrimEnd(new char[] { 'R' }).Trim(), CultureInfo.InstalledUICulture);
                        if (bPOSmode) { r = float.Parse(lines[index + ratio].Split(':')[6].Trim(), CultureInfo.InstalledUICulture); }
                        if (bXTND1mode)
                        {
                            r = float.Parse(lines[index + ratio].Split(':')[6].TrimEnd(new char[] { '1' }).Trim(), CultureInfo.InstalledUICulture);
                            ax7 = float.Parse(lines[index + ratio].Split(':')[7].Trim(), CultureInfo.InstalledUICulture);
                        }
                        if (bXTND2mode)
                        {
                            r = float.Parse(lines[index + ratio].Split(':')[6].TrimEnd(new char[] { '1' }).Trim(), CultureInfo.InstalledUICulture);
                            ax7 = float.Parse(lines[index + ratio].Split(':')[7].TrimEnd(new char[] { '2' }).Trim(), CultureInfo.InstalledUICulture);
                            ax8 = float.Parse(lines[index + ratio].Split(':')[8].Trim(), CultureInfo.InstalledUICulture);
                        }
                        //get cnfg flags from next line Line ex: CNFG: ''
                        cnfg = lines[index + ratio + 1].Replace("CNFG:", "").Replace("'", "").Trim();
                    }
                    //
                    //Console.WriteLine("robot: {9} File: '{0}'  Pos: '{1}' x: '{2}' y: '{3}' z: '{4}' a: '{5}' e: '{6}' r: '{7}' cnfg: '{8}'",
                    //    Path.GetFileNameWithoutExtension(fullFilePath), posname, x, y, z, a, e, r, cnfg,GetRobotName(fullFilePath));
                    //add to buffer 
                    row = Buffer.NewRow();
                    row["controller_id"] = Robotid;
                    row["_timestamp"] = DBNull.Value;
                    row["file_timestamp"] = ConvertComauDate(datestring);
                    row["Owner"] = Path.GetFileNameWithoutExtension(fullFilePath);
                    row["Pos"] = posname;
                    row["X"] = x;
                    row["Y"] = y;
                    row["Z"] = z;
                    row["A"] = a;
                    row["E"] = e;
                    row["R"] = r;
                    if (bXTND1mode | bXTND2mode) { row["ax7"] = ax7; } else { row["ax7"] = DBNull.Value; }
                    if (bXTND2mode) { row["ax8"] = ax8; } else { row["ax8"] = DBNull.Value; }
                    row["Cnfg"] =cnfg;
                    Buffer.Rows.Add(row);

                    if (Toolmode | Framemode) { index = index + 2; goto NextTF; }
                }

                index++;
            }
            return Buffer;
        }
        //Make datatable templates
        private static DataTable MakeErrorlogBufferTable()
        {
            DataTable Buffer = new DataTable("Buffer");

            DataColumn ID = new DataColumn();
            ID.DataType = System.Type.GetType("System.Int32");
            ID.ColumnName = "ID";
            ID.AutoIncrement = true;
            Buffer.Columns.Add(ID);

            DataColumn controller_id = new DataColumn();
            controller_id.DataType = System.Type.GetType("System.Int32");
            controller_id.ColumnName = "controller_id";
            Buffer.Columns.Add(controller_id);

            DataColumn error_timestamp = new DataColumn();
            error_timestamp.DataType = System.Type.GetType("System.DateTime");
            error_timestamp.ColumnName = "error_timestamp";
            Buffer.Columns.Add(error_timestamp);

            DataColumn error_number = new DataColumn();
            error_number.DataType = System.Type.GetType("System.Int32");
            error_number.ColumnName = "error_number";
            Buffer.Columns.Add(error_number);

            DataColumn error_severity = new DataColumn();
            error_severity.DataType = System.Type.GetType("System.Int32");
            error_severity.ColumnName = "error_severity";
            Buffer.Columns.Add(error_severity);

            DataColumn error_text_id = new DataColumn();
            error_text_id.DataType = System.Type.GetType("System.Int32");
            error_text_id.ColumnName = "error_text_id";
            Buffer.Columns.Add(error_text_id);
            
            DataColumn error_text = new DataColumn();
            error_text.DataType = System.Type.GetType("System.String");
            error_text.ColumnName = "error_text";
            Buffer.Columns.Add(error_text);
            DataColumn[] keys = new DataColumn[1];
            keys[0] = ID;
            Buffer.PrimaryKey = keys; 
            return Buffer;

//SQL target table
/*
 USE [GADATA]
GO
***** Object:  Table [RobotGA].[rt_alarm]    Script Date: 2/10/2014 6:19:35 *****
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [RobotGA].[rt_alarm](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[controller_id] [int] NULL,
	[error_timestamp] [datetime] NULL,
	[error_number] [int] NULL,
	[error_severity] [int] NULL,
	[error_text] [varchar](256) NULL,
	[RobotName] [varchar](20) NULL,
 CONSTRAINT [PK_rt_alarm] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


            use GADATA

CREATE UNIQUE NONCLUSTERED INDEX [IndexTableUniqueRows] ON gadata.robotga.rt_alarm
(
       [controller_id]
      ,[error_timestamp]
      ,[error_number]
  ASC

)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
 
 * */
        }
        private static DataTable MakeToollogBufferTable()
        {
            DataTable Buffer = new DataTable("Buffer");

            DataColumn ID = new DataColumn();
            ID.DataType = System.Type.GetType("System.Int32");
            ID.ColumnName = "ID";
            ID.AutoIncrement = true;
            Buffer.Columns.Add(ID);

            DataColumn error_timestamp = new DataColumn();
            error_timestamp.DataType = System.Type.GetType("System.DateTime");
            error_timestamp.ColumnName = "tool_timestamp";
            Buffer.Columns.Add(error_timestamp);

            DataColumn tool_id = new DataColumn();
            tool_id.DataType = System.Type.GetType("System.Int32");
            tool_id.ColumnName = "tool_id";
            Buffer.Columns.Add(tool_id);

            DataColumn Dmeas = new DataColumn();
            Dmeas.DataType = System.Type.GetType("System.Decimal");
            Dmeas.ColumnName = "Dmeas";
            Buffer.Columns.Add(Dmeas);

            DataColumn Dsetup = new DataColumn();
            Dsetup.DataType = System.Type.GetType("System.Decimal");
            Dsetup.ColumnName = "Dsetup";
            Buffer.Columns.Add(Dsetup);

            DataColumn ToolX = new DataColumn();
            ToolX.DataType = System.Type.GetType("System.Decimal");
            ToolX.ColumnName = "ToolX";
            Buffer.Columns.Add(ToolX);

            DataColumn ToolY = new DataColumn();
            ToolY.DataType = System.Type.GetType("System.Decimal");
            ToolY.ColumnName = "ToolY";
            Buffer.Columns.Add(ToolY);

            DataColumn ToolZ = new DataColumn();
            ToolZ.DataType = System.Type.GetType("System.Decimal");
            ToolZ.ColumnName = "ToolZ";
            Buffer.Columns.Add(ToolZ);

            DataColumn ToolA = new DataColumn();
            ToolA.DataType = System.Type.GetType("System.Decimal");
            ToolA.ColumnName = "ToolA";
            Buffer.Columns.Add(ToolA);

            DataColumn ToolE = new DataColumn();
            ToolE.DataType = System.Type.GetType("System.Decimal");
            ToolE.ColumnName = "ToolE";
            Buffer.Columns.Add(ToolE);

            DataColumn ToolR = new DataColumn();
            ToolR.DataType = System.Type.GetType("System.Decimal");
            ToolR.ColumnName = "ToolR";
            Buffer.Columns.Add(ToolR);

            DataColumn controller_id = new DataColumn();
            controller_id.DataType = System.Type.GetType("System.Int32");
            controller_id.ColumnName = "controller_id";
            Buffer.Columns.Add(controller_id);

            DataColumn Longcheck = new DataColumn();
            Longcheck.DataType = System.Type.GetType("System.Boolean");
            Longcheck.ColumnName = "Longcheck";
            Buffer.Columns.Add(Longcheck);

            DataColumn TcpUpdate = new DataColumn();
            TcpUpdate.DataType = System.Type.GetType("System.Boolean");
            TcpUpdate.ColumnName = "TcpUpdate";
            Buffer.Columns.Add(TcpUpdate);

            DataColumn[] keys = new DataColumn[1];
            keys[0] = ID;
            Buffer.PrimaryKey = keys;
 
            return Buffer;

//SQL target table script 
            /*
            USE [GADATA]
           GO
           ****** Object:  Table [RobotGA].[rt_toollog]    Script Date: 2/10/2014 6:17:40 *****
           SET ANSI_NULLS ON
           GO

           SET QUOTED_IDENTIFIER ON
           GO

           SET ANSI_PADDING ON
           GO

           CREATE TABLE [RobotGA].[rt_toollog](
               [ID] [int] IDENTITY(1,1) NOT NULL,
               [tool_timestamp] [datetime] NULL,
               [controller_id] [tinyint] NULL,
               [tool_id] [tinyint] NULL,
               [Dmeas] [real] NULL,
               [Dsetup] [real] NULL,
               [ToolX] [real] NULL,
               [Tooly] [real] NULL,
               [ToolZ] [real] NULL,
               [ToolA] [real] NULL,
               [ToolE] [real] NULL,
               [ToolR] [real] NULL,
               [Robotname] [varchar](20) NULL,
            CONSTRAINT [PK_rt_toollog] PRIMARY KEY CLUSTERED 
           (
               [ID] ASC
           )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
           ) ON [PRIMARY]

           GO

           SET ANSI_PADDING OFF
           GO

             CREATE UNIQUE NONCLUSTERED INDEX [IndexTableUniqueRows] ON gadata.robotga.rt_toollog
           (
                  [tool_timestamp]
                 ,[tool_id]
                 ,[Dmeas]
                 ,[Dsetup]
                 ,[ToolX]
                 ,[Tooly]
                 ,[ToolZ]
                 ,[ToolA]
                 ,[ToolE]
                 ,[ToolR]
                 ,[controller_id]
             ASC

           )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = ON, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
 
             */

        }
        private static DataTable MakePosBufferTable()
        {
            DataTable Buffer = new DataTable("Buffer");

            DataColumn ID = new DataColumn();
            ID.DataType = System.Type.GetType("System.Int32");
            ID.ColumnName = "ID";
            ID.AutoIncrement = true;
            Buffer.Columns.Add(ID);

            DataColumn _timestamp = new DataColumn();
            _timestamp.DataType = System.Type.GetType("System.DateTime");
            _timestamp.ColumnName = "_timestamp";
            Buffer.Columns.Add(_timestamp);

            DataColumn file_timestamp = new DataColumn();
            file_timestamp.DataType = System.Type.GetType("System.DateTime");
            file_timestamp.ColumnName = "file_timestamp";
            Buffer.Columns.Add(file_timestamp);

            DataColumn controller_id = new DataColumn();
            controller_id.DataType = System.Type.GetType("System.Int32");
            controller_id.ColumnName = "controller_id";
            Buffer.Columns.Add(controller_id);

            DataColumn Owner = new DataColumn();
            Owner.DataType = System.Type.GetType("System.String");
            Owner.ColumnName = "Owner";
            Buffer.Columns.Add(Owner);

            DataColumn Pos = new DataColumn();
            Pos.DataType = System.Type.GetType("System.String");
            Pos.ColumnName = "Pos";
            Buffer.Columns.Add(Pos);

            DataColumn X = new DataColumn();
            X.DataType = System.Type.GetType("System.Decimal");
            X.ColumnName = "X";
            Buffer.Columns.Add(X);

            DataColumn Y = new DataColumn();
            Y.DataType = System.Type.GetType("System.Decimal");
            Y.ColumnName = "Y";
            Buffer.Columns.Add(Y);

            DataColumn Z = new DataColumn();
            Z.DataType = System.Type.GetType("System.Decimal");
            Z.ColumnName = "Z";
            Buffer.Columns.Add(Z);

            DataColumn A = new DataColumn();
            A.DataType = System.Type.GetType("System.Decimal");
            A.ColumnName = "A";
            Buffer.Columns.Add(A);

            DataColumn E = new DataColumn();
            E.DataType = System.Type.GetType("System.Decimal");
            E.ColumnName = "E";
            Buffer.Columns.Add(E);

            DataColumn R = new DataColumn();
            R.DataType = System.Type.GetType("System.Decimal");
            R.ColumnName = "R";
            Buffer.Columns.Add(R);

            DataColumn ax7 = new DataColumn();
            ax7.DataType = System.Type.GetType("System.Decimal");
            ax7.ColumnName = "ax7";
            Buffer.Columns.Add(ax7);

            DataColumn ax8 = new DataColumn();
            ax8.DataType = System.Type.GetType("System.Decimal");
            ax8.ColumnName = "ax8";
            Buffer.Columns.Add(ax8);

            DataColumn Cnfg = new DataColumn();
            Cnfg.DataType = System.Type.GetType("System.String");
            Cnfg.ColumnName = "Cnfg";
            Buffer.Columns.Add(Cnfg);

            DataColumn[] keys = new DataColumn[1];
            keys[0] = ID;
            Buffer.PrimaryKey = keys;

            return Buffer;


//SQL script to make tabl
/*
USE [GADATA]
GO
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [RobotGA].[L_robotpositions](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[_timestamp] [datetime] NULL,
	[file_timestamp] [datetime] NULL,
	[controller_id] [tinyint] NULL,
	[Owner] [varchar](50) NULL,
	[Pos] [varchar](50) NULL,
	[X] [real] NULL,
	[Y] [real] NULL,
	[Z] [real] NULL,
	[a] [real] NULL,
	[e] [real] NULL,
	[r] [real] NULL,
	[ax7] [real] NULL,
	[ax8] [real] NULL,
	[Cnfg] [varchar](20) NULL,
 CONSTRAINT [PK_L_robotpositions] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO



 */

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
            // thrd posibile comau data type
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
            catch { Debug.Message("c3gTranslationErr", "robotid: " + GetC3GRobotID(GetRobotName(as_FullFilepath)) + " For: " + GetRobotName(as_FullFilepath)); }
        }
        static void TranslateC4G(String as_FullFilepath)
        {
         try {
            //extract the C4G decomplir from the resource into the executionpath
            byte[] exeBytes = Properties.Resources.c4gtr;
            string exeToRun = new Uri(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + @"\c4gtr.exe").LocalPath;
            if (!File.Exists(exeToRun)) { using (FileStream exeFile = new FileStream(exeToRun, FileMode.CreateNew)) { exeFile.Write(exeBytes, 0, exeBytes.Length); } }
            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = as_FullFilepath.Replace(Path.GetFileName(as_FullFilepath), "").Trim();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = exeToRun;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.RedirectStandardOutput = false;
            startInfo.Arguments = @"/B /V " + Path.GetFileName(as_FullFilepath);
            using (Process exeProcess = Process.Start(startInfo)) {exeProcess.WaitForExit(); }
            }
            catch (Exception ex) { Debug.Message("c4gTranslationErr", "For: "  + GetRobotName(as_FullFilepath) +"  M: " + ex.Message);
            }
        }
        //*****************************************************************************************************************************************
        //SQL
        //*****************************************************************************************************************************************  
        //Send Error to database. (c3g)
        static void SendrErrorC3G(int iRobotId, int iErrorNum, int iErrorSevr, string sErrorText)
        {
            // sql if ts is in db it will return the ts you send.. if not it wil return the last error ts    
            string connectionString = "user id=VCSCHelper; password=VCSCHelper; server=SQLA001.gen.volvocars.net; Trusted_Connection=no; database=gadata; connection timeout=5";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand commandInsertError = new SqlCommand(@"
                            INSERT INTO GADATA.C3G.rt_ALARM 
                            (controller_id,error_timestamp,error_number,error_severity,error_text)
                            VALUES(@robotid,getdate(),@ErrorNum,@ErrorSevr,@ErrorText)
                            ", connection);
                commandInsertError.Parameters.Add(new SqlParameter("@robotid", iRobotId));
                commandInsertError.Parameters.Add(new SqlParameter("@ErrorNum", iErrorNum));
                commandInsertError.Parameters.Add(new SqlParameter("@ErrorSevr", iErrorSevr));
                commandInsertError.Parameters.Add(new SqlParameter("@ErrorText", sErrorText));
                Debug.Message("INFO",string.Format("c3gEERORSEND Robotid:{0} Error:{1}-{2} {3} ",iRobotId,iErrorNum,iErrorSevr,sErrorText));
                var result = commandInsertError.ExecuteScalar();
                connection.Close();
                connection.Dispose();
            }
        }
    
    
        //function to check if there is dataloss
        static DataTable CheckDataConsistensyC3G(DataTable AS_intable)
        {
            DataRow[] Result = AS_intable.Select("", "error_timestamp ASC");
            DataRow firstrow = Result[1];
            DateTime OldestError = Convert.ToDateTime(firstrow[2]);
            // sql if ts is in db it will return the ts you send.. if not it wil return the last error ts    
            string connectionString = "user id=VCSCHelper; password=VCSCHelper; server=SQLA001.gen.volvocars.net; Trusted_Connection=no; database=gadata; connection timeout=5";
            DateTime ResultTs;
            using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand commandGetTS = new SqlCommand(
             "select TOP 1 ISNULL(error_timestamp," +
             "(select TOP 1 error_timestamp from GADATA.RobotGA.rt_alarm WHERE (controller_id LIKE @robotID) AND (error_timestamp < getdate()) ORDER BY error_timestamp DESC))" +
             "from GADATA.RobotGA.rt_alarm WHERE (controller_id LIKE @robotID)  AND  (error_timestamp <= @LasterrTS) ORDER BY error_timestamp DESC", connection);
                    commandGetTS.Parameters.Add(new SqlParameter("robotID", firstrow[1]));
                    commandGetTS.Parameters.Add(new SqlParameter("LasterrTS", OldestError));
                    ResultTs = System.Convert.ToDateTime(commandGetTS.ExecuteScalar());
                    connection.Close();
                    connection.Dispose();
                }
            TimeSpan duration = OldestError - ResultTs;
            // OK => no gap
            if (OldestError == ResultTs) { }//Console.WriteLine("NO Datagap"); }
            //NOK  => get last error in db with ts < one in db
            else
            {  //=> make entry in datatable with latest error ts in db and this date
                Console.WriteLine("!!!!!!!!!!!!************************************!!!!!!!!!!!!");
                Console.WriteLine("Datagap Detected WorstCaseLoss: {0}", duration);
                Console.WriteLine("!!!!!!!!!!!!************************************!!!!!!!!!!!!");
                DataRow row = AS_intable.NewRow();
                AS_intable.AcceptChanges();
                row = AS_intable.NewRow();
                row["controller_id"] = firstrow[1];
                row["error_timestamp"] = OldestError;
                row["error_number"] = 99001;
                row["error_severity"] = 4;
                row["error_text"] = "Datagap detected WorstCaseDataLoss: " + duration;
                AS_intable.Rows.Add(row);  
            }
            return AS_intable;
        }
        //function that gets the robot id from sql
        static Int32 GetC3GRobotID(String As_inString)
        {
            string connectionString = "user id=VCSCHelper; password=VCSCHelper; server=SQLA001.gen.volvocars.net; Trusted_Connection=no; database=gadata; connection timeout=5";
            using (SqlConnection connection =
                       new SqlConnection(connectionString))
            {
                connection.Open();
                // Perform an initial count on the destination table.
                SqlCommand commandGetId = new SqlCommand("select top 1 c_controller.id from GADATA.c3g.c_controller where c_controller.controller_name LIKE '%" + As_inString + "%'", connection);
                Int32 Robotid = System.Convert.ToInt16(commandGetId.ExecuteScalar());
                connection.Close();
                //  Console.WriteLine("c3g Got id {0} for robot {1} from sql", Robotid, As_inString);
                //  Console.ReadLine();
                connection.Dispose();
                return Robotid;
            }

        }
        static Int32 GetC4GRobotID(String As_inString)
        {
            string connectionString = "user id=VCSCHelper; password=VCSCHelper; server=SQLA001.gen.volvocars.net; Trusted_Connection=no; database=gadata; connection timeout=5";
            using (SqlConnection connection =
                       new SqlConnection(connectionString))
            {
                connection.Open();
                // Perform an initial count on the destination table.
                SqlCommand commandGetId = new SqlCommand("select top 1 c_controller.id from GADATA.c4g.c_controller where c_controller.controller_name LIKE '%" + As_inString + "%'", connection);
                Int32 Robotid = System.Convert.ToInt16(commandGetId.ExecuteScalar());
                connection.Close();
                //  Console.WriteLine("c4g Got id {0} for robot {1} from sql", Robotid, As_inString);
                //  Console.ReadLine();
                connection.Dispose();
                return Robotid;
            }

        }
        //Bulk Copy to Gadata
        static void BulkCopyToGadata (string as_schema, DataTable adt_table, string as_destination)
        {
            {
                string connectionString = "user id=VCSCHelper; password=VCSCHelper; server=SQLA001.gen.volvocars.net; Trusted_Connection=no; database=gadata; connection timeout=30";
                using (SqlConnection connection =
                           new SqlConnection(connectionString))
                {
                    connection.Open();
                    // Perform an initial count on the destination table.
                    SqlCommand commandRowCount = new SqlCommand("SELECT COUNT(*) FROM ["+as_schema+"].[" + as_destination + "];", connection);
                    long countStart = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
                    // Note that the column positions in the source DataTable  
                    // match the column positions in the destination table so  
                    // there is no need to map columns.  
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.DestinationTableName = "["+as_schema+"].[" + as_destination + "]";
                        try
                        {
                            // Write from the source to the destination.
                            bulkCopy.WriteToServer(adt_table);
                        }
                        catch (Exception ex)
                        {
                            Debug.Message("Bukcopy", ex.Message);
                            Console.WriteLine(ex.HelpLink);
                        }
                    }
                    //see how many rows were added. 
                    long countEnd = System.Convert.ToInt32(
                    commandRowCount.ExecuteScalar());
                    connection.Close();
                    connection.Dispose();
                    //Console.WriteLine("Detected: {0} rows {1} new rows were added to Gadata.",adt_table.Rows.Count ,(countEnd - countStart));
                }
            }
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
            if (fullPath.IndexOf(@"\\gnlsnm0101.gen.volvocars.net\6308-APP-NASROBOTBCK0001\logs\Comau\3\", 0, StringComparison.CurrentCultureIgnoreCase) != -1) 
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
} // end of namespace
