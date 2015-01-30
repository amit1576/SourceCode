using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CoralDBE;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.IO;


namespace BalanceSheetUtility
{
    public partial class BalanceSheetUtility : Form
    {
        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
                                                          string key, string def, System.Text.StringBuilder retVal,
                                                          int size, string filePath);

        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
                                                             string key, string val, string filePath);

        DBConnection DB;

        StringBuilder m_SysDB = new StringBuilder(255);
        StringBuilder m_SQLDB = new StringBuilder(255);
        StringBuilder m_SQLServer = new StringBuilder(255);
        StringBuilder m_SQLUser = new StringBuilder(255);
        StringBuilder m_SQLPassword = new StringBuilder(255);
        string m_SQLPwd = "";

        string __sqlConnectionString;
        string __DataSourceName;

        DataSet __DataSource;

        List<string> __storedProcedureList = new List<string>();
        List<string> __targetExcelSheet = new List<string>();
        List<CreateMappingString> __ExcelColumnMappingString = new List<CreateMappingString>();

        int __isConnected = 1;

        public string SQLUser { get; set; }
        public string SQLPwd { get; set; }

        static string m_AppName = "BalanceSheetUtility";

        static string m_INIPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.ini";
        static string m_Directory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + m_AppName;

        private string m_LogFilePath = "";

        public string LogFilePath
        {
            get { return m_LogFilePath; }
            set { m_LogFilePath = value; }
        }

        private string m_LogFileDirectory = "";

        public string LogFileDirectory
        {
            get { return m_LogFileDirectory; }
            set { m_LogFileDirectory = value; }
        }

        private string m_IsActivated;

        public string IsActivated
        {
            get { return m_IsActivated; }
            set { m_IsActivated = value; }
        }

        public string m_ExecuteStoredProcedure;

        public string ExecuteStoredProcedure
        {
            get { return m_ExecuteStoredProcedure; }
            set { m_ExecuteStoredProcedure = value; }
        }

        public string m_DataSourceName;

        public string DataSourceName
        {
            get { return m_DataSourceName; }
            set { m_DataSourceName = value; }
        }

        public string m_WhereCondition;

        public string WhereCondition
        {
            get { return m_WhereCondition; }
            set { m_WhereCondition = value; }
        }

        public string m_OrderBy;

        public string OrderBy
        {
            get { return m_OrderBy; }
            set { m_OrderBy = value; }
        }

        public string m_TargetExcelSheet;

        public string TargetExcelSheet
        {
            get { return m_TargetExcelSheet; }
            set { m_TargetExcelSheet = value; }
        }

        public string m_SP;

        public string SP
        {
            get { return m_SP; }
            set { m_SP = value; }
        }

        public string m_SourceColumnToRead;

        public string SourceColumnToRead
        {
            get { return m_SourceColumnToRead; }
            set { m_SourceColumnToRead = value; }
        }

        public string m_ExcelColumnToRead;

        public string ExcelColumnToRead
        {
            get { return m_ExcelColumnToRead; }
            set { m_ExcelColumnToRead = value; }
        }

        public string m_ExcelPath;

        public string ExcelPath
        {
            get { return m_ExcelPath; }
            set { m_ExcelPath = value; }
        }

        public string m_HeaderRowNumber;

        public string HeaderRowNumbers
        {
            get { return m_HeaderRowNumber; }
            set { m_HeaderRowNumber = value; }
        }

        public BalanceSheetUtility()
        {
            InitializeComponent();
        }

        private void BalanceSheetUtility_Load(object sender, EventArgs e)
        {
            //ConnectToDatabase();
        }

        private int ConnectToDatabase()
        {
            LogProgress("Connecting to DataBase");
            //if (!String.IsNullOrEmpty(SQLUser))
            //{
            //    SQLPwd = EncryptString(SQLPwd);

            //    WritePrivateProfileString("SQLSERVER", "USER", SQLUser, t_INIFile);
            //    WritePrivateProfileString("SQLSERVER", "Password", SQLPwd, t_INIFile);
            //}

            GetPrivateProfileString("SQLSERVER", "SERVER", "", m_SQLServer, 255, m_INIPath);
            GetPrivateProfileString("SQLSERVER", "USER", "", m_SQLUser, 255, m_INIPath);
            GetPrivateProfileString("SQLSERVER", "Password", "", m_SQLPassword, 255, m_INIPath);
            GetPrivateProfileString("SQLSERVER", "Database", "", m_SQLDB, 255, m_INIPath);

            //m_SQLPwd = DecryptString(m_SQLPassword.ToString());
            m_SQLPwd = m_SQLPassword.ToString();

            __sqlConnectionString = "Provider=SQLOLEDB;Data Source=" + m_SQLServer.ToString() + ";Initial Catalog=" + m_SQLDB.ToString() + ";User ID=" + m_SQLUser + ";Password=" + m_SQLPwd + ";Appilcation Name=" + m_AppName + "";

            try
            {
                DB = new CoralDBE.DBConnection(__sqlConnectionString, "MSSQL");
                if (DB != null)
                {
                    __isConnected = 0;
                }
            }
            catch (Exception Err)
            {
                LogProgress("Error while connecting to DataBase : " + Err.Message);
                MessageBox.Show(Err.Message);
                __isConnected = 1;
                return __isConnected;
            }
            return __isConnected;

        }

        private int LoadConfiguration()
        {
            LogProgress("Loding All Configuration");

            int _ret = 0;
            string t_LogFilePath = string.Empty;
            try
            {
                string[] t_Sections = IniReadValue(m_AppName, "INISection").Split(',');
                string t_Directory = IniReadValue(m_AppName, "LogFileDirectory");
                string t_FileName = "";

                if (!Directory.Exists(t_Directory + "\\LogFiles"))
                {
                    Directory.CreateDirectory(t_Directory + "\\LogFiles");
                }

                t_FileName = m_AppName + "-" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";

                t_LogFilePath = t_Directory + "\\LogFiles\\" + t_FileName;

                BalanceSheetUtility t_LoadConfiguration = new BalanceSheetUtility();
                t_LoadConfiguration.LogFileDirectory = t_Directory;
                t_LoadConfiguration.LogFilePath = t_LogFilePath;

                t_LoadConfiguration.IsActivated = IniReadValue(m_AppName, "Activated");
                t_LoadConfiguration.ExecuteStoredProcedure = IniReadValue(m_AppName, "ExecuteStoredProcedure");
                t_LoadConfiguration.ExcelPath = IniReadValue(m_AppName, "ExcelPath");

                t_LoadConfiguration.SP = IniReadValue("StoredProcedure", "SP");

                if (t_LoadConfiguration.SP != string.Empty)
                {
                    __storedProcedureList = GetListBack(t_LoadConfiguration.SP);
                }

                t_LoadConfiguration.DataSourceName = IniReadValue("DataSource", "DataSourceName");
                t_LoadConfiguration.WhereCondition = IniReadValue("DataSource", "WhereCondition");
                t_LoadConfiguration.OrderBy = IniReadValue("DataSource", "OrderBy");
                t_LoadConfiguration.SourceColumnToRead = IniReadValue("DataSource", "SourceColumnToRead");
                t_LoadConfiguration.ExcelColumnToRead = IniReadValue("DataSource", "ExcelColumnToRead");
                t_LoadConfiguration.HeaderRowNumbers = IniReadValue("DataSource", "HeaderRowNumber");

                t_LoadConfiguration.TargetExcelSheet = IniReadValue("ExcelSheetMapping", "TargetExcelSheet");

                __DataSourceName = t_LoadConfiguration.DataSourceName;
                m_WhereCondition = t_LoadConfiguration.WhereCondition;
                m_OrderBy = t_LoadConfiguration.OrderBy;
                m_SourceColumnToRead = t_LoadConfiguration.SourceColumnToRead;
                m_ExcelColumnToRead = t_LoadConfiguration.ExcelColumnToRead;
                m_ExcelPath = t_LoadConfiguration.ExcelPath;
                m_HeaderRowNumber = (!string.IsNullOrEmpty(t_LoadConfiguration.HeaderRowNumbers)) ? (t_LoadConfiguration.HeaderRowNumbers) : "1";
                m_ExecuteStoredProcedure = t_LoadConfiguration.ExecuteStoredProcedure;

                if (t_LoadConfiguration.DataSourceName == string.Empty)
                {
                    LogProgress("Please Define the Data Source Name");
                    MessageBox.Show("Please Define the Data Source Name");
                    return _ret = 1;
                }
                if (t_LoadConfiguration.SourceColumnToRead == string.Empty)
                {
                    LogProgress("Please Define the Data Source Column Name required for Matching");
                    MessageBox.Show("Please Define the Data Source Column Name required for Matching");
                    return _ret = 1;
                }
                if (t_LoadConfiguration.ExcelColumnToRead == string.Empty)
                {
                    LogProgress("Please Define the Excel Column Name required for Matching");
                    MessageBox.Show("Please Define the Excel Column Name required for Matching");
                    return _ret = 1;
                }

                if (t_LoadConfiguration.TargetExcelSheet != string.Empty)
                {
                    __targetExcelSheet = GetListBack(t_LoadConfiguration.TargetExcelSheet);
                }

                try
                {
                    if (__targetExcelSheet.Count > 0)
                    {
                        for (int i = 0; i < __targetExcelSheet.Count; i++)
                        {
                            var _Excelsheet = __targetExcelSheet[i].ToString();
                            var _ExcelColumnMapping = IniReadValue("ExcelColumnMapping", _Excelsheet);
                            if (_ExcelColumnMapping != string.Empty)
                            {
                                __ExcelColumnMappingString.Add(new CreateMappingString { ExcelSheetName = _Excelsheet, MappingString = _ExcelColumnMapping });
                            }
                            else
                            {
                                LogProgress("Please Define ExcelColumnMapping for " + _Excelsheet);
                                MessageBox.Show("Please Define ExcelColumnMapping for " + _Excelsheet);
                                return _ret = 1;
                            }
                        }
                    }
                    else
                    {
                        LogProgress("Please Define Atleast One Excel Sheet Name in 'ExcelSheetMapping' Tag");
                        MessageBox.Show("Please Define Atleast One Excel Sheet Name in 'ExcelSheetMapping' Tag");
                        return _ret = 1;
                    }
                }
                catch (Exception EX)
                {
                    LogProgress(EX.Message);
                    MessageBox.Show(EX.Message);
                    return _ret = 1;
                }
                _ret = 0;
            }
            catch (Exception ex)
            {
                LogProgress("Error while Loading Configuration : " + ex.Message);
                ProfileLogic(m_AppName, "Error In Timer Elapsed Event: " + ex.Message + Environment.NewLine + ex.StackTrace, t_LogFilePath);
                MessageBox.Show(ex.Message);
                return _ret = 1;
            }
            return _ret;
        }

        private string ExecuteProcess()
        {
            LogProgress("Start Executing DataBase Queries:");

            DataAccess t_DataAccess = new DataAccess();
            int _result = 0;
            string _ret = "";            

            DataSet _DataSourceDataSet = new DataSet();
            try
            {
                /// Executing Stored Procedure
                if (__storedProcedureList.Count > 0 && m_ExecuteStoredProcedure == "Y")
                {
                    for (int i = 0; i < __storedProcedureList.Count; i++)
                    {
                        var _spStatement = __storedProcedureList[i].ToString().Trim();
                        LogProgress("Executing Stored Procedure : " + _spStatement);
                        if (_spStatement != string.Empty)
                        {
                            _result = t_DataAccess.ExcuteStoredProcedure(_spStatement, __sqlConnectionString);
                            if (_result == 0)
                            {
                                LogProgress("Proceudre Executed Successfully : " + _spStatement);
                                continue;
                            }
                            else
                            {
                                LogProgress("Error in Execution of Strored Procedure : " + _spStatement);
                                MessageBox.Show("Error in Execution of Strored Procedure : " + _spStatement);
                                return "Error in Execution of Strored Procedure : " + _spStatement;
                            }
                        }
                    }
                }

                /// Fetching Data From Data Source
                if (__DataSourceName != string.Empty)
                {
                    var _sqlStatement = "Select * From " + __DataSourceName;
                   
                    if (m_WhereCondition != string.Empty)
                    {
                        _sqlStatement += " Where " + m_WhereCondition;
                    }

                    if (m_OrderBy != string.Empty)
                    {
                        _sqlStatement += " Order By " + m_OrderBy;
                    }

                    LogProgress("Fetching Data from Data Source : " + _sqlStatement);

                    _result = t_DataAccess.ExecuteDataAccessStatement(_sqlStatement, __sqlConnectionString);

                    if (_result == 0)
                    {
                        _DataSourceDataSet = t_DataAccess.ReturnDataSet;
                        __DataSource = _DataSourceDataSet;
                        LogProgress("Data Fetched Successfully : ");
                    }
                    else
                    {
                        LogProgress("Error in Fetching Data from Source : " + __DataSourceName);
                        MessageBox.Show("Error in Fetching Data from Source : " + __DataSourceName);
                        return "Error in Fetching Data from Source : " + __DataSourceName;
                    }
                }
                else
                {
                    LogProgress("Please Define the Data Source Name");
                    MessageBox.Show("Please Define the Data Source Name");
                    return "Please Define the Data Source Name";
                }

                /// Writing Data to Excel File
                if (m_ExcelPath != string.Empty)
                {
                    LogProgress("Excel Process Starts");
                    _ret = ProcessExcel();
                }
                else
                {
                    LogProgress("Please Define Excel Path");
                    MessageBox.Show("Please Define Excel Path");
                    return "Please Define Excel Path";
                }
            }
            catch (Exception EX)
            {
                LogProgress("Error in Execution " + EX.Message);
                MessageBox.Show(EX.Message);
                button1.Enabled = true;
            }
            return _ret;
        }

        private string ProcessExcel()
        {
            LogProgress("Populating Balance Sheet...............");
            ExcelFile t_ExcelFile = new ExcelFile();
            string _ret = "";
            int _result = 0;
            try
            {
                if (__ExcelColumnMappingString.Count > 0)
                {
                    t_ExcelFile.excelFilePath = m_ExcelPath;
                    t_ExcelFile.HeaderRowNumber = Convert.ToInt32(m_HeaderRowNumber);
                    _result = t_ExcelFile.openExcel();
                    if (_result == 0)
                    {
                        _ret = t_ExcelFile.ManuplateExcel(__DataSource, __ExcelColumnMappingString, m_SourceColumnToRead, m_ExcelColumnToRead);                        
                        t_ExcelFile.closeExcel();
                    }
                    else
                    {
                        t_ExcelFile.closeExcel();
                        MessageBox.Show("Error While opening Excel File");
                        _ret = "Error While opening Excel File";
                        LogProgress(_ret);
                        return _ret;
                    }
                }
                else
                {
                    MessageBox.Show("No Excel Sheet for Writing Data");
                    _ret = "No Excel Sheet for Writing Data";
                    LogProgress(_ret);
                    return _ret;
                }
            }
            catch (Exception EX)
            {
                button1.Enabled = true;
                t_ExcelFile.closeExcel();
                MessageBox.Show(EX.Message);
                _ret = EX.Message.ToString();
                LogProgress(_ret);
                return _ret;                
            }
            return _ret;
        }

        private static string IniReadValue(string Section, string Key)
        {
            System.Text.StringBuilder temp = new System.Text.StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, m_INIPath);
            return temp.ToString();
        }

        static void ProfileLogic(string p_Service, string p_Statement, string p_FilePath)
        {
            string t_FilePath = p_FilePath;

            if (!File.Exists(t_FilePath))
            {
                File.Create(t_FilePath).Dispose();
            }
            StreamWriter writer = new StreamWriter(t_FilePath, true);
            writer.WriteLine("-------------------Profile Log - " + p_Service + " -----------------------------------------------");
            writer.WriteLine("Log Time: " + DateTime.Now);
            //writer.WriteLine(DateTime.Now);
            writer.WriteLine("Statement : " + p_Statement);
            //writer.WriteLine(p_Statement);
            writer.WriteLine("-------------------Profile Log - " + p_Service + " -----------------------------------------------");

            writer.Close();
            writer.Dispose();
        }

        private List<string> GetListBack(string p_String)
        {
            List<string> _returnList = new List<string>();

            _returnList = p_String.Split(',').ToList();

            return _returnList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int _result = 10;
            string _ret = "";
            try
            {
                ConnectToDatabase();

                if (__isConnected == 0)
                {
                    LogProgress("Connected to DataBase Successfully");
                    _result = LoadConfiguration();
                }
                else
                {
                    ConnectToDatabase();
                }
                if (_result == 0)
                {
                    LogProgress("All Configuration Loaded Successfully");
                    button1.Enabled = false;
                    /// Summary 
                    /// Calling Execute Process
                    _ret = ExecuteProcess();
                }
                else if (_result == 1)
                {
                    LogProgress("Some Error Comming in Loading Configuration");
                    MessageBox.Show("Some Error Comming in Loading Configuration");
                    button1.Enabled = true;
                }
                if (_ret == "OK")
                {
                    LogProgress("**********  Balance Sheet Populated Successfully *********************");
                    MessageBox.Show("Balance Sheet Populated Successfully");
                    button1.Enabled = true;
                }
                else if (_ret != "")
                {
                    LogProgress("Some Error while Populating BalanceSheet !!!" + _ret);
                    MessageBox.Show("Some Error while Populating BalanceSheet !!!" + _ret);
                    ExcelFile t_ExcelFile = new ExcelFile();
                    t_ExcelFile.closeExcel();
                    button1.Enabled = true;
                }
            }
            catch (Exception EX)
            {
                LogProgress("Error " + EX.Message);
                MessageBox.Show(EX.Message);
                button1.Enabled = true;
            }
        }

        public List<ExcelColumnMapping> GetExcelColumnMappingList(string p_ExcelSheetName, string p_String)
        {
            List<ExcelColumnMapping> _returnList = new List<ExcelColumnMapping>();
            List<string> _list1 = new List<string>();

            try
            {
                _list1 = p_String.Split('|').ToList();
                for (int i = 0; i < _list1.Count; i++)
                {
                    string[] _list2 = _list1[i].Split(':');
                    _returnList.Add(new ExcelColumnMapping() { ExcelSheetName = p_ExcelSheetName, DataBaseColumnName = _list2[0], ExcelColumnNumber = Convert.ToInt32(_list2[1]) });
                }
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
            return _returnList;
        }

        public struct ExcelColumnMapping
        {
            public string ExcelSheetName;
            public string DataBaseColumnName;
            public int ExcelColumnNumber;
        }

        public struct CreateMappingString
        {
            public string ExcelSheetName;
            public string MappingString;
        }

        public void LogProgress(string p_LogInfo)
        {
            LogBox.AppendText("-----------" + System.DateTime.Now.ToShortDateString() + " " + System.DateTime.Now.ToLongTimeString() + "--------------" + System.Environment.NewLine);
            LogBox.AppendText(p_LogInfo + Environment.NewLine);
            Application.DoEvents();
        }

        private void logclear_Click(object sender, EventArgs e)
        {
            LogBox.Clear();
        }

    }
}
