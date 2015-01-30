using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace BalanceSheetUtility
{
    public class LoadConfiguration : CoralService.BaseService
    {
        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section,
                                                          string key, string def, System.Text.StringBuilder retVal,
                                                          int size, string filePath);

        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section,
                                                             string key, string val, string filePath);

        static string m_AppName = "BalanceSheetUtility";

        static string m_INIPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.ini";
        static string m_FileName = "", m_FilePath = "";
        static string m_Directory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + m_AppName;



        public void ExecuteService()
        {
            string t_WorkFlowFilePath = "";

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

                t_WorkFlowFilePath = t_Directory + "\\LogFiles\\" + t_FileName;
                LoadConfiguration t_LoadConfiguration = new LoadConfiguration();
                t_LoadConfiguration.LogFileDirectory = t_Directory;
                t_LoadConfiguration.LogFilePath = t_WorkFlowFilePath;

                t_LoadConfiguration.TotalTry = IniReadValue(m_AppName, "TotalTry");
                t_LoadConfiguration.IsActivated = IniReadValue(m_AppName, "Activated");
                t_LoadConfiguration.FTPSourceFolder = IniReadValue(m_AppName, "FTPSourceFolder");
                t_LoadConfiguration.FTPDestinationFolder = IniReadValue(m_AppName, "FTPDestinationFolder");
                t_LoadConfiguration.FileMovement = IniReadValue(m_AppName, "FileMovement");

                for (int i = 0; i < t_Sections.Length; i++)
                {
                    t_LoadConfiguration.Connect(m_INIPath, t_Sections[i]);
                    //t_LoadConfiguration.Execute();
                }

                DataAccess dd = new DataAccess();
                dd.ExecuteDataAccessStatement("select * from usermaster", t_LoadConfiguration.DB);

            }
            catch (Exception ex)
            {
                ProfileLogic(m_AppName, "Error In Timer Elapsed Event: " +
                    ex.Message + Environment.NewLine + ex.StackTrace, t_WorkFlowFilePath);
            }
        }

        private static string IniReadValue(string Section, string Key)
        {
            System.Text.StringBuilder temp = new System.Text.StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp,255, m_INIPath);
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
            writer.WriteLine("Log Time: ");
            writer.WriteLine(DateTime.Now);
            writer.WriteLine("Statement : ");
            writer.WriteLine(p_Statement);
            writer.WriteLine("-------------------Profile Log - " + p_Service + " -----------------------------------------------");

            writer.Close();
            writer.Dispose();
        }


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

        private string m_TotalTry;

        public string TotalTry
        {
            get { return m_TotalTry; }
            set { m_TotalTry = value; }
        }

        private string m_IsActivated;

        public string IsActivated
        {
            get { return m_IsActivated; }
            set { m_IsActivated = value; }
        }

        public string m_FTPSourceFolder;

        public string FTPSourceFolder
        {
            get { return m_FTPSourceFolder; }
            set { m_FTPSourceFolder = value; }
        }

        public string m_FileMovement;

        public string FileMovement
        {
            get { return m_FileMovement; }
            set { m_FileMovement = value; }
        }

        public string m_FTPDestinationFolder;

        public string FTPDestinationFolder
        {
            get { return m_FTPDestinationFolder; }
            set { m_FTPDestinationFolder = value; }
        }
    }
}
