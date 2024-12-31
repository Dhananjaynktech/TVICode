using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Configuration;


namespace EYEinvoicingInward
{
    public static class clsConnection
    {
        #region GLOBAL VARIABLES
        public static string gPath = String.Empty;
        public static string gFirstParameter = String.Empty; // To keep first parameter
        public static SqlConnection gSQLCon = null, gSQLCon_GRIR = null; // To keep sql connection
        public static DataTable gDT_AlertSettings = new DataTable();
        public static string gConnectionString = String.Empty; // To keep connection string
        public static string[] gDBUserDetails = null;
        public static string gDocEntry = "";
        public static string gDocType = "";
        public static SAPbobsCOM.Company gCompany = null;
        #endregion

        #region CLASS LEVEL VARIABLE DECLARATION
        static string _Server = String.Empty;
        static string _UserId = String.Empty;
        static string _UserName = String.Empty;
        static string _Password = String.Empty;
        static string _Database = String.Empty;
        #endregion

        #region clsFunctions

        /// <summary>
        /// Funtion to set the sql connection settings
        /// </summary>
        public static void Connect()
        {
            try
            {
                //1. Checking whether DBLogin.INI file exists or not
                //2. Checking whether DBLogin.INI file is having valid login id and password  

                //Getting Login ID and Password from DBLogin.INI
               // _Server = "sapnew";
               // _UserId = "sa";
               // _Password = "B1admin";
               //// _Database = "TVI_PilotCompany";
               //   _Database = "ZTest_03_08_2020_TVI_PilotCompany";

                _Server = ConfigurationManager.AppSettings["Server"].ToString(); //"sapnew";
                // _Database = "TVI_PilotCompany";
                _Database = ConfigurationManager.AppSettings["Database"].ToString();//"ZTest_03_08_2020_TVI_PilotCompany";
                _UserId = ConfigurationManager.AppSettings["UserID"].ToString();// "sa";
                _Password = ConfigurationManager.AppSettings["Password"].ToString();// "B1admin";
                //------------------------------------------------------------------------------------

                //Initialize ADO.Net Connection Object
                gConnectionString = "Server=" + _Server + ";Database=" + _Database + ";UID=" + _UserId + ";PWD=" + _Password + ";MultipleActiveResultSets=true";
                gSQLCon = new SqlConnection(gConnectionString);
                gSQLCon_GRIR = new SqlConnection(gConnectionString);
                //--------------------------------------------------------------------------------------

                // Setting CRDispaly Details
                //CRDisplay.clsCRDisplay.SQLConnection = clsConnection.gSQLCon;
                //CRDisplay.clsCRDisplay.Database = _Database;
                //CRDisplay.clsCRDisplay.UserID = _UserId;
                //CRDisplay.clsCRDisplay.Password = _Password;

                // Creating temp folder
                gPath = System.Windows.Forms.Application.StartupPath + @"\Temp";
                if (Directory.Exists("Temp") == false)
                {
                    Directory.CreateDirectory(gPath);
                }
                clsCommon.LogEntry("", "");
                clsCommon.LogEntry("################NEW CONNECTION###############***********************", "****************************");
                clsCommon.LogEntry(": " + _Database + "", "Connected With Database");
            }
            catch (Exception ex)
            {
                clsCommon.LogEntry(ex.Message, "Connect");
                Application.Exit();
            }
        }
        #endregion
    }
}