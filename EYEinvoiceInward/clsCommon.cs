using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using System.Data;
using System.Net.Mail;
using System.Reflection;
using System.Net.Security;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Collections;
using System.Threading;
//using System.IO;
using Tamir.SharpSsh;
using Tamir.Streams;
using System.Configuration;

namespace EYEinvoicingInward
{
    class clsCommon
    {

        // Global Variables
        public static string gJobName = String.Empty;

        public static string MailSubject = String.Empty;
        public static string MailBody = String.Empty;
        public static string MailStatus = String.Empty;
        public static string AttachSuccessFile = String.Empty;
        public static string AttachErrorFile = String.Empty;

        // public static string MailBody = String.Empty;


        /// <summary>
        /// Accepts a sql query as a parameter executes it which returns data in data set
        /// </summary>
        /// <param name="Sql"></param>
        /// <returns></returns>
        public static DataSet ExecuteDataSet(string Sql)
        {
            // Class level variable declaration
            DataSet dataSet = null; // Decalaring DataTable to return DataTable
            //------------------------------------------------------------------------------------
            try
            {
                dataSet = new DataSet(); // Initializing data table object
                if (clsConnection.gSQLCon.State == ConnectionState.Closed) clsConnection.gSQLCon.Open();
                SqlCommand cmd = new SqlCommand(Sql, clsConnection.gSQLCon);
                cmd.CommandTimeout = 10000000;
                SqlDataAdapter oAd = new SqlDataAdapter(cmd);
                oAd.Fill(dataSet);
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ExecuteDataSet");  /* Throwing error message */

                Application.Exit();
            }
            finally
            {
                if (clsConnection.gSQLCon.State == ConnectionState.Open)
                    clsConnection.gSQLCon.Close();
            }
            return dataSet; // Returning data set 
        }

        /// <summary>
        /// Accepts a sql query as a parameter executes it which returns a single value & returns its value as a string
        /// </summary>
        /// <param name="Sql"></param>
        /// <returns></returns>
        public static string ExecuteScalarValue_String(string Sql)
        {
            // Class level variable declaration
            string scalarvalue = String.Empty; // Decalaring string variable to return value
            Object obj = null; // Declararing a variable Object 
            //------------------------------------------------------------------------------------

            try
            {
                if (clsConnection.gSQLCon.State == ConnectionState.Closed) clsConnection.gSQLCon.Open();
                SqlCommand cmd = new SqlCommand(Sql, clsConnection.gSQLCon);
                cmd.CommandTimeout = 10000000;
                obj = cmd.ExecuteScalar();

                if (obj != null && !String.IsNullOrEmpty(obj.ToString()))
                    scalarvalue = obj.ToString();

                cmd.Dispose();
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ExecuteScalarValue_String"); /* Throwing error message */
                Application.Exit();
            }
            finally
            {
                if (clsConnection.gSQLCon.State == ConnectionState.Open)
                    clsConnection.gSQLCon.Close();
            }
            return scalarvalue; // returing string value after executing passed query
        }

        /// <summary>
        /// Accepts a sql query as a parameter executes it which returns data in data table
        /// </summary>
        /// <param name="Sql"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataSet_DataTable(string Sql, SqlConnection sqlCon)
        {
            // Class level variable declaration
            DataTable dataTable = null; // Decalaring DataTable to return DataTable
            //------------------------------------------------------------------------------------
            try
            {
                dataTable = new DataTable(); // Initializing data table object
                if (sqlCon.State == ConnectionState.Closed) sqlCon.Open();
                SqlCommand cmd = new SqlCommand(Sql, sqlCon);
                cmd.CommandTimeout = 10000000;
                SqlDataAdapter oAd = new SqlDataAdapter(cmd);
                oAd.Fill(dataTable);
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ExecuteDataTable"); /* Throwing error message */
            }
            finally
            {
                if (sqlCon.State == ConnectionState.Open)
                    sqlCon.Close();
            }
            return dataTable; // Returning data set 
        }

        /// <summary>
        /// Accepts a sql query as a parameter executes it which returns data in data table
        /// </summary>
        /// <param name="Sql"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataSetDataTable(string Sql)
        {
            // Class level variable declaration
            DataTable dataTable = null; // Decalaring DataTable to return DataTable
            //------------------------------------------------------------------------------------
            try
            {
                dataTable = new DataTable(); // Initializing data table object
                if (clsConnection.gSQLCon.State == ConnectionState.Closed) clsConnection.gSQLCon.Open();
                SqlCommand cmd = new SqlCommand(Sql, clsConnection.gSQLCon);
                cmd.CommandTimeout = 10000000;
                SqlDataAdapter oAd = new SqlDataAdapter(cmd);
                oAd.Fill(dataTable);
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ExecuteDataTable"); /* Throwing error message */
            }
            finally
            {
                if (clsConnection.gSQLCon.State == ConnectionState.Open)
                    clsConnection.gSQLCon.Close();
            }
            return dataTable; // Returning data set 
        }

        /// <summary>
        /// Accepts a sql query as a parameter executes it which returns a single value & returns its value as a string
        /// </summary>
        /// <param name="Sql"></param>
        /// <returns></returns>
        public static bool ExecuteNonQuery(string Sql)
        {
            try
            {
                if (!String.IsNullOrEmpty(Sql)) // Checking query to be executed should not be blank
                {
                    if (clsConnection.gSQLCon.State == ConnectionState.Closed) clsConnection.gSQLCon.Open();
                    SqlCommand cmd = new SqlCommand(Sql, clsConnection.gSQLCon);
                    cmd.CommandTimeout = 10000000;
                    int aa = cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ExecuteNonQuery"); /* Throwing error message */
                Application.Exit();
            }
            finally
            {
                if (clsConnection.gSQLCon.State == ConnectionState.Open)
                    clsConnection.gSQLCon.Close();
            }
            return true; // returing string value after executing passed query
        }

        public static bool ExecuteNonQuery_II(string Sql)
        {
            try
            {
                if (!String.IsNullOrEmpty(Sql)) // Checking query to be executed should not be blank
                {
                    if (clsConnection.gSQLCon.State == ConnectionState.Closed) clsConnection.gSQLCon.Open();
                    SqlCommand cmd = new SqlCommand(Sql, clsConnection.gSQLCon);
                    cmd.CommandTimeout = 10000000;
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ExecuteNonQuery"); /* Throwing error message */
                Application.Exit();
            }
            finally
            {
                if (clsConnection.gSQLCon.State == ConnectionState.Open)
                    clsConnection.gSQLCon.Close();
            }
            return true; // returing string value after executing passed query
        }

        public static bool SendMail1(string To, string Subject, string Message, string AttachFile1, string AttachFile2)
        {
            // Local Variable
            Attachment oAttch = null;
            //===================================================================
            try
            {
                MailMessage oMail = new MailMessage();
                SmtpClient oSmtp = null;

                oSmtp = new SmtpClient("smtp.gmail.com", 25);

                oSmtp.Credentials = new System.Net.NetworkCredential("swati.gupta@ksetechnologies.in", "dolandduck");
                oMail.From = new MailAddress("swati.gupta@ksetechnologies.in");

                oSmtp.EnableSsl = true;
                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

                foreach (string ToReceipient in To.Split(','))
                {
                    oMail.To.Add(new MailAddress(ToReceipient.Trim()));
                }
                oMail.Subject = Subject;
                oMail.Body = Message;
                oMail.IsBodyHtml = true;

                if (!String.IsNullOrEmpty(AttachFile1))
                {
                    oAttch = new Attachment(AttachFile1);
                    oMail.Attachments.Add(oAttch);
                }
                oSmtp.Send(oMail);
                return true;
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, " E-Mail");
                return false;
            }
        }

        public static void LogEntry2(string Error, string ProcessName)
        {
            try
            {
                string oPath = System.Windows.Forms.Application.StartupPath + @"\SysLogs";
                if (Directory.Exists("SysLogs") == false)
                {
                    Directory.CreateDirectory(oPath);
                }

                oPath += @"\AlertErrorLog.log";
                StreamWriter sw = new StreamWriter(oPath, true);
                sw.WriteLine(System.DateTime.Now + "-----Error:" + ProcessName + ": " + Error);
                sw.Close();
                Application.Exit();
            }
            catch 
            {
            }
        }
        public static void LogEntry(string messages, string ProcessName)
        {
            try
            {
                string oPath = System.Windows.Forms.Application.StartupPath + @"\SysLogs\" + System.DateTime.Now.ToString("ddMMyyyy") + "";
                if (Directory.Exists(oPath) == false)
                {
                    Directory.CreateDirectory(oPath);
                }

                oPath += @"\" + System.DateTime.Now.ToString("ddMMyyyy") + ".log";
                StreamWriter sw = new StreamWriter(oPath, true);
                sw.WriteLine(System.DateTime.Now + "---Messages--:" + ProcessName + ": " + messages);
                sw.Close();
                Application.Exit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DateTime ToDate(string sDate)
        {
            try
            {
                sDate = sDate.Trim().Insert(4, "/").Insert(7, "/");
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ToDate"); /* Throwing error message */
            }
            return DateTime.Parse(sDate);
        }
        public static string ToDateString(string sDate, string ClaimID)
        {
            string date = "";
            try
            {

                string[] dateArray = sDate.Split('/');
                date = dateArray[1] + "/" + dateArray[0] + "/" + dateArray[2];
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "ToDate"); /* Throwing error message */
            }
            if (ClaimID == "36729")
            {
              ///  string aa = "";
            }
            return date;
        }

        public static string connectSAPCompany()
        {
            int retValue = 0;
            string retMsg = String.Empty;
            try
            {
                clsConnection.gCompany = new SAPbobsCOM.Company();
                clsConnection.gCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                clsConnection.gCompany.Server = "sapnew";
                clsConnection.gCompany.CompanyDB = "TVI_PilotCompany";
                clsConnection.gCompany.DbUserName = "sa";
                clsConnection.gCompany.DbPassword = "B1admin";
                clsConnection.gCompany.UserName = "Sachi B";
                clsConnection.gCompany.Password = "Smartsb@#5";
                clsConnection.gCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
                clsConnection.gCompany.UseTrusted = false;

                retValue = clsConnection.gCompany.Connect();

                if (retValue != 0)
                {
                    clsConnection.gCompany.GetLastError(out retValue, out retMsg);
                    clsCommon.LogEntry(retMsg, "ConnectSAP");
                    return retValue + ": " + retMsg;
                }
                else
                    return "Connected";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static void ExecuteNonQuery_recordSet(string oSql)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = (SAPbobsCOM.Recordset)clsConnection.gCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(oSql);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oRecordSet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static string getApplicationPath()
        {
            string sPath = String.Empty;
            try
            {
                sPath = System.Windows.Forms.Application.StartupPath.Trim();
            }
            catch (Exception ex)
            {
                LogEntry(ex.Message, "getApplicationPath");
            }
            return sPath;
        }

        public static string getCurrentPeriod()
        {
            try
            {
                return ExecuteScalarValue_String("SELECT TOP 1 AbsEntry FROM OFPR (NOLOCK) WHERE CONVERT(VARCHAR, GETDATE(), 112) BETWEEN F_RefDate AND T_RefDate");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GetSeries(string objectType, string type)
        {
            DataTable dtSeries_Prefix = new DataTable();
            try
            {
                return ExecuteScalarValue_String("EXEC TVIPL_PROC_GetDefaultSeries '" + objectType + "', '" + type + "'");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void GetCurrentUserDetails(out string InternalK, out string SAPUserName, out string Dept)
        {
            InternalK = String.Empty;
            SAPUserName = String.Empty;
            Dept = String.Empty;
            DataTable dtUserDetails = new DataTable();
            try
            {
                dtUserDetails = ExecuteDataSetDataTable(@"SELECT T0.Internal_K, T0.U_Name, T1.Name Dept FROM OUSR T0 (NOLOCK) INNER JOIN OUDP T1 (NOLOCK) ON T0.Department = T1.Code  WHERE User_Code = 'M4U'");
                if (dtUserDetails.Rows.Count > 0)
                {
                    InternalK = dtUserDetails.Rows[0]["Internal_K"].ToString();
                    SAPUserName = dtUserDetails.Rows[0]["U_Name"].ToString();
                    Dept = dtUserDetails.Rows[0]["Dept"].ToString();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        #region Get Max Column Value
        public static string getMaxColumnValue(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRS = null;
            string sSQL, sCode;
            int MaxCode;

            sSQL = "SELECT MAX(CAST(" + Column + " AS Numeric)) FROM [" + Table + "]";
            // ExecuteNonQuery(ref oRS, sSQL);

            if (Convert.ToString(oRS.Fields.Item(0).Value).Length > 0)
                MaxCode = int.Parse(oRS.Fields.Item(0).Value.ToString()) + 1;
            else
                MaxCode = 1;

            sCode = MaxCode.ToString("00000000");

            return sCode;

        }
        #endregion

        #region Send Mail

        public static void sendEmail()
        {
            String userName = "sapb1system@tower-vision.com";
            String password = "Tvipl@420";
            MailMessage msg = new MailMessage("sapb1system@tower-vision.com", "saphelpdesk@tower-vision.com");
            msg.Subject = "Testing mail through Office365";

            msg.Body = "<html><head></head><body>" +
                                "Dear Sir <br><br> This is  test mail" +


                                "<br><br>---------------------<br>Thanks & regards<BR><BR>Tower Vision India Pvt. Ltd. " +
                                "<br><br><br><br><i><b>>" +

                                "</body></html>";

            msg.IsBodyHtml = true;
            SmtpClient SmtpClient = new SmtpClient();
            SmtpClient.Credentials = new System.Net.NetworkCredential(userName, password);
            SmtpClient.Host = "smtp.office365.com";
            SmtpClient.Port = 587;

            SmtpClient.EnableSsl = true;
            SmtpClient.Send(msg);
        }

        public static void SendMail(string UserOrSystem, string MailTo, string MailSubject, string MailBody)
        {
            try
            {


                /// String userName = "sapb1system@tower-vision.com";
                string userName = "sapb1system@tower-vision.com";
                // String password = "Tvipl@420";
                string password = "Tvipl@420";// ConfigurationManager.AppSettings["Password"].ToString();
                // string m = HttpContext.Current.Session["UserMailID"].ToString();
                MailMessage eMail;
                if (UserOrSystem == "U")
                {
                    // eMail = new MailMessage(HttpContext.Current.Session["UserMailID"].ToString(), MailTo);
                    //eMail = new MailMessage(HttpContext.Current.Session["UserMailID"].ToString(), MailTo);
                    eMail = new MailMessage(userName, MailTo);
                }
                else
                {
                    eMail = new MailMessage(userName, MailTo);
                }

                eMail.Subject = MailSubject;
                eMail.Body = MailBody;
                eMail.IsBodyHtml = true;
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Credentials = new System.Net.NetworkCredential(userName, password);
                SmtpClient.Host = "smtp.office365.com";
                SmtpClient.Port = 587;

                SmtpClient.EnableSsl = true;
                SmtpClient.Send(eMail);
            }
            catch (Exception ex)
            {

                ex.ToString();
            }
        }

        public static void SendMailWithAttachment(string UserOrSystem, string MailTo, string MailSubject, string MailBody)
        {
            try
            {

                string AppLocation = "";
                AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                AppLocation = AppLocation.Replace("file:\\", "");
                string ExcelFile = ConfigurationManager.AppSettings["ExcelFile"].ToString();
                string file = UserOrSystem + @"\" + ExcelFile;
                string userName = "sapb1system@TowerVisionIndia1.onmicrosoft.com";
                string password = "Tvipl@420";
                // string m = HttpContext.Current.Session["UserMailID"].ToString();
                MailMessage eMail;

                string EmailTo = ConfigurationManager.AppSettings["EmailTo"].ToString();
                eMail = new MailMessage(userName,EmailTo);
                // mail.To.Add(MailTo);                            // Sending MailTo

                List<string> li = new List<string>();
                string CC1 = ConfigurationManager.AppSettings["EmailCC1"].ToString();
                li.Add(CC1);
                if (ConfigurationManager.AppSettings["EmailCC2"].ToString()!="")
                {
                    li.Add(ConfigurationManager.AppSettings["EmailCC2"].ToString()); 
                }
                if (ConfigurationManager.AppSettings["EmailCC3"].ToString() != "")
                {
                    li.Add(ConfigurationManager.AppSettings["EmailCC3"].ToString());
                }
                if (ConfigurationManager.AppSettings["EmailCC4"].ToString() != "")
                {
                    li.Add(ConfigurationManager.AppSettings["EmailCC4"].ToString());
                }
                
                //li.Add("saihacksoft@gmail.com");
                //li.Add("saihacksoft@gmail.com");
                //li.Add("saihacksoft@gmail.com");

                eMail.CC.Add(string.Join<string>(",", li));
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(file); //Attaching File to Mail
                eMail.Attachments.Add(attachment);

                eMail.Subject = MailSubject;
                eMail.Body = MailBody;
                eMail.IsBodyHtml = true;
                SmtpClient SmtpClient = new SmtpClient();
                SmtpClient.Credentials = new System.Net.NetworkCredential(userName, password);
                SmtpClient.Host = "smtp.office365.com";
                SmtpClient.Port = 587;

                SmtpClient.EnableSsl = true;
                SmtpClient.Send(eMail);
            }
            catch (Exception ex)
            {

                ex.ToString();
            }
        }
        #endregion
    }
}