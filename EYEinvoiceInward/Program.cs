using System;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Collections.Generic;
using Tamir.SharpSsh;
using Tamir.Streams;
using System.Collections;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Spire.Xls;


//--------
using Renci.SshNet;
using Renci.SshNet.Sftp;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using EYEinvoicingInward;

using System.Net.Mail;

namespace EYEinvoicingInward
{
    static class Program
    {
        // DataTable _dtMTNDetails = new DataTable();

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                clsConnection.Connect();

                //clsCommon.sendEmail();
                //return;

                ClsSFTPClass oSuccessFactor = new ClsSFTPClass();
                // clsCommon.SendMailWithAttachment("", "saphelpdesk@tower-vision.com", "!!Success factor Data is imported in SAP!!", "Success factor Data imported in SAP!!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");  
                clsCommon.LogEntry("Download file from SFTP Server         ", "*********START*********");
                oSuccessFactor.GetFileFromRemortServer();
                clsCommon.LogEntry("File download completed         ", "*********END***********");
                clsCommon.LogEntry("Move file into SFTP server to Source folder", "*********START*********");
              oSuccessFactor.MoveFileBetweenRemortServerDirectory();
               clsCommon.LogEntry("Move file into SFTP server to Source folder", "*********END***********");

                //   oSuccessFactor.ReadCSVFile();
                //   oSuccessFactor.ImportDataFromClaimTable();
                // oSuccessFactor.MovedAllFiles();
                // clsCommon.SendMailWithAttachment(clsCommon.AttachSuccessFile, "saphelpdesk@tower-vision.com", " !!!Success factor Data successsfully imported in SAP  !!", "This Excel file is imported from Success factor SMTP Server  ,<br>Success factor Data imported in SAP ,<br>Excel File  is impoted from Success Factor SMTP Server!!<br>  <br><b>NOTE:</b>  This is a system generated email, do not reply to this email id.");
                return;

            }
            catch (Exception ex)
            {
                //throw ex;
                ClsSFTPClass oSuccessFactor1 = new ClsSFTPClass();
                clsCommon.LogEntry("Error:- File Move Failed:" + ex.Message, "*******FAILED:-  Main()*******");
              // oSuccessFactor1.MovedAllErrorFiles();

                //clsCommon.LogEntry(ex.Message, "Main");
                //clsCommon.SendMailWithAttachment(clsCommon.AttachErrorFile, "saphelpdesk@tower-vision.com", "!!!Error While Importing Data From SFTP  !!", "Either  Data is not available or Something wrong in Excel file ,<br> Please  contact with IT Team if data is available in excell file !!<br>  <br><b>NOTE:</b>  This is a system generated email, do not reply to this email id.");
                ////if (ex.Message == "Benefit Employee Claim ID-TAD' does not belong to table")
                ////{
                ////    //  clsCommon.SendMail("", "saphelpdesk@tower-vision.com", "Success Factor Services has been Executed No Data Found", "Success Factor Services has been  Executed No Data Found !!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");
                ////}

                ////else
                ////{
                ////    clsCommon.SendMail("", "saphelpdesk@tower-vision.com", "!!Success factor Data is not imported in SAP" + ex.Message + "!!", "Success factor Data is not imported in SAP Something wrong in data !!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");
                ////}


                //Application.Exit();
            }
            finally
            {
                //if (clsCommon.AttachSuccessFile != "" && clsCommon.MailStatus == "T")
                //{
                //    clsCommon.SendMailWithAttachment(clsCommon.AttachSuccessFile, "saphelpdesk@tower-vision.com", "!!Success factor Data is imported in SAP!!", "Success factor Data imported in SAP!!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");
                //}
                //else if (clsCommon.AttachSuccessFile != "" && clsCommon.MailStatus != "T")
                //{
                //    clsCommon.SendMailWithAttachment(clsCommon.AttachErrorFile, "saphelpdesk@tower-vision.com", "!!Success factor Data is not imported in SAP!!", "Success factor Data imported in SAP ,<br> Either Data not found in excel or Something wrong with data Or all Data is duplicate!!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");
                //}
                //else
                //{
                //    clsCommon.SendMailWithAttachment(clsCommon.AttachErrorFile, "saphelpdesk@tower-vision.com", "!!Success factor Data is not imported in SAP!!", "Success factor Data imported in SAP ,<br> Either Data not found in excel or Something wrong with data Or all Data is duplicate!!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");
                //}
                //clsCommon.SendMail("", "saphelpdesk@tower-vision.com", "!!Success factor Data is not imported in SAP" + ex.Message + "!!", "Success factor Data is not imported in SAP Something wrong in data !!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");

            }
        }

        #region MyRegion


        /* 1.---Get data from Database 2.---- E invoice Create CSV file  3.----Read CST file---------*/
        static void GetDatafromDatabase()
        {

            //IRN_Details@SAP@<any_string>.csv 
            string suBNmae = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            suBNmae = suBNmae.Replace("/", "").Replace(":", "").Replace(" ", "");
            DataTable _DT = clsCommon.ExecuteDataSetDataTable("EXEC UTL_E_Incoice_Details '20200701', '20200731'");
            string strFilePath = @"D:\Dhananjay Development Code\Code\IRN_Details@SAP@" + suBNmae + ".xls";
            //ToCSV(_DT, strFilePath);

        }

        static void ToCSV(this DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        static void ReadCSVFile()
        {


            string[] lines = File.ReadAllLines(@"D:\Dhananjay Development Code\Code\IRN_Details@SAP@19082020123408.csv");
            foreach (string line in lines)
            {

                string[] TotalFields = line.Split(',');
                string Refrence_No = TotalFields[5];

            }


        }
        #endregion
        static void GetFileListFromSFTP()
        {
            try
            {

                string _ftpURL = "sftp10.successfactors.com"; //Host URL or address of the SFTP server
                string _UserName = "10121426P";     //User Name of the SFTP server
                string _Password = "i5iQQlsl92";  //Password of the SFTP server
                int _Port = 22;                 //Port No of the SFTP server (if any)
                string _ftpDirectory = "/FEED/"; //The directory in SFTP server where the files are present
                Sftp oSftp = new Sftp(_ftpURL, _UserName, _Password);
                oSftp.Connect(_Port);
                string aa = oSftp.ClientVersion;
                ArrayList FileList = oSftp.GetFileList(_ftpDirectory);
                oSftp.Close();
            }
            catch (Exception EF)
            {

                string error = EF.Message;
                //oSftp.Close();
            }

        }

        static void GetFileFromSFTP()
        {
            string _ftpURL = "testsftp.com"; //Host URL or address of the SFTP server
            string _UserName = "admin";     //User Name of the SFTP server
            string _Password = "admin123";  //Password of the SFTP server
            int _Port = 22;                 //Port No of the SFTP server (if any)
            string _ftpDirectory = "Receipts"; //The directory in SFTP server where the files are present
            string LocalDirectory = "D:\\FilePuller"; //Local directory where the files will be downloaded

            Sftp oSftp = new Sftp(_ftpURL, _UserName, _Password);
            oSftp.Connect(_Port);
            ArrayList FileList = oSftp.GetFileList(_ftpDirectory);
            FileList.Remove(".");
            FileList.Remove("..");          //Remove . from the file list
            FileList.Remove("Processed");   //Remove folder name from the file list. If there is no folder name remove the code.

            for (int i = 0; i < FileList.Count; i++)
            {
                if (!File.Exists(LocalDirectory + "/" + FileList[i]))
                {
                    oSftp.Get(_ftpDirectory + "/" + FileList[i], LocalDirectory + "/" + FileList[i]);
                    Thread.Sleep(100);
                }
            }
            oSftp.Close();

        }


        static DataTable GetDataTableFromCsv()
        {
            string header = "Yes";
            string path = @"D:\\Dhananjay Development Code\\Code\\Testing Development\\SF\\TotalBenefitEm.csv";
           // string excelFileName = @"D:\\Dhananjay Development Code\\Code\\Testing Development\\SF\\CSVTOEXCEL.xls";

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);

            string sql = @"SELECT * FROM [" + fileName + "]";

            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                      ";Extended Properties=\"Text;HDR=" + header + "\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {
                DataTable dataTable = new DataTable();
                dataTable.Locale = CultureInfo.CurrentCulture;
                adapter.Fill(dataTable);
                return dataTable;
            }
        }

        static void ImportExcel()
        {
            try
            {

                DataTable _dtMTNDetails = new DataTable();
                // _dtMTNDetails = new DataTable();
                if (true)
                {
                    // Utilities.ShowWarningMessage("Importing data");

                    string Sql = @"SELECT * FROM [Sheet1$]";
                    string strFileName = @"D:\Dhananjay Development Code\EXCEL File\ChargeTemplate_Updated Jun20.xls"; ;

                    string oConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + strFileName + "; Extended Properties='Excel 8.0; HDR=Yes'";
                    OleDbConnection oOledbConn = new OleDbConnection(oConnectionString);
                    if (oOledbConn.State == ConnectionState.Closed) oOledbConn.Open();
                    OleDbCommand oCmdSelect = new OleDbCommand(Sql, oOledbConn);
                    OleDbDataAdapter oOledbAdapter = new OleDbDataAdapter();
                    oOledbAdapter.SelectCommand = oCmdSelect;
                    oOledbAdapter.Fill(_dtMTNDetails);



                }
                else
                {
                    //
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //public static class SftpTest
        //{
        //    private const string Host = "https://sftp10.successfactors.com";
        //    private const int Port = 22;
        //    private const string Username = "10121426P";
        //    private const string Password = "i5iQQlsl92";
        //    private const string Source = "/myfilders/tmp";
        //    private const string Destination = @"c:\temp";

        //    public static void Main()
        //    {
        //        var connectionInfo = new KeyboardInteractiveConnectionInfo(Host, Port, Username);

        //        connectionInfo.AuthenticationPrompt += delegate(object sender, AuthenticationPromptEventArgs e)
        //        {
        //            foreach (var prompt in e.Prompts)
        //            {
        //                if (prompt.Request.Equals("Password: ", StringComparison.InvariantCultureIgnoreCase))
        //                {
        //                    prompt.Response = Password;
        //                }
        //            }
        //        };

        //        using (var client = new SftpClient(connectionInfo))
        //        {
        //            client.Connect();
        //            DownloadDirectory(client, Source, Destination);
        //        }
        //    }

        //    private static void DownloadDirectory(SftpClient client, string source, string destination)
        //    {
        //        var files = client.ListDirectory(source);
        //        foreach (var file in files)
        //        {
        //            if (!file.IsDirectory && !file.IsSymbolicLink)
        //            {
        //                DownloadFile(client, file, destination);
        //            }
        //            else if (file.IsSymbolicLink)
        //            {
        //                Console.WriteLine("Ignoring symbolic link {0}", file.FullName);
        //            }
        //            else if (file.Name != "." && file.Name != "..")
        //            {
        //                var dir = Directory.CreateDirectory(Path.Combine(destination, file.Name));
        //                DownloadDirectory(client, file.FullName, dir.FullName);
        //            }
        //        }
        //    }

        //    private static void DownloadFile(SftpClient client, SftpFile file, string directory)
        //    {
        //        Console.WriteLine("Downloading {0}", file.FullName);
        //        using (Stream fileStream = File.OpenWrite(Path.Combine(directory, file.Name)))
        //        {
        //            client.DownloadFile(file.FullName, fileStream);
        //        }
        //    }
        //}
        //static void CallingMTNDetails()
        //{
        //    try
        //    {
        //        clsSendAlert oSendAlert = new clsSendAlert();
        //        while (true)
        //        {
        //            oSendAlert.UpdateMTNDetails();
        //            oSendAlert.UpdateFATaxAmount();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        clsCommon.LogEntry(ex.Message, "CallingMTNDetails");
        //        Application.Exit();
        //    }
        //}

        //static void CallingProcessGRIREntry()
        //{
        //    try
        //    {
        //        clsCommon.connectSAPCompany();
        //        clsSendAlert oSendAlert = new clsSendAlert();
        //        while (true)
        //        {

        //            oSendAlert.ProcessGRIREntry();

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        clsCommon.LogEntry(ex.Message, "ProcessGRIREntry");
        //        Application.Exit();
        //    }
        //}


        /// <summary>
        /// List a remote directory in the console.
        /// </summary>
        static void listFiles()
        {
            string host = @"sftp10.successfactors.com";
            string username = "10121426P";
            string password = "i5iQQlsl92";

            string remoteDirectory = "/FEED";

            using (SftpClient sftp = new SftpClient(host, username, password))
            {
                try
                {
                    sftp.Connect();

                    var files = sftp.ListDirectory(remoteDirectory);

                    foreach (var file in files)
                    {
                        Console.WriteLine(file.Name);
                    }

                    sftp.Disconnect();
                    Console.ReadLine();
                }
                catch (Exception e)
                {
                    Console.WriteLine("An exception has been caught " + e.ToString());
                }
            }
        }

        /// <summary>
        /// Downloads a remote directory into a local directory
        /// </summary>
        /// <param name="client"></param>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        static void DownloadDirectory(SftpClient client, string source, string destination, bool recursive = false)
        {
            // List the files and folders of the directory
            var files = client.ListDirectory(source);

            // Iterate over them
            foreach (SftpFile file in files)
            {
                // If is a file, download it
                if (!file.IsDirectory && !file.IsSymbolicLink)
                {
                    DownloadFile(client, file, destination);
                }
                // If it's a symbolic link, ignore it
                else if (file.IsSymbolicLink)
                {
                    Console.WriteLine("Symbolic link ignored: {0}", file.FullName);
                }
                // If its a directory, create it locally (and ignore the .. and .=) 
                //. is the current folder
                //.. is the folder above the current folder -the folder that contains the current folder.
                else if (file.Name != "." && file.Name != "..")
                {
                    var dir = Directory.CreateDirectory(Path.Combine(destination, file.Name));
                    // and start downloading it's content recursively :) in case it's required
                    if (recursive)
                    {
                        DownloadDirectory(client, file.FullName, dir.FullName);
                    }
                }
            }
        }

        /// <summary>
        /// Downloads a remote file through the client into a local directory
        /// </summary>
        /// <param name="client"></param>
        /// <param name="file"></param>
        /// <param name="directory"></param>
        static void DownloadFile(SftpClient client, SftpFile file, string directory)
        {
            //Console.WriteLine("Downloading {0}", file.FullName);

            using (Stream fileStream = File.OpenWrite(Path.Combine(directory, file.Name)))
            {
                client.DownloadFile(file.FullName, fileStream);
            }
        }

        /// <summary>
        /// Downloads a file in the desktop synchronously
        /// </summary>
        static void downloadFile()
        {
            string host = @"yourSftpServer.com";
            string username = "root";
            string password = @"p4ssw0rd";

            // Path to file on SFTP server
            string pathRemoteFile = "/var/www/vhosts/some-folder/file_server.txt";
            // Path where the file should be saved once downloaded (locally)
            string pathLocalFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "download_sftp_file.txt");

            using (SftpClient sftp = new SftpClient(host, username, password))
            {
                try
                {
                    sftp.Connect();

                    Console.WriteLine("Downloading {0}", pathRemoteFile);

                    using (Stream fileStream = File.OpenWrite(pathLocalFile))
                    {
                        sftp.DownloadFile(pathRemoteFile, fileStream);
                    }

                    sftp.Disconnect();
                }
                catch (Exception er)
                {
                    Console.WriteLine("An exception has been caught " + er.ToString());
                }
            }
        }



    }
}
