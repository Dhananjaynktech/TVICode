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
using EYEinvoicingInward;
using System.Configuration;

using System.Runtime.InteropServices;


//--------
using Renci.SshNet;
using Renci.SshNet.Sftp;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Text;
using System.Runtime.InteropServices;

namespace EYEinvoicingInward
{

    public class ClsSFTPClass
    {
        #region CLASS LEVEL DECLARATION

        Excel.Workbook xlWorkBook;
        Excel.Application xlApp;
        Excel.Worksheet xlWorkSheet;
        Excel.Range Startrange;
        Excel.Range HeaderStartrange;

        string _errorMsg = String.Empty;
        string _errorMessage = String.Empty;
        DataTable _DataTable = new DataTable();
        //  DataTable _dtJEDetails = new DataTable();
        StringBuilder sbJEHeader = new StringBuilder();
        StringBuilder sbJELines = new StringBuilder();
        //SAPbobsCOM.JournalEntries _journalEntries = null;
        #endregion
        #region CLASS LEVEL VARIABLES[

        //int _entryIDINS = 1;
        //bool _isFADocQueryInit = false;
        string _internal_K = String.Empty;
        string _userName = String.Empty;
        string _dept = String.Empty;

        // DataTable _dtImportDetails = new DataTable();
        DataTable _dtFADesignation = new DataTable();
        DataTable _dtFAGroups = new DataTable();

        DataTable _dtMTNDetails = new DataTable();
        DataTable _DtFileName = new DataTable("MyTable");

        Dictionary<string, string> _dictManagementType = new Dictionary<string, string>();

        StringBuilder _sbFAInsertQuery_1 = new StringBuilder();
        StringBuilder _sbFADocumentInsertQuery_1 = new StringBuilder();


        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        private static OpenFileDialog openFileDialog = new OpenFileDialog();

        string filePath = String.Empty;
        #endregion

        /// <summary>
        /// Downloads a remote directory into a local directory
        public void GetFileFromRemortServer()
        {
            try
            {
                //Thread myThread = new System.Threading.Thread(delegate()
               // Excel.Application excelApp = new Excel.Application();
                //excelApp.DisplayAlerts = true;
                //{
                string host = ConfigurationManager.AppSettings["SFTPURL"].ToString(); 
                string username = ConfigurationManager.AppSettings["SFTPUID"].ToString();
                string password = ConfigurationManager.AppSettings["SFTPPWD"].ToString();//"i5iQQlsl92";
                int Port = 22;
                // Path to folder on SFTP server
                string pathRemoteDirectory = ConfigurationManager.AppSettings["EinvoiceResponse"].ToString();
                string DestpathRemoteDirectory = ConfigurationManager.AppSettings["EinvoiceArchive"].ToString();
                // Path where the file should be saved once downloaded (locally)
                // string pathLocalDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Success Factot");
                string pathLocalDirectory = ConfigurationManager.AppSettings["Download"].ToString();// @"D:\\Dhananjay Development Code\\Code\\Testing Development\\SF";
               // string pathLocalDirectory2 = ConfigurationManager.AppSettings["Download2"].ToString();
                using (SftpClient sftp = new SftpClient(host, Port, username, password))
                {
                    try
                    {
                        sftp.Connect();

                        // By default, the method doesn't download subdirectories
                        // therefore you need to indicate that you want their content too
                        bool recursiveDownload = true;
                        // Start download of the directory
                        DownloadDirectory(sftp, pathRemoteDirectory, pathLocalDirectory, recursiveDownload);
                      //  DownloadDirectory(sftp, pathRemoteDirectory, pathLocalDirectory2, recursiveDownload);
                        /// <summary>
                        /// <Move file>
                        #region Move file
                        //SftpFile eachRemoteFile = sftp.ListDirectory(pathRemoteDirectory).FirstOrDefault();//Get first file in the Directory            
                        //if (eachRemoteFile.IsDirectory)//Move only file
                        //{
                        //    bool eachFileExistsInArchive = CheckIfRemoteFileExists(sftp, pathRemoteDirectory, eachRemoteFile.Name);

                        //    //MoveTo will result in error if filename alredy exists in the target folder. Prevent that error by cheking if File name exists
                        //    string eachFileNameInArchive = eachRemoteFile.Name;

                        //    if (eachFileExistsInArchive)
                        //    {
                        //        eachFileNameInArchive = eachFileNameInArchive + "_" + DateTime.Now.ToString("MMddyyyy_HHmmss");//Change file name if the file already exists
                        //    }
                        //    eachRemoteFile.MoveTo(DestpathRemoteDirectory + eachFileNameInArchive);
                        //}
                        #endregion
                        sftp.Disconnect();
                    }
                    catch (Exception er)
                    {
                        Console.WriteLine("An exception has been caught " + er.ToString());
                    }
                }

                //**************When SFTP has multifactor authentication<!--- Privatekey---->*************
                //var keybInterMethod = new KeyboardInteractiveAuthenticationMethod(username);
                //keybInterMethod.AuthenticationPrompt +=
                //    (sender, e) => { e.Prompts.First().Response = password; };

                //AuthenticationMethod[] methods = new AuthenticationMethod[] {
                //           new PrivateKeyAuthenticationMethod(username, new PrivateKeyFile(PrivateKey)), keybInterMethod
                //     };
                //ConnectionInfo connectionInfo = new ConnectionInfo(host, username, methods);

                //using (var sftp = new SftpClient(connectionInfo))
                //{
                //    sftp.Connect();
                //    bool recursiveDownload = true;
                //    DownloadDirectory(sftp, pathRemoteDirectory, pathLocalDirectory, recursiveDownload);
                //    // ...
                //}
                //********************************

            }
            catch (Exception ex)
            {

                clsCommon.LogEntry( "Error:- File move successfully:" + ex.Message, "*******FAILED: -GetFileFromRemortServer()*******");
            }

        }

        public void MoveFileBetweenRemortServerDirectory()
        {
            string host = ConfigurationManager.AppSettings["SFTPURL"].ToString();
            string username = ConfigurationManager.AppSettings["SFTPUID"].ToString();
            string password = ConfigurationManager.AppSettings["SFTPPWD"].ToString();//"i5iQQlsl92";
            int Port = 22;
            // Path to folder on SFTP server
            string pathRemoteDirectory = ConfigurationManager.AppSettings["EinvoiceResponse"].ToString();
            string DestpathRemoteDirectory = ConfigurationManager.AppSettings["EinvoiceArchive"].ToString();
           // string ArchuveDirectory = ConfigurationManager.AppSettings["EinvoiceArchive"].ToString();
            // Path where the file should be saved once downloaded (locally)
            // string pathLocalDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Success Factot");
            string pathLocalDirectory = ConfigurationManager.AppSettings["Download"].ToString();// @"D:\\Dhananjay Development Code\\Code\\Testing Development\\SF";
            //using (SftpClient sftp = new SftpClient(host, Port, username, password))
            //{
            //    try
            //    {
            //        sftp.Connect();
            //        SftpFile eachRemoteFile = sftp.ListDirectory(pathRemoteDirectory).FirstOrDefault();//Get first file in the Directory            
            //        if (eachRemoteFile.IsDirectory)//Move only file
            //        {
            //            bool eachFileExistsInArchive = CheckIfRemoteFileExists(sftp, pathRemoteDirectory, eachRemoteFile.Name);

            //            //MoveTo will result in error if filename alredy exists in the target folder. Prevent that error by cheking if File name exists
            //            string eachFileNameInArchive = eachRemoteFile.Name;
            //            if (eachFileExistsInArchive)
            //            {
            //                eachFileNameInArchive = eachFileNameInArchive + "_" + DateTime.Now.ToString("MMddyyyy_HHmmss");//Change file name if the file already exists
            //            }
            //            eachRemoteFile.MoveTo(DestpathRemoteDirectory + eachFileNameInArchive);
            //        }
            //    }
            //    catch (Exception)
            //    {

            //        throw;
            //    }
            //}

            try
            {
                using (SftpClient sftp = new SftpClient(host, Port, username, password))
                {
                    sftp.Connect();
                    var files = sftp.ListDirectory(pathRemoteDirectory);
                    string fileName = "";
                    foreach (SftpFile file in files)
                    {
                        string name = file.Name;
                        try
                        {

                            if (name.Contains("ISR_"))
                            {
                               // string fname = _DtFileName.Select("FileName = '" + name + "'").AsEnumerable().Distinct().ToString();

                             var fname =   (from row in _DtFileName.Select("FileName = '" + name + "'").AsEnumerable()
                                            select row.Field<string>("FileName")).Distinct();

                             foreach (string Items in fname)
                             {
                                 fileName = Items;
                             }

                             if (name == fileName)
                                {
                                    file.MoveTo(DestpathRemoteDirectory + file.Name);
                                    clsCommon.LogEntry("file moved to SFTP Archived folder:-" + file.Name, "*******File Move*******");
                                }

                            }
                        }
                        catch (Exception exp)
                        {
                            if (exp.Message == "No such file" && name.Contains("_ERR"))
                            {
                                file.Delete();
                                clsCommon.LogEntry("******File Name  :" + file.Name, "******#  Duplicate file deleted #*******:-" + pathRemoteDirectory + "/" + file);
                            }
                            else
                                clsCommon.LogEntry("******File Name  :" + file.Name, "******#  ERROR1 #*******:-" + exp.Message);
                            continue;
                        }
                    }
                }
            }
            catch
            {

                clsCommon.LogEntry("*******##########*******", "******#  ERROR #*******:-When move file moving into Archive folder");

            }
            finally
            {
                _DtFileName.Dispose(); 
            }
        }
        // <summary>
        /// Checks if Remote folder contains the given file name
        /// </summary>
        public bool CheckIfRemoteFileExists(SftpClient sftpClient, string remoteFolderName, string remotefileName)
        {
            bool isFileExists = sftpClient
                                .ListDirectory(remoteFolderName)
                                .Any(
                                        f => f.IsRegularFile &&
                                        f.Name.ToLower() == remotefileName.ToLower()
                                    );
            return isFileExists;
        }

        /// <summary>
        /// Downloads a remote directory into a local directory
        /// </summary>
        /// <param name="client"></param>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        public void DownloadDirectory(SftpClient client, string source, string destination, bool recursive = false)
        {
            // List the files and folders of the directory
            var files = client.ListDirectory(source);
            // Date 11 Nov 2024 to prevent archieve file without download  in LOcal system
           
            DataColumn column1 = new DataColumn("FileName", typeof(string));
            _DtFileName.Columns.Add(column1);

            // Iterate over them
            foreach (SftpFile file in files)
            {
               
                // Create a new row.
                DataRow row = _DtFileName.NewRow();
                // Add the row to the table.
               

                // If is a file, download it
                if (!file.IsDirectory && !file.IsSymbolicLink)
                {
                    row["FileName"] = file.Name;
                    DownloadFile(client, file, destination);
                    _DtFileName.Rows.Add(row);
                }
                // If it's a symbolic link, ignore it
                else if (file.IsSymbolicLink)
                {
                    Console.WriteLine("Symbolic link ignored: {0}", file.FullName);
                }
                // If its a directory, create it locally (and ignore the .. and .=) 
                //. is the current folder
                //.. is the folder above the current folder -the folder that contains the current folder.
                else if (file.Name != "." && file.Name != ".." && file.Name != "EINV_IRN_JSON")
                {
                    var dir = Directory.CreateDirectory(Path.Combine(destination, file.Name));
                    // and start downloading it's content recursively :) in case it's required
                    if (recursive)
                    {
                        row["FileName"] = file.Name;
                        DownloadDirectory(client, file.FullName, dir.FullName);
                        _DtFileName.Rows.Add(row);
                    }
                }

               
            }
        }

        public void MovedAllFiles()
        {

            //string sourcePath = @"C:\Source Folder";
            //string targetPath = @"D:\Destination Folder";
            string sourcePath = ConfigurationManager.AppSettings["Download"].ToString();
            string targetPath = ConfigurationManager.AppSettings["moved"].ToString();
            string date = System.DateTime.Now.ToString("yyyyMMdd");
            Random r = new Random();
            int genRand = r.Next(10, 1000);
            targetPath = targetPath + date + "_" + genRand;
            clsCommon.AttachSuccessFile = targetPath;
            System.IO.Directory.CreateDirectory(targetPath);
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }
            string[] sourcefiles = Directory.GetFiles(sourcePath);
            foreach (string sourcefile in sourcefiles)
            {
                try
                {
                    string fileName = Path.GetFileName(sourcefile);
                    string destFile = Path.Combine(targetPath, fileName);
                    File.Move(sourcefile, destFile);
                }
                catch
                {
                    clsCommon.LogEntry("*******##########*******", "**ERROR:- When file  move from dowload folder into Archive folder in local machine after successfully data impo");
                }
                finally
                {
                }
            }
        }

        public void MovedAllErrorFiles()
        {

            //string sourcePath = @"C:\Source Folder";
            //string targetPath = @"D:\Destination Folder";
            string sourcePath = ConfigurationManager.AppSettings["Download"].ToString();
            string targetPath = ConfigurationManager.AppSettings["movedErrorFile"].ToString();

            string date = System.DateTime.Now.ToString("yyyyMMdd");
            Random r = new Random();
            int genRand = r.Next(10, 1000);
            targetPath = targetPath + date + "_" + genRand;
            clsCommon.AttachErrorFile = targetPath;
            System.IO.Directory.CreateDirectory(targetPath);
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }
            string[] sourcefiles = Directory.GetFiles(sourcePath);
            foreach (string sourcefile in sourcefiles)
            {
                try
                {
                    string fileName = Path.GetFileName(sourcefile);
                    string destFile = Path.Combine(targetPath, fileName);
                    File.Move(sourcefile, destFile);
                }
                catch
                {


                }
                finally
                {


                }

            }

        }
        /// <summary>
        /// Currentlly this metod in not in user it is replace by ImportDataFromClaimTable when in csv format
        /// </summary> 
        public void ImportExcel()
        {

            //OleDbConnection oOledbConn = null;
            //OleDbCommand oCmdSelect = null;
            //OleDbDataAdapter oOledbAdapter = null;
            try
            {
                if (true)
                {

                    // Utilities.ShowWarningMessage("Importing data");
                    string sourcePath = ConfigurationManager.AppSettings["Download"].ToString();
                    string[] files = Directory.GetFiles(sourcePath);
                    // string[] files = Directory.GetFiles(@"D:\Dhananjay Development Code\Code\Testing Development\SFDownloadFile");

                    foreach (string file in files)
                    {
                        string fileName = file;
                        string strFileName = fileName;
                        //string Sql = @"SELECT * FROM [Excel Output$]";
                        //
                        //string oConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + strFileName + "; Extended Properties='Excel 8.0; HDR=Yes'";
                        //oOledbConn = new OleDbConnection(oConnectionString);
                        //if (oOledbConn.State == ConnectionState.Closed) oOledbConn.Open();
                        //oCmdSelect = new OleDbCommand(Sql, oOledbConn);
                        //oOledbAdapter = new OleDbDataAdapter();
                        //oOledbAdapter.SelectCommand = oCmdSelect;
                        //oOledbAdapter.Fill(_dtMTNDetails);


                        using (var con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + strFileName + "; Extended Properties='Excel 8.0; HDR=Yes'"))
                        {
                            con.Open();
                            var dtSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            var strSheetName = "";
                            strSheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();
                            var cmd = new OleDbCommand();
                            var da = new OleDbDataAdapter();
                            cmd.Connection = con;
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "SELECT * FROM [" + strSheetName + "]";
                            da = new OleDbDataAdapter(cmd);
                            da.Fill(_dtMTNDetails);
                        }
                        // Remove last undetscore( _ ) from the Datatable Column
                        for (int i = 0; i < _dtMTNDetails.Columns.Count; i++)
                        {
                            // 
                            _dtMTNDetails.Columns[i].ColumnName = (_dtMTNDetails.Columns[i].ColumnName).ToString().TrimEnd('_');
                            _dtMTNDetails.AcceptChanges();
                        }


                        if (_dtMTNDetails.Rows.Count > 0)
                        {

                            this.DeleteDuplicateRecord();
                            //Thread oThread1 = new Thread(new ThreadStart(DeleteDuplicateRecord)); // CALLING SHOW REPORT FUNCTION 
                            //oThread1.SetApartmentState(ApartmentState.STA);
                            //oThread1.Priority = ThreadPriority.Highest;
                            //oThread1.Start();
                        }
                        if (_dtMTNDetails.Rows.Count > 0)
                        {
                            AddEmployeeData();
                        }


                    }
                }
                else
                {

                }
            }
            catch
            {
                throw;
            }
            finally
            {
                //if (oOledbConn.State == ConnectionState.Open) oOledbConn.Close();
            }
        }

        public void ReadCSVFile()
        {

            #region VARIABLE DECLARATION
            string data1 = "";
            string data2 = "";
            string data3 = "";
            string data4 = "";
            string data5 = "";
            string data6 = "";
            string data7 = "";
            string data8 = "";
            string data9 = "";
            string data10 = "";
            string data11 = "";
            string data12 = "";

            int Count = 0;
            StringBuilder InsertQuery = new StringBuilder();
            StringBuilder InsertQuery2 = new StringBuilder();
            // Utilities.ShowWarningMessage("Importing data");
            string sourcePath = ConfigurationManager.AppSettings["Download"].ToString();
            string[] files = Directory.GetFiles(sourcePath);
            #endregion
            //**********************LOOPING ON FILES IN INWARD FOLDER
            #region LOOPING ON FILES

            foreach (string file in files)
            {
                Count = 0;
                InsertQuery.Length = 0;
                string[] lines = File.ReadAllLines(file);
                string fileName = Path.GetFileName(file);

                int TotalLine = lines.Length;
                try
                {
                    #region LOOPING ON LINES OF EACH COMMA SEPERATED VALUES

                    foreach (string line in lines)
                    {
                        #region Inser Data into temparory table

                        #region DATE ASSIGN IN VARIABLE

                        string[] TotalFields = line.Split(',');
                        data1 = TotalFields[0];
                        data2 = TotalFields[1];
                        data3 = TotalFields[2];
                        data4 = TotalFields[3];
                        data5 = TotalFields[4];
                        data6 = TotalFields[5];
                        data7 = TotalFields[6];
                        data8 = TotalFields[7];
                        data9 = TotalFields[8];
                        data10 = TotalFields[9];
                        data11 = TotalFields[10];
                        data12 = TotalFields[11];

                        #endregion
                        InsertQuery.Append(
                                      @"INSERT INTO [TBL_EmployeeClaim] (
                                                            ClaimCode,ApprovedAmount ,employeecode,StageStartDate,StageCompleteDate    
                                                            ,ExpenseCategoryName,ExpenseHeadName,Employeefullname,[Business Unit]   
                                                           ,[Benefit Name] ,[Claim Date] ,[Total Amount] ,createdate        
                                                         )  VALUES");
                        if (Count != 0)
                        {
                            InsertQuery.Append(String.Format(@"( '{0}','{1}','{2}','{3}','{4}','{5}','{5}','{7}','{8}','{9}','{10}','{11}','{12}')"
                                                           , data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, data12
                                                            , System.DateTime.Now.ToString("yyyyMMdd")));
                            clsCommon.ExecuteNonQuery(InsertQuery.ToString());
                        }
                        InsertQuery.Length = 0;
                        Count = Count + 1;
                        #endregion
                    }
                    #endregion
                }
                catch { throw; }
            }

            #endregion
        }

        public void ImportDataFromClaimTable()
        {
            try
            {
                if (true)
                {
                    _dtMTNDetails = clsCommon.ExecuteDataSet_DataTable("EXEC TVI_PilotCompany.[dbo].[UTL_GETEMPLOYEECLAIMDATA]", clsConnection.gSQLCon_GRIR);
                 
                    if (_dtMTNDetails.Rows.Count > 0)
                    {
                        this.DeleteDuplicateRecord();
                    }

                    if (_dtMTNDetails.Rows.Count > 0)
                    {
                        AddEmployeeData();
                    }
                    if (_dtMTNDetails.Rows.Count == 0)
                    {
                        clsCommon.LogEntry("-------", "Duplicate Data or No Data found");
                    }
                    // if (oOledbConn.State == ConnectionState.Open) oOledbConn.Close();
                    //oOledbConn.Open();

                }
                else
                {

                }
            }
            catch
            {
                throw;
            }
            finally
            {
                //if (oOledbConn.State == ConnectionState.Open) oOledbConn.Close();
            }
        }
        public void DeleteTopTwoBlankRow()
        {
            string sourcePath = ConfigurationManager.AppSettings["Download"].ToString();
            string[] sourcefiles = Directory.GetFiles(sourcePath);


            //initialize a new Workbook object

            Workbook workbook = new Workbook();
            Excel.Application excelApp = new Excel.Application();
            foreach (string sourcefile in sourcefiles)
            {
                //open an excel file
                workbook.LoadFromFile(sourcefile);
                //get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                //delete the first two row of the first sheet
                excelApp.DisplayAlerts = false;
                sheet.DeleteRow(1);
                sheet.DeleteRow(1);
                excelApp.DisplayAlerts = true;
                // sheet.DeleteRow(12);
                //delete the second column of the first sheet
                //  sheet.DeleteColumn(2);
                //save the excel file
                workbook.SaveToFile(sourcefile, ExcelVersion.Version2007);
                //launch the excel file
                //System.Diagnostics.Process.Start(workbook.FileName);
            }
        }

        public void DeleteTop2Rows()
        {
            string sourcePath = ConfigurationManager.AppSettings["Download"].ToString();
            string[] sourcefiles = Directory.GetFiles(sourcePath);
            Excel.Application excelApp = new Excel.Application();
            try
            {
                foreach (string sourcefile in sourcefiles)
                {
                    Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(sourcefile, 1, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, null, false);

                    Excel.Sheets excelWorkSheet = excelWorkbook.Sheets;

                    foreach (Excel.Worksheet work in excelWorkSheet)
                    {
                        Excel.Range range = work.get_Range("A1", "A2");
                        Excel.Range entireRow = range.EntireRow; // update
                        entireRow.Delete(Excel.XlDirection.xlUp);
                        //for (int i = 1; i <= 2; i++)
                        //{
                        //    entireRow.Delete(Excel.XlDirection.xlUp);
                        //}

                    }
                    excelApp.DisplayAlerts = false;
                    // excelWorkbook.Close(false, sourcefile, null);
                    excelWorkbook.Save();
                    excelApp.DisplayAlerts = true;
                    //excelWorkbook.Close(false, mstrFilePath, null);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.DisplayAlerts = false;
                excelApp.Quit();
                excelApp.DisplayAlerts = true;
            }
        }


        /// <summary>
        /// Downloads a remote file through the client into a local directory
        /// </summary>
        /// <param name="client"></param>
        /// <param name="file"></param>
        /// <param name="directory"></param>
        public void DownloadFile(SftpClient client, SftpFile file, string directory)
        {
            //Console.WriteLine("Downloading {0}", file.FullName);

            using (Stream fileStream = File.OpenWrite(Path.Combine(directory, file.Name)))
            {
                client.DownloadFile(file.FullName, fileStream);
            }
        }

        #region Employee Claim Function

        public void AddEmployeeData()
        {
            // Local Variables
            int lineId = 1;
            int docEntryINS = 0;
            string docEntry = String.Empty;
            string docNum = String.Empty;
            string period = String.Empty;
            string series = String.Empty;
            string internal_K = String.Empty;
            string userName = String.Empty;
            string dept = String.Empty;
            StringBuilder sbHeaderInsertQuery = new StringBuilder();
            StringBuilder sbDetailInsertQuery = new StringBuilder();
            //===========================================================
            try
            {
                //this.EnableDisable_Cntrls(false);
                period = clsCommon.getCurrentPeriod();
                series = clsCommon.GetSeries("TVI_OEXP", "");
                clsCommon.GetCurrentUserDetails(out internal_K, out userName, out dept);

                var distinctSuccessFactor = (from row in _dtMTNDetails.AsEnumerable()
                                             select row.Field<string>("employeecode")).Distinct();
                sbHeaderInsertQuery.Append(
                              @"INSERT INTO [@TVI_OEXP] (
                                Canceled,CreateDate,CreateTime,Creator,DataSource,DocEntry,DocNum,Handwrtten,Instance,LogInst
                                ,[Object],Period,Remark,RequestStatus,Series,[Status],Transfered
                                ,U_CardCode,	U_CardName	,U_SFClaimID,	U_DocDate	,U_Status,	U_PrjCode,	U_Indicator	
                                ,U_Location	,U_Remarks,	U_SAPInv,	U_SAPInvDE	,U_DocTotal,	U_LastError
                                ,UpdateDate,UpdateTime,UserSign
                               )
                               VALUES");

                sbDetailInsertQuery.Append(@"INSERT INTO [@TVI_EXP1] (DocEntry,LineId,LogInst,[Object]
                                          ,U_Description	,U_AcctCode,	U_TrvlFrom,	U_TrvlTo,	U_FromDate,	U_ToDate,	U_TaxCode,	U_Price,	U_ActAcctCode,	U_ActPrice,	U_SFCliamID,	U_IntAcctCode
                                          ,VisOrder,U_SFCLAIMID2) VALUES");


                string EnployeeID = "", BenefitClaimID = "", SelectedBenefit = "", ClaimDate = "";
                string Remarks = "", TotalClaimAmount = "";
                string BusinessUnitJobInformation = "";
                int Count = 0;

                #region MyRegion
                string TravelTO = "", Travelfrom = "", FromDate = "", ToDate = "", ClaimAmount = "";
                #endregion
                try
                {
                    foreach (string uniqueSF in distinctSuccessFactor)
                    {
                        if (uniqueSF != null)
                        {
                            DataRow mrfHeader = _dtMTNDetails.Select("[employeecode] = '" + uniqueSF + "'")[0];

                            if (docEntryINS > 0)
                                sbHeaderInsertQuery.Append(",");
                            #region Variable

                            // BenifitName = mrfHeader["Benefit Name"].ToString();
                            EnployeeID = mrfHeader["employeecode"].ToString();
                            if (mrfHeader["employeecode"].ToString() == "")
                            {

                            }

                            BenefitClaimID = mrfHeader["ClaimCode"].ToString().Trim();
                            SelectedBenefit = mrfHeader["Benefit Name"].ToString();
                            ClaimDate = mrfHeader["Claim Date"].ToString();
                            Remarks = mrfHeader["ExpenseHeadName"].ToString().Replace("'", "");
                            if (Remarks.Length > 200) { Remarks = Remarks.Substring(0, 200).Replace("'", ""); }
                            TotalClaimAmount = mrfHeader["Total Amount"].ToString();
                            BusinessUnitJobInformation = mrfHeader["Business Unit"].ToString();
                            EnployeeID = GetEmployeeCode(mrfHeader["employeecode"].ToString());
                            string Circle = "";

                            try
                            {
                                Circle = BusinessUnitJobInformation;//BusinessUnitJobInformation.Substring(BusinessUnitJobInformation.Length - 2, 2);
                                if (Circle == "BR") { Circle = "BH"; }
                                if (Circle == "KA") { Circle = "KT"; }
                                if (Circle == "PE") { Circle = "UP(E)"; }
                                if (Circle == "PW") { Circle = "UP(W)"; }
                                if (Circle == "OL") { Circle = "WB"; }
                            }
                            catch { throw; }
                            #endregion

                            // string company = Utilities.Application.Company.UserName;
                            //--//Canceled,CreateDate,CreateTime,
                            //Creator,DataSource,DocEntry,DocNum,Handwrtten,Instance,LogInst
                            //            ,[Object],Period,Remark,RequestStatus,Series,[Status],Transfered
                            //            ,U_CardCode,	U_CardName	,U_SFClaimID,	U_DocDate	,U_Status,	U_PrjCode,	U_Indicator	
                            //            ,U_Location	,U_Remarks,	U_SAPInv,	U_SAPInvDE	,U_DocTotal,	U_LastError
                            //            ,UpdateDate,UpdateTime,UserSign

                            //*************Index Help*************
                            //0=Creator, 1=DataSource, 2=DocEntry, 2=DocNum, 3=Handwrtten, 4=Instance, 5=LogInst
                            //            ,6=[Object], 7=Period, 8=Remark, 9=RequestStatus, 10= Series,11=[Status],  12=Transfered
                            //            ,13=U_CardCode,	14=U_CardName	,15 =U_SFClaimID,	16=U_DocDate	,17=U_Status,	18=U_PrjCode,	19=U_Indicator	
                            //            ,20=U_Location	, 21=U_Remarks, 22=	U_SAPInv,	23=U_SAPInvDE	, 24=U_DocTotal,	25=U_LastError
                            //            ,27=UpdateDate,  27=UpdateTime,  28=UserSign
                            decimal Doctotal = 0;
                            sbHeaderInsertQuery.Append(String.Format(@"('N',CONVERT(VARCHAR, GETDATE(), 112),REPLACE(LEFT(CAST(GETDATE() AS TIME), 5), ':', ''),
                          '{0}' ,'{1}',DocEntryStart + {2}, DocNumStart + {2},'{3}' ,'{4}','{5}', '{6}' ,'{7}','{8}', '{9}' ,'{10}','{11}', '{12}'
                           ,'{13}','{14}', '{15}' ,'{16}','{17}', '{18}' ,'{19}','{20}', '{21}' ,'{22}','{23}',
                         '{24}' ,'{25}','{26}', '{27}' ,'{28}')"

                           , 497, "I", docEntryINS.ToString(), "N", 0, null, "TVI_OEXP", period, null, "W", null, "O", "N",
                          EnployeeID, GetEmployeeName(EnployeeID), EnployeeID + "#" + Circle + "/" + BenefitClaimID, ClaimDate, "O", Circle, "I4", 20, Remarks.Replace("'", ""), "", "", Doctotal, "", "", "", "497"));

                            lineId = 1;
                            Doctotal = 0;
                            foreach (DataRow mrfRow in _dtMTNDetails.Select("[employeecode] = '" + uniqueSF + "'"))
                            {
                                //clsCommon.LogEntry("---------", "Preparing For Success Factor Row Addition");
                                // BenefitClaimID = mrfHeader["Benefit Employee Claim ID-" + BenifitName].ToString(); //mrfRow["Benefit Employee Claim ID-COV"].ToString();
                                BenefitClaimID = mrfRow["ClaimCode"].ToString().Trim();
                                clsCommon.LogEntry("-------#####--Employee ID:- " + uniqueSF + "--------", "-#####--Claim ID:- " + BenefitClaimID + "");

                                if (lineId > 0 && Count != 0)
                                    sbDetailInsertQuery.Append(",");
                                Count = Count + 1;
                                EnployeeID = GetEmployeeCode(mrfHeader["employeecode"].ToString());
                                try
                                {
                                    Circle = BusinessUnitJobInformation;//BusinessUnitJobInformation.Substring(BusinessUnitJobInformation.Length - 2, 2);
                                    if (Circle == "BR") { Circle = "BH"; }
                                    if (Circle == "KA") { Circle = "KT"; }
                                    if (Circle == "PE") { Circle = "UP(E)"; }
                                    if (Circle == "PW") { Circle = "UP(W)"; }
                                    if (Circle == "OL") { Circle = "WB"; }
                                }
                                catch { throw; }

                                string _accSysCode = "", _accCode = "", _internalAccCode = "", _accDescription = "";
                                SelectedBenefit = mrfRow["Benefit Name"].ToString();
                                //1
                                if (SelectedBenefit == "LC")
                                {
                                    _accSysCode = "_SYS00000000860";  _accCode = "525095045-1-01"; _internalAccCode = "525095045-1-01";
                                    _accDescription = "CONVEYANCE EXPENSES(IN, HO )";
                                    //if (Remarks.Length > 200) { Remarks = Remarks.Substring(0, 200).Replace("'", ""); }
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }
                                //---2
                                else if (SelectedBenefit == "ATC" || SelectedBenefit == "TBC" || SelectedBenefit == "TCC" || SelectedBenefit == "TB" || SelectedBenefit == "TTC")
                                {
                                    _accSysCode = "_SYS00000000881";_accCode = "525105050-1-01";_internalAccCode = "525105050-1-01";
                                    _accDescription = "TRAVELLING EXP. DOMESTIC(IN, HO )";
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }
                                //**3
                                else if (SelectedBenefit == "MOB")
                                {
                                    _accSysCode = "_SYS00000000991"; _accCode = "525125077-1-01"; _internalAccCode = "525125077-1-01";
                                    _accDescription = "REIMBURSEMENT OF MOBILE  (IN, HO )";
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }
                                //**4
                                else if (SelectedBenefit == "LSA")
                                {
                                    _accSysCode = "_SYS00000001513"; _internalAccCode = "515025203-1-01"; _accDescription = "REWARDS & RECOGNITION EXPENSES (IN, HO)";
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }
                                //**5
                                else if (SelectedBenefit == "HBLBR")
                                {
                                    _accSysCode = "_SYS00000000910"; _accCode = "525115057-1-01"; _internalAccCode = "525115057-1-01";
                                    _accDescription = "HOTEL  EXPENSES(IN, HO )";
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }
                                 //**6	TA / DA EXP.(IN, HO )	525105051-1-01
                                else if (SelectedBenefit == "TAD")
                                {
                                    _accSysCode = "_SYS00000000885";_accCode = "525105050-1-01"; _internalAccCode = "525105051-1-01";
                                    _accDescription = "TA / DA EXP.(IN, HO )";
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }
                                // *****7--
                                else if (SelectedBenefit == "RC")
                                {
                                    _accSysCode = "_SYS00000000768";_accCode = "515025024-1-01";_internalAccCode = "525105051-1-01";
                                    _accDescription = "STAFF RELOCATION EXPENSES(IN, HO )";
                                    Travelfrom = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    TravelTO = mrfRow["ExpenseHeadName"].ToString().Trim();
                                    FromDate = mrfRow["StageStartDate"].ToString().Trim();
                                    ToDate = mrfRow["StageCompleteDate"].ToString().Trim();
                                    ClaimAmount = mrfRow["ApprovedAmount"].ToString().Trim();
                                }

                                if (Travelfrom.Length > 65) { Travelfrom = Travelfrom.Substring(0, 65).Replace("'", ""); }
                                if (TravelTO.Length > 65) { TravelTO = TravelTO.Substring(0, 65).Replace("'", ""); }
                                //DocEntry,LineId,LogInst,[Object]
                                // ,U_Description	,U_AcctCode,	U_TrvlFrom,	U_TrvlTo,	U_FromDate,	U_ToDate,	U_TaxCode,	U_Price,	
                                //U_ActAcctCode,	U_ActPrice,	U_SFCliamID,	U_IntAcctCode
                                //                  VisOrder

                                // 0=DocEntry, 2=LineId,LogInst,3=[Object]
                                // ,4=U_Description	,  5=U_AcctCode,	6=U_TrvlFrom,7=	U_TrvlTo,	8=U_FromDate,	9=U_ToDate,10=	U_TaxCode,	11=U_Price,	
                                //12=U_ActAcctCode,	13=U_ActPrice,	14=U_SFCliamID,15=	U_IntAcctCode
                                //  16= VisOrder , 17= U_SFCLAIMID2+1

                                // //ClaimCode	ApprovedAmount	employeecode	StageStartDate	StageCompleteDate	ExpenseCategoryName	
                                //ExpenseHeadName	Employeefullname	ClaimID	Business Unit	Benefit Name	Claim Date	Total Amount

                                string ClaimID = mrfRow["ClaimCode"].ToString().Trim();

                               
                                int len = FromDate.Length;
                                int len1 = ToDate.Length;


                                sbDetailInsertQuery.Append(String.Format(@"(DocEntryStart + {0}, {1}, '{2}', '{3}', '{4}','{5}', '{6}', '{7}', '{8}', '{9}',
                                                   '{10}',
                                                  '{11}', '{12}', '{13}', '{14}', '{15}', '{16}','{17}')"
                                                        , docEntryINS.ToString(), lineId.ToString(), "", "TVI_OEXP", _accDescription, _accCode
                                                        , Travelfrom, TravelTO, FromDate, ToDate
                                                        , "Exempt",
                                                        ClaimAmount, _internalAccCode, ClaimAmount
                                                        , ClaimID, _accSysCode, lineId.ToString(), ClaimID));
                                //  int  TotalClaim= Convert.ToInt32(mrfRow["Total Claim Amount"].ToString());34561
                                // Doctotal = Doctotal + TotalClaim;
                                Travelfrom = "";
                                TravelTO = "";
                                FromDate = "";
                                ToDate = "";
                                ClaimDate = "";
                                lineId++;
                            }

                            docEntryINS++;
                        }
                    }
                }
                catch (Exception ex)
                {

                    string error = ex.Message;
                }
                if (_dtMTNDetails.Rows.Count > 0)
                {
                    clsCommon.LogEntry("-------", "Adding Success Factor");
                }
                if (_dtMTNDetails.Rows.Count == 0)
                {
                    clsCommon.LogEntry("-------", "Duplicate Data or No Data found");
                }


                // Utilities.StartTransaction();
                if (series == null || series == "")
                    series = "0";

                docEntry = clsCommon.ExecuteScalarValue_String("SELECT MAX(CAST(DocEntry AS Numeric)) +1 FROM [@TVI_OEXP]");//clsc.getMaxColumnValue("@TVI_OEXP", "DocEntry");
                docNum = clsCommon.ExecuteScalarValue_String("SELECT MAX(CAST(DocNum AS Numeric))+1 FROM [@TVI_OEXP]");//Utilities.getMaxColumnValue("@TVI_OEXP", "DocNum");

                clsCommon.ExecuteDataSet(sbHeaderInsertQuery.ToString().Replace("DocEntryStart", docEntry).Replace("DocNumStart", docNum));
                clsCommon.ExecuteDataSet(sbDetailInsertQuery.ToString().Replace("DocEntryStart", docEntry));

                clsCommon.ExecuteNonQuery(String.Format(@"UPDATE T0 
                                                                    SET NextNumber = (SELECT ISNULL(MAX(DocNum), 0) FROM [@TVI_OEXP] WHERE Series = {0}) + 1 
                                                                    FROM NNM1 T0 
                                                                    WHERE ObjectCode = 'TVI_OEXP' AND Series = {0}", series));

                clsCommon.ExecuteNonQuery(String.Format(@"UPDATE T0
                                                                    SET AutoKey = (SELECT ISNULL(MAX(DocEntry), 0) FROM [@TVI_OEXP]) + 1
                                                                    FROM ONNM T0
                                                                    WHERE ObjectCode = 'TVI_OEXP' "));

                // Utilities.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                clsCommon.LogEntry("-------------------", "Success Factor is added successfully");
                // ((SAPbouiCOM.StaticText)Form.Items.Item("lbResult").Specific).Caption = String.Concat("Success Factor is added successfully. Document No From ", (Convert.ToInt32(docEntry)).ToString(), " To ", (Convert.ToInt32(docEntry) + (docEntryINS - 1)).ToString());

                string Query = "Select DocEntry from [@TVI_OEXP] where DocEntry>=" + docEntry + " order by DocEntry";
                DataTable _DtDocEntry = clsCommon.ExecuteDataSetDataTable(Query);
                for (int i = 0; i <= _DtDocEntry.Rows.Count - 1; i++)
                {
                    string aa = _DtDocEntry.Rows[i][0].ToString();
                    this.UpdateDocTotal(Convert.ToInt32(_DtDocEntry.Rows[i][0]));

                }
                if (_DtDocEntry.Rows.Count > 0)
                {
                    clsCommon.MailStatus = "T";
                    // clsCommon.SendMailWithAttachment("", "saphelpdesk@tower-vision.com", "!!Success factor Data is imported in SAP!!", "Success factor Data imported in SAP!!<br>  <br><b>NOTE:</b>This is a system generated email, do not reply to this email id.");  
                }


            }
            catch (Exception ex)
            {
                // Utilities.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                // this.EnableDisable_Cntrls(true);
                throw ex;
            }
        }

        private void DeleteDuplicateRecord()
        {
            DataRow dr = null;
            try
            {
                for (int i = _dtMTNDetails.Rows.Count - 1; i >= 0; i--)
                {
                    dr = _dtMTNDetails.Rows[i];
                    string ClaimId = dr["ClaimCode"].ToString().Trim();

                    if (ClaimId == GetSFClaimID(ClaimId))
                        dr.Delete();
                }
                _dtMTNDetails.AcceptChanges();
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                _dtMTNDetails.AcceptChanges();
            }

        }
        private string GetSFClaimID(string SFCID)
        {
            string _sfCid = clsCommon.ExecuteScalarValue_String("Select distinct * from (Select isnull(U_SFCLAIMID2 ,'') SFID from  [@TVI_EXP1] where U_SFCLAIMID2='" + SFCID + "' Union  Select isnull(U_SFCliamID ,'') SFID from  [@TVI_EXP1] where U_SFCliamID='" + SFCID + "') AA where isnull(AA.SFID,'')!='' ");
            if (_sfCid == "")
                _sfCid = "-1";
            return _sfCid;
        }

        private string GetEmployeeName(string EMPID)
        {
            string EMPCODE = "";
            string _empName = clsCommon.ExecuteScalarValue_String("Select CardName from OCRD where CardCode='" + EMPID + "' Union   Select CardName from OCRD where cast(AliasName as varchar)='" + EMPID + "'");
            if (_empName != "-1") { EMPCODE = _empName; }
            else { EMPCODE = EMPID; }
            return EMPCODE;
        }
        private string GetEmployeeCode(string EMPID)
        {
            string EMPCODE = "";
            string _empCode = clsCommon.ExecuteScalarValue_String("Select CardCode from OCRD where CardCode='" + EMPID + "' Union   Select CardCode from OCRD where cast(AliasName as varchar)='" + EMPID + "'");
            if (_empCode != "") { EMPCODE = _empCode; }
            else { EMPCODE = EMPID; }
            return EMPCODE;
        }




        private void UpdateDocTotal(int DocEntry)
        {
            string _empName = clsCommon.ExecuteScalarValue_String(" Update [@TVI_OEXP] Set [@TVI_OEXP].U_DocTotal= (sELECT SUM(TT.u_PRICE) FROM [@TVI_EXP1] TT WHERE TT.DocEntry=[@TVI_OEXP].DocEntry) from [@TVI_OEXP] inner Join [@TVI_EXP1] On [@TVI_OEXP].DocEntry =[@TVI_EXP1].DocEntry  WHERE [@TVI_EXP1].DocEntry=" + DocEntry + "");

        }

        #endregion

        public bool StartExport(DataTable dtbl, bool isFirst, bool isLast, string strOutputPath, string TemplateLocation, string TemplateFullName, int SectionOrder, int totalNoOfSheets)
        {
            bool isSuccess = false;
            try
            {
                if (isFirst)
                {
                    //CopyTemplate(TemplateLocation, strOutputPath, TemplateFullName);
                    xlApp = new Excel.Application();
                    if (xlApp == null)
                    {
                        throw new Exception("Excel is not properly installed!!");
                    }
                    xlWorkBook = xlApp.Workbooks.Open(@strOutputPath + TemplateFullName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                    // To Add Sheets Dynamically
                    for (int i = 0; i <= totalNoOfSheets; i++)
                    {
                        int count = xlWorkBook.Worksheets.Count;
                        Excel.Worksheet addedSheet = xlWorkBook.Worksheets.Add(Type.Missing,
                                xlWorkBook.Worksheets[count], Type.Missing, Type.Missing);
                        addedSheet.Name = "Sheet " + i.ToString();
                    }
                }
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(SectionOrder);
                Startrange = xlWorkSheet.get_Range("A2");
                HeaderStartrange = xlWorkSheet.get_Range("A1");
                FillInExcel(Startrange, HeaderStartrange, dtbl);
                xlWorkSheet.Name = dtbl.TableName;
                if (isLast)
                {
                    xlApp.DisplayAlerts = false;
                    xlWorkBook.SaveAs(@strOutputPath + TemplateFullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
                    xlWorkBook.Close(true, null, null);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    isSuccess = true;
                }
            }
            catch (Exception ex)
            {
                isSuccess = false;
                throw ex;
            }
            return isSuccess;
        }

        public void FillInExcel(Excel.Range startrange, Excel.Range HeaderStartRange, DataTable dtblData)
        {
            int rw = 0;
            int cl = 0;
            try
            {
                // Fill The Report Content Data Here
                rw = dtblData.Rows.Count;
                cl = dtblData.Columns.Count;
                string[,] data = new string[rw, cl];
                // Adding Columns Here
                for (var row = 1; row <= rw; row++)
                {
                    for (var column = 1; column <= cl; column++)
                    {
                        data[row - 1, column - 1] = dtblData.Rows[row - 1][column - 1].ToString();
                    }
                }
                Excel.Range endRange = (Excel.Range)xlWorkSheet.Cells[rw + (startrange.Cells.Row - 1), cl + (startrange.Cells.Column - 1)];
                Excel.Range writeRange = xlWorkSheet.Range[startrange, endRange];
                writeRange.Value2 = data;
                writeRange.Formula = writeRange.Formula;
                data = null;
                startrange = null;
                endRange = null;
                writeRange = null;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

    }
}
