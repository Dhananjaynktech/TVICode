using System;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Xml;
using System.Xml.Xsl;
using System.Threading;
using System.Reflection;
using System.Net;
using System.Net.Mail;
//using Syncfusion.XlsIO;

namespace TVIPLScheduler
{
    /// <summary>
    /// Summary description for Utilities.
    /// Basic utilities are defined in this class
    /// </summary>
    public sealed class Utilities
    {
       // private static EventListener oApplication;
        private static int FormCounter;
        public enum genmPropertyType { SOURCE_PROPERTY, DESTINATION_PROPERTY };
        public enum genmDocType { PO, DRAFT_GRPO };

        private Utilities()
        { }

       

        

        #region Get Application Path
        public static string getApplicationPath()
        {
            string sPath;

            sPath = System.Windows.Forms.Application.StartupPath.Trim();
            //sPath = System.IO.Directory.GetParent(sPath).ToString(); 

            return sPath;
        }
        #endregion

        #region Execute Query
        public static void ExecuteSQL(ref SAPbobsCOM.Recordset RecordSet, string Sql)
        {
            if (RecordSet == null)
                RecordSet = (SAPbobsCOM.Recordset)oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            RecordSet.DoQuery(Sql);
        }
        #endregion

       

     

        #region To Date
        public static DateTime ToDate(string sDate)
        {
            sDate = sDate.Trim().Insert(4, "/").Insert(7, "/");
            return DateTime.Parse(sDate);
            //return Convert.ToDateTime(sDate);  
        }
        #endregion

        #region Fill Combo
        public static void FillCombo(ref SAPbouiCOM.ComboBox oCombo, string sSQL, bool isDefault)
        {
            try
            {
                while (oCombo.ValidValues.Count > 0)
                {
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                SAPbobsCOM.Recordset oRS = null;
                ExecuteSQL(ref oRS, sSQL);

                oRS.MoveFirst();
                while (!oRS.EoF)
                {
                    oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(0).Value.ToString());
                    oRS.MoveNext();
                }

                oRS.MoveFirst();
                oCombo.Select(oRS.Fields.Item(1).Value.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public static void FillCombo_Diff(ref SAPbouiCOM.ComboBox oCombo, string sSQL)
        {
            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

            SAPbobsCOM.Recordset oRS = null;
            ExecuteSQL(ref oRS, sSQL);

            oRS.MoveFirst();
            while (!oRS.EoF)
            {
                oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                oRS.MoveNext();
            }
        }
        #endregion

        #region GridRowCount
        public static void SetRowIndex(SAPbouiCOM.Grid oGrid)
        {
            if (oGrid != null)
            {
                if (oGrid.Rows.Count > 0)
                {
                    int count;
                    count = oGrid.Rows.Count;
                    for (int Row = 1; Row <= count; Row++)
                    {
                        //oGrid.DataTable.Columns.Item(0).Cells.Item(0).Value = "N";

                    }
                }
            }
        }
        #endregion

        #region IsNumeric
        public static bool IsNumeric(object Expression)
        {
            bool isNum;
            double retNum;

            isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);

            return isNum;
        }
        #endregion

        #region Get Max Column Value
        public static string getMaxColumnValue(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRS = null;
            string sSQL, sCode;
            int MaxCode;

            sSQL = "SELECT MAX(CAST(" + Column + " AS Numeric)) FROM [" + Table + "]";
            ExecuteSQL(ref oRS, sSQL);

            if (Convert.ToString(oRS.Fields.Item(0).Value).Length > 0)
                MaxCode = int.Parse(oRS.Fields.Item(0).Value.ToString()) + 1;
            else
                MaxCode = 1;

            sCode = MaxCode.ToString("00000000");

            return sCode;

        }
        #endregion

        #region Transactions
        public static void StartTransaction()
        {
            if (!oApplication.Company.InTransaction)
                oApplication.Company.StartTransaction();
        }

        public static void EndTransaction(SAPbobsCOM.BoWfTransOpt TransEndType)
        {
            if (oApplication.Company.InTransaction)
                oApplication.Company.EndTransaction(TransEndType);
        }
        #endregion

        #region Create Folders
        public static void CreateFolders()
        {
            try
            {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports");

                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Purchase"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Purchase");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\PurchaseReturn"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\PurchaseReturn");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Sales"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Sales");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\SalesReturn"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\SalesReturn");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\InputToProduction"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\InputToProduction");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\OutputFromProduction"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\OutputFromProduction");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Customers"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Customers");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Suppliers"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Suppliers");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Items"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Items");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Show Report

        //		Public Shared Sub ShowReport(ByVal rptName As String, ByVal SourceXML As String)
        //		Constants.REPORT_NAME = rptName
        //		Constants.REPORT_SOURCE_XML = SourceXML
        //
        //		oNewReportThread = New Threading.Thread(AddressOf OpenReport)
        //		oNewReportThread.Priority = Threading.ThreadPriority.Highest
        //		oNewReportThread.Start()
        //
        //		End Sub

        //public static void ReportView(string rptName, string SourceXML, Boolean PrintGo)
        //{
        //    Constants.REPORT_NAME = rptName;
        //    Constants.REPORT_SOURCE_XML = SourceXML;
        //    Constants.PRINT_STATUS = PrintGo;		
        //    Thread myThread = new Thread(new ThreadStart(ReportCall));
        //    myThread.Priority = ThreadPriority.Highest;
        //    myThread.Start();            
        //}

        //public static void ReportCall()
        //{
        //    CrystalDecisions.CrystalReports.Engine.SubreportObject oSubReport;
        //    CrystalDecisions.CrystalReports.Engine.ReportDocument rptSubReportDoc;

        //    ReportViewer rptView = new ReportViewer();
        //    string rptPath = Utilities.getApplicationPath() + @"\Reports\" + Constants.REPORT_NAME;

        //    CrystalDecisions.CrystalReports.Engine.ReportDocument rptDoc = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        //    rptDoc.Load(rptPath);

        //    foreach(CrystalDecisions.CrystalReports.Engine.Table oMainReportTable in rptDoc.Database.Tables)
        //        oMainReportTable.Location = Utilities.getApplicationPath() + @"\XML Files\" + Constants.REPORT_SOURCE_XML;

        //    foreach(CrystalDecisions.CrystalReports.Engine.Section rptSection in rptDoc.ReportDefinition.Sections)
        //    {
        //        foreach(CrystalDecisions.CrystalReports.Engine.ReportObject rptObject in rptSection.ReportObjects)
        //        {
        //            if( rptObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
        //            {
        //                oSubReport = (CrystalDecisions.CrystalReports.Engine.SubreportObject)rptObject;
        //                rptSubReportDoc = oSubReport.OpenSubreport(oSubReport.SubreportName);

        //                foreach(CrystalDecisions.CrystalReports.Engine.Table oSubTable in rptSubReportDoc.Database.Tables)
        //                    oSubTable.Location = Utilities.getApplicationPath() + @"\XML Files\" + Constants.REPORT_SOURCE_XML;
        //            }
        //        }
        //    }

        //    rptView.crystalReportViewer1.ReportSource = rptDoc;
        //    rptView.crystalReportViewer1.RefreshReport();

        //    if (Constants.PRINT_STATUS == false)
        //    {
        //        rptView.ShowDialog();
        //        rptView.Select();
        //        rptView.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        //    }
        //    else
        //    {				
        //        rptView.crystalReportViewer1.PrintReport();
        //        //				rptView.ShowDialog();
        //        //				rptView.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        //    }

        //}


        //		public static void ShowReport(string rptName, string SourceXML, Boolean PrintGo)
        //		{
        //			CrystalDecisions.CrystalReports.Engine.SubreportObject oSubReport;
        //			CrystalDecisions.CrystalReports.Engine.ReportDocument rptSubReportDoc;
        //
        //			ReportViewer rptView = new ReportViewer();
        //			string rptPath = Utilities.getApplicationPath() + @"\Reports\" + rptName;
        //				
        //			CrystalDecisions.CrystalReports.Engine.ReportDocument rptDoc = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        //			rptDoc.Load(rptPath);
        //
        //			foreach(CrystalDecisions.CrystalReports.Engine.Table oMainReportTable in rptDoc.Database.Tables)
        //				oMainReportTable.Location = Utilities.getApplicationPath() + @"\XML Files\" + SourceXML;
        //
        //			foreach(CrystalDecisions.CrystalReports.Engine.Section rptSection in rptDoc.ReportDefinition.Sections)
        //			{
        //				foreach(CrystalDecisions.CrystalReports.Engine.ReportObject rptObject in rptSection.ReportObjects)
        //				{
        //					if( rptObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
        //					{
        //						oSubReport = (CrystalDecisions.CrystalReports.Engine.SubreportObject)rptObject;
        //						rptSubReportDoc = oSubReport.OpenSubreport(oSubReport.SubreportName);
        //
        //						foreach(CrystalDecisions.CrystalReports.Engine.Table oSubTable in rptSubReportDoc.Database.Tables)
        //							oSubTable.Location = Utilities.getApplicationPath() + @"\XML Files\" + SourceXML;
        //					}
        //				}
        //			}
        //
        //			rptView.crystalReportViewer1.ReportSource = rptDoc;
        //			rptView.crystalReportViewer1.RefreshReport();
        //
        //			if (PrintGo == false)
        //			{
        //				rptView.ShowDialog();
        //				rptView.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        //			}
        //			else
        //			{				
        //				rptView.crystalReportViewer1.PrintReport();
        ////				rptView.ShowDialog();
        ////				rptView.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        //			}
        //					
        //		}
        #endregion

        #region SQL Connection String
        public static string SQLConnectionString()
        {
            string ConnStr;
            ConnStr = "user id=" + Constants.USER_ID + ";data source=" + Constants.SERVER + ";pwd=" + Constants.USER_PASSWORD + ";initial catalog=" + Utilities.Application.Company.CompanyDB;

            return ConnStr;
        }
        #endregion

        /// <summary>
        /// Accepts a sql query as a parameter executes it which returns data in data table
        /// </summary>
        /// <param name="Sql"></param>
        /// <returns></returns>
        public static System.Data.DataTable ExecuteDataSet_DataTable(string Sql)
        {
            DataTable dataTable = new DataTable();
            try
            {
                if (Constants.gobjSQLCon.State == ConnectionState.Closed) Constants.gobjSQLCon.Open();
                Constants.gobjSQLCon.ChangeDatabase(Utilities.Application.Company.CompanyDB);
                SqlCommand oCmd = new SqlCommand(Sql, Constants.gobjSQLCon);
                oCmd.CommandTimeout = 1000000000;
                SqlDataAdapter oAd = new SqlDataAdapter(oCmd);
                oAd.Fill(dataTable);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (Constants.gobjSQLCon.State == ConnectionState.Open)
                    Constants.gobjSQLCon.Close();
            }
            return dataTable; // Returning data table 
        }

        public static System.Data.DataSet ExecuteDataSet(string Sql)
        {
            DataSet dataSet = new DataSet();
            try
            {
                if (Constants.gobjSQLCon.State == ConnectionState.Closed) Constants.gobjSQLCon.Open();
                Constants.gobjSQLCon.ChangeDatabase(Utilities.Application.Company.CompanyDB);
                SqlCommand oCmd = new SqlCommand(Sql, Constants.gobjSQLCon);
                oCmd.CommandTimeout = 10000000;
                SqlDataAdapter oAd = new SqlDataAdapter(oCmd);
                oAd.Fill(dataSet);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (Constants.gobjSQLCon.State == ConnectionState.Open)
                    Constants.gobjSQLCon.Close();
            }
            return dataSet; // Returning data table 
        }

        public static string ExecuteScalarSql(string oSQL)
        {
            string scalarvalue = String.Empty;
            Object obj = null;
            try
            {
                if (Constants.gobjSQLCon.State == ConnectionState.Closed) Constants.gobjSQLCon.Open();
                Constants.gobjSQLCon.ChangeDatabase(Utilities.Application.Company.CompanyDB);

                SqlCommand oCmd = new SqlCommand(oSQL, Constants.gobjSQLCon);
                obj = oCmd.ExecuteScalar();

                if (obj != null && !String.IsNullOrEmpty(obj.ToString()))
                    scalarvalue = obj.ToString();
                else
                    scalarvalue="-1";

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (Constants.gobjSQLCon.State == ConnectionState.Open)
                    Constants.gobjSQLCon.Close();
            }
            return scalarvalue;
        }

        public static void ExecuteNonQuery_recordSet(string oSql)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            try
            {
                oRecordSet = (SAPbobsCOM.Recordset)Utilities.oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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

        #region ADD CONTROLS
        public static void AddButton(SAPbouiCOM.Form form, string newBtnID, string oldCntrlID, int distanceFrmLeft, int extraWidth, int formPane, int toPane, string caption)
        {
            SAPbouiCOM.Item oldItem = null, newItem = null;
            try
            {
                oldItem = form.Items.Item(oldCntrlID);
                newItem = form.Items.Add(newBtnID, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                newItem.Top = oldItem.Top;
                newItem.Width = oldItem.Width + extraWidth;
                newItem.Left = oldItem.Left + oldItem.Width + distanceFrmLeft;
                newItem.FromPane = formPane;
                newItem.ToPane = toPane;
                newItem.Visible = true;
                newItem.LinkTo = oldItem.UniqueID;
                ((SAPbouiCOM.Button)newItem.Specific).Caption = caption;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AddStatic(SAPbouiCOM.Form form, string newStaticID, string oldCntrlID, int distanceFrmLeft, int extraWidth, int formPane, int toPane, string caption)
        {
            SAPbouiCOM.Item oldItem = null, newItem = null;
            try
            {
                oldItem = form.Items.Item(oldCntrlID);
                newItem = form.Items.Add(newStaticID, SAPbouiCOM.BoFormItemTypes.it_STATIC);
                newItem.Top = oldItem.Top + oldItem.Height + 1;
                newItem.Width = oldItem.Width + extraWidth;
                newItem.Left = oldItem.Left + distanceFrmLeft;
                newItem.FromPane = formPane;
                newItem.ToPane = toPane;
                newItem.Visible = true;
                newItem.LinkTo = oldItem.UniqueID;
                ((SAPbouiCOM.StaticText)newItem.Specific).Caption = caption;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AddTextBox(SAPbouiCOM.Form form, string newTextID, string oldCntrlID, int distanceFrmLeft, int extraWidth, int formPane, int toPane, string tableName, string columnName)
        {
            SAPbouiCOM.Item oldItem = null, newItem = null;
            try
            {
                oldItem = form.Items.Item(oldCntrlID);
                newItem = form.Items.Add(newTextID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                newItem.Top = oldItem.Top + oldItem.Height + 1;
                newItem.Width = oldItem.Width + extraWidth;
                newItem.Left = oldItem.Left + distanceFrmLeft;
                newItem.FromPane = formPane;
                newItem.ToPane = toPane;
                newItem.Visible = true;
                newItem.LinkTo = oldItem.UniqueID;
                ((SAPbouiCOM.EditText)newItem.Specific).DataBind.SetBound(true, tableName, columnName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AddComboBox(SAPbouiCOM.Form form, string newTextID, string oldCntrlID, int distanceFrmLeft, int extraWidth, int formPane, int toPane, string tableName, string columnName)
        {
            SAPbouiCOM.Item oldItem = null, newItem = null;
            try
            {
                oldItem = form.Items.Item(oldCntrlID);
                newItem = form.Items.Add(newTextID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                newItem.Top = oldItem.Top + oldItem.Height + 1;
                newItem.Width = oldItem.Width + extraWidth;
                newItem.Left = oldItem.Left + distanceFrmLeft;
                newItem.FromPane = formPane;
                newItem.ToPane = toPane;
                newItem.Visible = true;
                newItem.LinkTo = oldItem.UniqueID;
                ((SAPbouiCOM.ComboBox)newItem.Specific).DataBind.SetBound(true, tableName, columnName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AddLinkButton(SAPbouiCOM.Form form, string newTextID, string oldCntrlID, int extraWidth, int formPane, int toPane, SAPbouiCOM.BoLinkedObject linkedObject, string linkedObjectType)
        {
            SAPbouiCOM.Item oldItem = null, newItem = null;
            try
            {
                oldItem = form.Items.Item(oldCntrlID);
                newItem = form.Items.Add(newTextID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                newItem.Top = oldItem.Top;
                newItem.Width = 19;
                newItem.Left = oldItem.Left - 20;
                newItem.FromPane = formPane;
                newItem.ToPane = toPane;
                newItem.Visible = true;
                newItem.LinkTo = oldItem.UniqueID;
                ((SAPbouiCOM.LinkedButton)newItem.Specific).LinkedObject = linkedObject;
                ((SAPbouiCOM.LinkedButton)newItem.Specific).LinkedObjectType = linkedObjectType;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region ADD CHOOSE FROM LIST
        public static void AddChooseFromList(string oFormUID, string oCFL_Text, string oCFL_Button, SAPbouiCOM.BoLinkedObject oObjectType, string oAliasName, string oCondVal
                                            , SAPbouiCOM.BoConditionOperation oOperation, string objType)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.Conditions oCons;
            SAPbouiCOM.Condition oCon;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;

            try
            {
                oCFLs = oApplication.SBO_Application.Forms.Item(oFormUID).ChooseFromLists;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //' Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;

                if (((int)oObjectType) == 0)
                {
                    oCFLCreationParams.ObjectType = objType;
                }
                else
                    oCFLCreationParams.ObjectType = ((int)oObjectType).ToString();

                oCFLCreationParams.UniqueID = oCFL_Text;
                oCFL = oCFLs.Add(oCFLCreationParams);



                //'Adding Conditions to CFL
                oCons = oCFL.GetConditions();
                if (oAliasName != "")
                {
                    oCon = oCons.Add();
                    oCon.Alias = oAliasName;
                    oCon.Operation = oOperation;
                    oCon.CondVal = oCondVal;
                    oCFL.SetConditions(oCons);
                }

                oCFLCreationParams.UniqueID = oCFL_Button;
                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }

        }
        #endregion

        public static void SerializedMartix(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(0).Cells.Item(i).Specific).Value = i.ToString();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Update();
            }
        }

        #region FILL SERIES
        public static void FillSeries(ref SAPbouiCOM.ComboBox seriesCombo, string objectType, string type)
        {
            try
            {
                Utilities.FillCombo_Diff(ref seriesCombo, "SELECT Series, SeriesName FROM NNM1 WHERE Locked <> 'Y' AND ObjectCode = '" + objectType + "' AND (ISNULL(Remark, '') = '" + type + "' OR '" + type + "' = '')");


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region SET DEFAULT SERIES
        public static void SetDfltSeries(string objectType, SAPbouiCOM.DBDataSource headerDataSource, string series, string type)
        {
            DataTable dtDefaultSeries_Prefix = new DataTable();
            try
            {
                dtDefaultSeries_Prefix = ExecuteDataSet_DataTable("EXEC TVIPL_PROC_GetDefaultSeries '" + objectType + "', '" + type + "'");

                if (dtDefaultSeries_Prefix.Rows.Count > 0)
                {
                    if (!String.IsNullOrEmpty(dtDefaultSeries_Prefix.Rows[0]["DfltSeries"].ToString().Trim()) && String.IsNullOrEmpty(series))
                        headerDataSource.SetValue("Series", 0, dtDefaultSeries_Prefix.Rows[0]["DfltSeries"].ToString().Trim());
                }
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
                return ExecuteScalarSql("EXEC TVIPL_PROC_GetDefaultSeries '" + objectType + "', '" + type + "'");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        public static string GetCurrentTime_In24Hrs()
        {
            // Local Variables
            string oCurrentTime = String.Empty;
            string oSql = String.Empty;
            //============================================================
            try
            {
                oSql = "SELECT REPLACE(CONVERT(VARCHAR(5), CAST(GETDATE() AS TIME), 108), ':', '')";
                oCurrentTime = ExecuteScalarSql(oSql);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return oCurrentTime;
        }

        #region TODAY'S DATE
        public static DateTime TodayDate()
        {
            SAPbobsCOM.Recordset oRS = null;
            string oSQL = "select getdate()";
            ExecuteSQL(ref oRS, oSQL);
            if (oRS.RecordCount > 0)
            {
                return Convert.ToDateTime(oRS.Fields.Item(0).Value);
            }
            return DateTime.Today;
        }
        #endregion

        public static void GetCurrentUserDetails(out string InternalK, out string SAPUserName,  out string Dept)
        {
            InternalK = String.Empty;
            SAPUserName = String.Empty;
            Dept = String.Empty;
            DataTable dtUserDetails = new DataTable();
            try
            {
                dtUserDetails = Utilities.ExecuteDataSet_DataTable(@"SELECT T0.Internal_K, T0.U_Name, T1.Name Dept FROM OUSR T0 (NOLOCK) INNER JOIN OUDP T1 (NOLOCK) ON T0.Department = T1.Code  WHERE User_Code = '" + Utilities.Application.Company.UserName + "'");
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

        public static void SendMail(string To, string Subject, string Message, string AttachFile1, string AttachFile2)
        {
            // Local Variable
            int oPort = 0;
            Attachment oAttch = null;
            //===================================================================
            try
            {


                //if (!String.IsNullOrEmpty(clsConnection.gDT_AlertSettings.Rows[0]["U_EmailPort"].ToString().Trim()))
                //    oPort = Convert.ToInt32(clsConnection.gDT_AlertSettings.Rows[0]["U_EmailPort"].ToString().Trim());
                oPort = 25;

                MailMessage oMail = new MailMessage();
                SmtpClient oSmtp = null;

                if (oPort > 0)
                    oSmtp = new SmtpClient("10.0.0.4", 25);
                else
                    oSmtp = new SmtpClient("10.0.0.4");

                //_DcryptEMailPwd = clsConnection.gDT_AlertSettings.Rows[0]["U_EmailPwd"].ToString().Trim(); // Getting EMail password to decrypt
                //oSmtp.Credentials = new System.Net.NetworkCredential("sapb1system@tower-vision.com", clsEncryption.Decrypt(_DcryptEMailPwd));

                //if (clsConnection.gDT_AlertSettings.Rows[0]["U_EnbSSL"].ToString().Trim() == "Y")
                //{
                //oSmtp.EnableSsl = true;
                //ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                //}

                oMail.From = new MailAddress("sapb1system@tower-vision.com");
                oMail.To.Add(new MailAddress("swatig@tower-vision.com"));
                
                //foreach (string ToReceipient in To.Split(','))
                //{
                //    oMail.To.Add(new MailAddress(ToReceipient.Trim()));
                //}

                oMail.Subject = Subject;
                oMail.Body = Message;
                oMail.IsBodyHtml = true;

                if (!String.IsNullOrEmpty(AttachFile1))
                {
                    oAttch = new Attachment(AttachFile1);
                    oMail.Attachments.Add(oAttch);
                }
                if (!String.IsNullOrEmpty(AttachFile2))
                {
                    oAttch = new Attachment(AttachFile2);
                    oMail.Attachments.Add(oAttch);
                }

                oSmtp.Send(oMail);
                if (oMail != null)
                    oMail.Dispose();
            }
            catch (Exception ex)
            {
                //LogEntry(ex.Message, " E-Mail");
                //Application.Exit();
            }

        }

        public static bool validateProperty(string property)
        {
            string sql = String.Empty;
            try
            {
                if (!String.IsNullOrEmpty(ExecuteScalarSql("SELECT Name2 FROM M4U_LL_Properties (NOLOCK) T0 WHERE T0.Name2 = '" + property + "'"))) return true; else return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool isActiveProperty(string property)
        {
            string sql = String.Empty;
            try
            {
                if (!String.IsNullOrEmpty(ExecuteScalarSql("SELECT Name2 FROM M4U_LL_Properties (NOLOCK) T0 WHERE T0.Name2 = '" + property + "' AND T0.Status = '1'"))) return true; else return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool isValidPurpose(string purpose)
        {
            try
            {
                if (!String.IsNullOrEmpty(ExecuteScalarSql("SELECT Code FROM [@TVIPL_MT_PURPOSE] WHERE Code = '" + purpose + "'"))) return true; else return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool isSODetailsMandaory(string purpose)
        {
            try
            {
                if (String.IsNullOrEmpty(ExecuteScalarSql("SELECT Code FROM [@TVIPL_MT_PURPOSE] WHERE Code = '" + purpose + "' AND ISNULL(U_IsSOMan, '') = 'Y' "))) return true; else return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string getCurrentPeriod()
        {
            try
            {
                return ExecuteScalarSql("SELECT TOP 1 AbsEntry FROM OFPR (NOLOCK) WHERE CONVERT(VARCHAR, GETDATE(), 112) BETWEEN F_RefDate AND T_RefDate");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataTable getCurrentUserCircleDetails()
        {
            try
            {
                return Utilities.ExecuteDataSet_DataTable("EXEC TVIPL_PROC_GetSelectedUserBranch " + Utilities.Application.Company.UserSignature.ToString());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataTable getPropertyDetails(string property)
        {
            try
            {
                return ExecuteDataSet_DataTable("SELECT Name2, EntryId, TreePath, WhsCode, PrjCode FROM TVIPL_VW_ListProperty WHERE Name2 = '"+ property +"'  ORDER BY Name2 ASC");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void OpenExcel(string fileName, DataTable dataToExportInExcel)
        {
            fileName = String.Concat(fileName, DateTime.Now.ToString().Replace(":", "").Replace("/", ""), ".xlsx");
            try
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2007;
                    IWorkbook workbook = application.Workbooks.Create(1);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    //Initialize the DataTable
                    //Import DataTable to the worksheet.
                    worksheet.ImportDataTable(dataToExportInExcel, true, 1, 1);
                    workbook.SaveAs(fileName);
                    System.Diagnostics.Process.Start(Utilities.getApplicationPath() + "/" + fileName);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public static string getDraftGRPONo(string poNo)
        {
            try
            {
                return ExecuteScalarSql("SELECT DocEntry FROM ODRF (NOLOCK) T0 WHERE U_BgBsEntr = '" + poNo + "' ");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool RemoveDraft(string DocNo, genmDocType docType)
        {
            int retValue = 0;
            int draftNo = 0;
            int error = 0;
            string draftGRPONo = String.Empty;
            string errorMessage = String.Empty;
            SAPbobsCOM.Documents draft = null;
            try
            {
                if (docType == genmDocType.PO)
                {
                    draftGRPONo = getDraftGRPONo(DocNo);
                }

                if (!String.IsNullOrEmpty(DocNo) && Int32.TryParse(draftGRPONo, out draftNo))
                {
                    draft = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

                    if (draft.GetByKey(draftNo))
                    {
                        retValue = draft.Remove();

                        if (retValue != 0)
                        {
                            Utilities.Application.Company.GetLastError(out error, out errorMessage);
                            Utilities.ShowErrorMessage("Error: " + error.ToString() + " " + errorMessage);
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        public static bool IsSuperUser()
        {
            try
            {
                if (ExecuteScalarSql("SELECT ISNULL(SuperUser, '') FROM OUSR (NOLOCK) WHERE User_Code = '" + Utilities.Application.Company.UserName + "'") == "Y")
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void Get_PO_Circle (String PONo, genmDocType docType, out string PO_Circle)
        {
            try
            {
                if (docType == genmDocType.PO)
                    PO_Circle = ExecuteScalarSql("SELECT T0.HeaderPrjCode FROM M4U_BG_PurchaseDocuments (NOLOCK) T0 WHERE T0.EntryId = '" + PONo + "'");
                else
                    PO_Circle = ExecuteScalarSql("SELECT T0.HeaderPrjCode FROM M4U_BG_PurchaseDocuments (NOLOCK) T0 WHERE T0.EntryId = '" + PONo + "'");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
