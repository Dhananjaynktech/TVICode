using System;
using System.IO;
using System.Data;
using System.Text;
using System.Linq;
using System.Collections;
using System.Collections.Generic;


namespace EYEinvoicingInward
{
    public class clsSendAlert
    {
        #region CLASS LEVEL DECLARATION
        string _errorMsg = String.Empty;
        string _errorMessage = String.Empty;
        DataTable _DataTable = new DataTable();
        DataTable _dtJEDetails = new DataTable();
        StringBuilder sbJEHeader = new StringBuilder();
        StringBuilder sbJELines = new StringBuilder();
        SAPbobsCOM.JournalEntries _journalEntries = null;
        #endregion

        public void UpdateMTNDetails()
        {
            try
            {
                clsCommon.ExecuteNonQuery("EXEC TVIPL_PROC_LoadTaxAmontONFACard");
            }
            catch (Exception ex)
            {
                clsCommon.LogEntry(ex.Message, "UpdateMTNDetails");
            }
        }

        public void UpdateFATaxAmount()
        {
            try
            {
                clsCommon.ExecuteNonQuery("EXEC TVIPL_PROC_PostMRNTransations");
            }
            catch (Exception ex)
            {
                clsCommon.LogEntry(ex.Message, "UpdateFATaxAmount");
            }
        }

        private void JE_HeaderDetails(string uniqueKey)
        {
            try
            {
                DataRow dataRow = _dtJEDetails.Select("UniqueKey = '" + uniqueKey + "'")[0];

                _journalEntries = (SAPbobsCOM.JournalEntries)clsConnection.gCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                sbJEHeader.Append(@"<?xml version=""1.0"" encoding=""utf-8""?>");
                sbJEHeader.Append("<BOM>");
                sbJEHeader.Append("   <BO>");
                sbJEHeader.Append("       <AdmInfo>");
                sbJEHeader.Append("           <Object>30</Object>");
                sbJEHeader.Append("       </AdmInfo>");
                sbJEHeader.Append("       <OJDT>");
                sbJEHeader.Append("           <row>");
                //sbJEHeader.Append("               <DocType>-1</DocType>"); // Setting series
                sbJEHeader.Append("               <RefDate>" + dataRow["DocDate"].ToString() + "</RefDate>"); 
                sbJEHeader.Append("               <DueDate>" + dataRow["DocDate"].ToString() + "</DueDate>");
                sbJEHeader.Append("               <TaxDate>" + dataRow["DocDate"].ToString() + "</TaxDate>");
                sbJEHeader.Append("               <Project>" + dataRow["Circle"].ToString() + "</Project>");
                sbJEHeader.Append("               <Memo>" + dataRow["Remarks"].ToString()+ "</Memo>");
                sbJEHeader.Append("               <Indicator>" + dataRow["Indicator"].ToString() + "</Indicator>");
                sbJEHeader.Append("               <U_OriginNo>" + dataRow["DocEntry"].ToString() + "</U_OriginNo>");
                sbJEHeader.Append("           </row>");
                sbJEHeader.Append("       </OJDT>");
            }
            catch (Exception ex)
            {
                clsCommon.LogEntry(ex.Message, "JE_HeaderDetails");
            }
        }

        public void ProcessGRIREntry()
        {
            int JELine = 0;
            string lastUniqueKey = String.Empty;
            try
            {
                //_dtJEDetails = clsCommon.ExecuteDataSet_DataTable("EXEC TVI_PilotCompany.[dbo].[TVIPL_SAP_GetGRIRDetails]", clsConnection.gSQLCon_GRIR);

                _dtJEDetails = clsCommon.ExecuteDataSet_DataTable("EXEC TVI_PilotCompany.[dbo].[TVIPL_SAP_GetGRIRDetails_BackUp_05may2020]", clsConnection.gSQLCon_GRIR);
              
                if (!clsConnection.gCompany.Connected)
                    clsCommon.connectSAPCompany();

                foreach (DataRow row in _dtJEDetails.Select("1 = 1", "UniqueKey ASC"))
                {
                    if (lastUniqueKey == String.Empty || lastUniqueKey != row["UniqueKey"].ToString().Trim())
                    {
                        if (!String.IsNullOrEmpty(lastUniqueKey))
                        { Add_JE(lastUniqueKey); }

                        sbJEHeader = new StringBuilder();
                        sbJELines = new StringBuilder();

                        JELine = 0;
                        lastUniqueKey = row["UniqueKey"].ToString().Trim();
                        JE_HeaderDetails(lastUniqueKey);
                    }

                    sbJELines.Append("           <row>");
                    sbJELines.Append("               <Line_ID>" + JELine.ToString() + "</Line_ID>");
                    sbJELines.Append("               <Account>" + row["Account"].ToString() + "</Account>");
                    sbJELines.Append("               <Project>" + row["Circle"].ToString() + "</Project>");
                    sbJELines.Append("               <LineMemo>" + row["Remarks"].ToString() + "</LineMemo>");
                    sbJELines.Append("               <U_bgBsType>" + row["ObjType"].ToString() + "</U_bgBsType>");
                    sbJELines.Append("               <U_bgBsEntr>" + row["DocEntry"].ToString() + "</U_bgBsEntr>");
                    sbJELines.Append("               <U_bgBsLine>" + row["LineId"].ToString() + "</U_bgBsLine>");
                    sbJELines.Append("               <U_bgCat1>" + row["BudgetCategory"].ToString() + "</U_bgCat1>");
                    sbJELines.Append("               <U_FromDate>" + row["FromDate"].ToString() + "</U_FromDate>");
                    sbJELines.Append("               <U_ToDate>" + row["ToDate"].ToString() + "</U_ToDate>");

                    if (row["JETranType"].ToString().Trim().ToUpper() == "DEBIT")
                        sbJELines.Append("               <Debit>" + row["Amount"].ToString() + "</Debit>");
                    else
                        sbJELines.Append("               <Credit>" + row["Amount"].ToString() + "</Credit>");

                    sbJELines.Append("           </row>");

                    JELine++;
                }

                if (!String.IsNullOrEmpty(lastUniqueKey))
                { Add_JE(lastUniqueKey); }
            }
            catch (Exception ex)
            {
                clsCommon.LogEntry(ex.Message, "ProcessGRIREntry");
            }
        }

        private void Add_JE(string uniqueKey)
        {
            int oErrorCode = 0;
            string trnsID = String.Empty;
            string _DIError = String.Empty;
            string oPath = clsCommon.getApplicationPath() + "\\GRIR.xml";
            try
            {
                if (!String.IsNullOrEmpty(sbJELines.ToString()))
                {
                    sbJEHeader.Append("       <JDT1>");
                    sbJEHeader.Append(sbJELines);
                    sbJEHeader.Append("       </JDT1>");
                    sbJEHeader.Append("   </BO>");
                    sbJEHeader.Append("</BOM>");

                    System.IO.File.WriteAllText(oPath, sbJEHeader.ToString()); // Writting xml file
                    
                    _journalEntries = (SAPbobsCOM.JournalEntries)clsConnection.gCompany.GetBusinessObjectFromXML(oPath, 0); // Passing xml file having Invoice detils
                    int oRetValue = _journalEntries.Add();

                    if (oRetValue == 0)
                    {
                        string userSign = _dtJEDetails.Select("UniqueKey = '" + uniqueKey + "'")[0]["UserSign"].ToString();
                        trnsID = clsConnection.gCompany.GetNewObjectKey();
                        clsCommon.ExecuteNonQuery_recordSet("UPDATE OJDT SET UserSign = '" + userSign + "', UserSign2 = '" + userSign + "' WHERE TransId = '" + trnsID + "' ");
                        clsCommon.ExecuteNonQuery_recordSet("UPDATE JDT1 SET UserSign = '" + userSign + "' WHERE TransId = '" + trnsID + "' ");
                        clsCommon.ExecuteNonQuery_recordSet("UPDATE OACT SET UserSign2 = '" + userSign + "' WHERE UserSign2 = '540' ");
                    }

                    // In case Journal entry is addition fails, transaction is rolled back.
                    if (oRetValue != 0)
                    {   
                        //oRtrnVal = false;
                        clsConnection.gCompany.GetLastError(out oErrorCode, out _DIError);
                        clsCommon.LogEntry(String.Concat(_DIError, " : ", oErrorCode.ToString()), "ADDING JE");

                        //if (Utilities.Application.Company.InTransaction)
                        //    Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);    // Transaction is rolled back
                        //Utilities.ShowErrorMessage(_DIError);
                    }
                }
            }
            catch (Exception ex)
            {
                clsCommon.LogEntry(ex.Message, "Add_JE");
            }
        }
    }
}
