#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using ReportingTool.BAL;
using System.Data;

using Microsoft.Office.Core;
using Excel1 = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using Excel;
using System.Drawing;

#endregion NameSpace


namespace ReportingTool
{
    public partial class ImportDC : System.Web.UI.Page
    {

        #region Events

        #region Page_Load
        /// <summary>
        /// Page_Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            BindPONumber();
            BindSONumber();
            //BindStore();
            lblMessage.Text = "";
            
        }
        #endregion Page_Load

        #region btnImport_Click
        /// <summary>
        /// btnImport_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnImport_Click(object sender, EventArgs e)
        {

            try
            {
                switch (ddlType.SelectedItem.Value)
                {
                    case "1":
                        if (Validate())
                        {
                            ImportPackExtract();
                        }
                        break;
                    case "2":
                        if (Validate())
                        {
                            ImportContainerExtract();
                        }
                        break;
                    case "3":
                        if (Validate())
                        {
                            ImportPOGrn();
                        }
                        break;
                    case "4":
                        if (Validate())
                        {
                            ImportAllocation();
                            BindStore();
                        }
                        break;

                    case "5":
                        if (Validate())
                        {
                            ImportSOIssueNote();
                        }
                        break;

                    case "6":
                        if (Validate())
                        {
                            ImportProductGroupListing();
                        }
                        break;
                    case "7":
                        if (Validate())
                        {
                            ImportFamilyListing();
                        }
                        break;
                    case "8":
                        if (Validate())
                        {
                            CreateSOFromPO();
                        }
                        break;

                    case "9":
                        if (Validate())
                        {
                            DeleteDocs(txtPONumber.Text.Trim(), "1");
                        }
                        break;

                    case "10":
                        if (Validate())
                        {
                            DeleteDocs(txtSONumber.Text.Trim(), "2");
                        }
                        break;

                    case "11":
                        if (Validate())
                        {
                            DeleteDocs(txtAllocationNo.Text.Trim(), "3");
                        }
                        break;

                    case "12":
                        if (Validate())
                        {
                            ImportExtraContainerExtract();
                        }
                        break;

                    case "13":
                        if (Validate())
                        {
                            ImportAdjustment(); 
                        }
                        break;
                    case "14":
                        if (Validate())
                        {
                            ImportHSCodes();
                        }
                        break;

                }
            }
            catch(Exception ex)
            {
                lblMessage.Text = ex.ToString();
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion btnImport_Click

        #region DwdLog_Click
        /// <summary>
        /// DwdLog_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void DwdLog_Click(object sender, EventArgs e)
        {
            string fileName = ViewState["FileName"].ToString();
            FileDownload(fileName);
        }
        #endregion DwdLog_Click


        #region btnSaveStockLedger_Click
        /// <summary>
        /// btnSaveStockLedger_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnSaveStockLedger_Click(object sender, EventArgs e)
        {
            TatiBAL objDC = new TatiBAL();
            bool Result = false;
            switch (ddlType.SelectedItem.Value)
            {
                case "3":
                    objDC.ReportType = "1";
                    objDC.GrnNo = txtGRNNo.Text.Trim();
                    objDC.IssueNo = "";
                    Result = objDC.InsertStockLedger();
                    break;
                case "5":
                    objDC.ReportType = "2";
                    objDC.IssueNo = txtIssueNoteNo.Text.Trim();
                    objDC.GrnNo = "";
                    Result = objDC.InsertStockLedger();
                    break;

                case "13":
                    objDC.ReportType = "3";
                    objDC.IssueNo =txtAdjustmentNo.Text.Trim();//Passing Adjustment No.
                    objDC.GrnNo = "";
                    Result = objDC.InsertStockLedger();
                    break;
            }
          

            if(Result)
            {
                lblMessage.Text = "Successfully Inserted";
                lblMessage.ForeColor = Color.Green;
                btnSaveStockLedger.Visible = false;
            }
            else
            {
                lblMessage.Text = "Faild";
                lblMessage.ForeColor = Color.Red;
            }

        }
        #endregion btnSaveStockLedger_Click


        #region ddlType_SelectedIndexChanged
        /// <summary>
        /// ddlType_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlType_SelectedIndexChanged(object sender, EventArgs e)
        {

            btnImport.Text = "Import";

            if (ddlType.SelectedItem.Value=="2" || ddlType.SelectedItem.Value == "12")
            {
                txtPONumber.Visible = true;
                CmbPONumber.Visible = false;
            }
            else
            {
                txtPONumber.Visible = false;
                CmbPONumber.Visible = true;
            }


            switch(ddlType.SelectedItem.Value)
            {
                case "1":trPONumber.Visible = false;
                         trGRNNo.Visible = false;
                         trContainerReference.Visible = false;
                         trAllocationNo.Visible = false;

                         trStoreNo.Visible = false;
                         trSONumber.Visible = false;
                         trIssueNoteNo.Visible = false;
                         trtxtSONumber.Visible = false;

                         trAsOfDate.Visible = false;
                         trCompany.Visible = false;
                         fileuploadExcel.Visible = true;

                         trAdjustment.Visible = false;
                         trSelectAdjustment.Visible = false;
                         trDocNo.Visible = false;
                         trAdjustmentDate.Visible = false;

                         break;
                case "2":trPONumber.Visible = true;
                         trGRNNo.Visible = false;
                         trContainerReference.Visible = true;
                         trAllocationNo.Visible = false;

                         trStoreNo.Visible = false;
                         trSONumber.Visible = false;
                         trIssueNoteNo.Visible = false;
                         trtxtSONumber.Visible = false;

                         trAsOfDate.Visible = false;
                         trCompany.Visible = false;
                         fileuploadExcel.Visible = true;

                         trAdjustment.Visible = false;
                         trSelectAdjustment.Visible = false;
                         trDocNo.Visible = false;
                         trAdjustmentDate.Visible = false;
                         break;

                case "3":trPONumber.Visible = true;
                         trGRNNo.Visible = true;
                         trContainerReference.Visible = false;
                         trAllocationNo.Visible = false;

                         trStoreNo.Visible = false;
                         trSONumber.Visible = false;
                         trIssueNoteNo.Visible = false;
                         trtxtSONumber.Visible = false;

                         trAsOfDate.Visible = false;
                         trCompany.Visible = false;
                         fileuploadExcel.Visible = true;

                         trAdjustment.Visible = false;
                         trSelectAdjustment.Visible = false;
                         trDocNo.Visible = false;
                         trAdjustmentDate.Visible = false;
                    break;
                case "4":trPONumber.Visible = false;
                         trGRNNo.Visible = false;
                         trContainerReference.Visible = false;
                         trAllocationNo.Visible = true;
                         BindStore();
                         trStoreNo.Visible = true;
                         trSONumber.Visible = false;
                         trIssueNoteNo.Visible = false;
                         trtxtSONumber.Visible = false;

                         trAsOfDate.Visible = false;
                         trCompany.Visible = false;
                         fileuploadExcel.Visible = true;

                        trAdjustment.Visible = false;
                        trSelectAdjustment.Visible = false;
                        trDocNo.Visible = false;
                        trAdjustmentDate.Visible = false;

                    break;
                case "5":
                        trPONumber.Visible = false;
                        trGRNNo.Visible = false;
                        trContainerReference.Visible = false;
                        trAllocationNo.Visible = false;

                        trStoreNo.Visible = false;
                        trSONumber.Visible = true;
                        trIssueNoteNo.Visible = true;
                        trtxtSONumber.Visible = false;

                        trAsOfDate.Visible = false;
                        trCompany.Visible = false;
                        fileuploadExcel.Visible = true;

                        trAdjustment.Visible = false;
                        trSelectAdjustment.Visible = false;
                        trDocNo.Visible = false;
                        trAdjustmentDate.Visible = false;
                     break;
                    case "6":
                        trPONumber.Visible = false;
                        trGRNNo.Visible = false;
                        trContainerReference.Visible = false;
                        trAllocationNo.Visible = false;
                        trStoreNo.Visible = true;
                        trSONumber.Visible = false;
                        trIssueNoteNo.Visible = false;
                        trtxtSONumber.Visible = false;
                        BindStore();
                        trAsOfDate.Visible = false;
                        trCompany.Visible = false;
                        fileuploadExcel.Visible = true;

                        trAdjustment.Visible = false;
                        trSelectAdjustment.Visible = false;
                        trDocNo.Visible = false;
                        trAdjustmentDate.Visible = false;
                    break;
                    case "7":
                        trPONumber.Visible = false;
                        trGRNNo.Visible = false;
                        trContainerReference.Visible = false;
                        trAllocationNo.Visible = false;
                        trStoreNo.Visible = true;
                        trSONumber.Visible = false;
                        trIssueNoteNo.Visible = false;
                        trtxtSONumber.Visible = false;
                        BindStore();
                        trAsOfDate.Visible = false;
                        trCompany.Visible = false;
                        fileuploadExcel.Visible = true;

                        trAdjustment.Visible = false;
                        trSelectAdjustment.Visible = false;
                        trDocNo.Visible = false;
                        trAdjustmentDate.Visible = false;

                        break;
                    case "8":
                        trPONumber.Visible = false;
                        trGRNNo.Visible = true;
                        trContainerReference.Visible = false;
                        trAllocationNo.Visible = false;

                        trStoreNo.Visible = true;
                        trSONumber.Visible = false;
                        trIssueNoteNo.Visible = false;
                        trtxtSONumber.Visible = true;

                        trAsOfDate.Visible = true;
                        trCompany.Visible = true;
                        fileuploadExcel.Visible = false;
                        btnImport.Text = "Post";

                        BindStore();
                        trAdjustment.Visible = false;
                        trSelectAdjustment.Visible = false;
                        trDocNo.Visible = false;

                        trAdjustmentDate.Visible = false;
                    break;

                        case "9":

                        trPONumber.Visible = true;
                        txtPONumber.Visible = true;
                        CmbPONumber.Visible = false;
                        trGRNNo.Visible = false;

                        trContainerReference.Visible = false;
                        trAllocationNo.Visible = false;
                        trStoreNo.Visible = false;
                        trSONumber.Visible = false;

                        trIssueNoteNo.Visible = false;
                        trtxtSONumber.Visible = false;
                        trAsOfDate.Visible = false;
                        trCompany.Visible = false;

                        fileuploadExcel.Visible = false;
                        btnImport.Text = "Delete PO";
                        trAdjustment.Visible = false;

                        trSelectAdjustment.Visible = false;
                        trDocNo.Visible = false;
                        trAdjustmentDate.Visible = false;
                    break;

                case "10":

                    trPONumber.Visible = false;
                    txtPONumber.Visible = false;
                    CmbPONumber.Visible = false;
                    trGRNNo.Visible = false;

                    trContainerReference.Visible = false;
                    trAllocationNo.Visible = false;
                    trStoreNo.Visible = false;
                    trSONumber.Visible = false;

                    trIssueNoteNo.Visible = false;
                    trtxtSONumber.Visible = true;
                    trAsOfDate.Visible = false;
                    trCompany.Visible = false;

                    fileuploadExcel.Visible = false;
                    btnImport.Text = "Delete SO";

                    trAdjustment.Visible = false;
                    trSelectAdjustment.Visible = false;
                    trDocNo.Visible = false;
                    trAdjustmentDate.Visible = false;
                    break;

                case "11":

                    trPONumber.Visible = false;
                    txtPONumber.Visible = false;
                    CmbPONumber.Visible = false;
                    trGRNNo.Visible = false;

                    trContainerReference.Visible = false;
                    trAllocationNo.Visible = true;
                    trStoreNo.Visible = false;
                    trSONumber.Visible = false;

                    trIssueNoteNo.Visible = false;
                    trtxtSONumber.Visible = false;
                    trAsOfDate.Visible = false;
                    trCompany.Visible = false;

                    fileuploadExcel.Visible = false;
                    btnImport.Text = "Delete Allo.";

                    trAdjustment.Visible = false;
                    trSelectAdjustment.Visible = false;
                    trDocNo.Visible = false;
                    trAdjustmentDate.Visible = false;
                    break;

                case "12":
                    trPONumber.Visible = true;
                    trGRNNo.Visible = false;
                    trContainerReference.Visible = true;
                    trAllocationNo.Visible = false;

                    trStoreNo.Visible = false;
                    trSONumber.Visible = false;
                    trIssueNoteNo.Visible = false;
                    trtxtSONumber.Visible = false;

                    trAsOfDate.Visible = false;
                    trCompany.Visible = false;
                    fileuploadExcel.Visible = true;

                    trAdjustment.Visible = false;
                    trSelectAdjustment.Visible = false;
                    trDocNo.Visible = false;
                    trAdjustmentDate.Visible = false;
                    break;

                case "13":

                    trPONumber.Visible = false;
                    trGRNNo.Visible = false;
                    trContainerReference.Visible = false;
                    trAllocationNo.Visible = false;

                    trStoreNo.Visible = false;
                    trSONumber.Visible = false;
                    trIssueNoteNo.Visible = false;
                    trtxtSONumber.Visible = false;

                    trAsOfDate.Visible = false;
                    trCompany.Visible = false;
                    fileuploadExcel.Visible = true;
                    trAdjustment.Visible = true;

                    trSelectAdjustment.Visible = true;
                    trDocNo.Visible = true;
                    trAdjustmentDate.Visible = true;
                    BindDoc("9", "PONumber");
                    break;

                case "14":
                    trPONumber.Visible = false;
                    trGRNNo.Visible = false;
                    trContainerReference.Visible = false;
                    trAllocationNo.Visible = false;

                    trStoreNo.Visible = false;
                    trSONumber.Visible = false;
                    trIssueNoteNo.Visible = false;
                    trtxtSONumber.Visible = false;

                    trAsOfDate.Visible = false;
                    trCompany.Visible = false;
                    fileuploadExcel.Visible = true;
                    trAdjustment.Visible = false;

                    trSelectAdjustment.Visible = false;
                    trDocNo.Visible = false;
                    trAdjustmentDate.Visible = false;
                    break;
            }

        }
        #endregion ddlType_SelectedIndexChanged

        #region rdlSelectDoc_SelectedIndexChanged
        /// <summary>
        /// rdlSelectDoc_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void rdlSelectDoc_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(rdlSelectDoc.SelectedItem.Value)
            {
                case "1": BindDoc("9", "PONumber");
                          trDocNo.Visible = true;
                          break;

                case "2":BindDoc("10", "SONumber");
                         trDocNo.Visible = true;
                         break;

                case "3":
                         trDocNo.Visible = false;
                         break;
            }

        }
        #endregion rdlSelectDoc_SelectedIndexChanged

        #endregion Events

        #region Methods

        #region BindDoc
        /// <summary>
        /// BindDoc
        /// </summary>
        private void BindDoc(string ReportType, string FieldName)
        {
            TatiBAL objDC = new TatiBAL();
            objDC.DocNo = "";
            objDC.ReportType = ReportType;
            DataTable dt = objDC.GetAllPOHeader();
            cmbDocNumber.DataSource = dt;
            cmbDocNumber.DataMember = FieldName;
            cmbDocNumber.DataValueField = FieldName;

            cmbDocNumber.DataBind();
            cmbDocNumber.Items.Insert(0, "<-- Select -->");
        }
        #endregion BindDoc


        #region ImportPackExtract
        /// <summary>
        /// ImportPackExtract
        /// </summary>
        private void ImportPackExtract()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "Pack_Data" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);
            if(!dtImport.Columns.Contains("Linked Pack ID"))
            {
                dtImport.Columns.Add("Linked Pack ID");
            }

            dtImport.Columns.Add("Pack Id1");
            dtImport.Columns.Add("Linked Pack ID1");

            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                if (dtImport.Rows[i]["Pack Id"].ToString().Length == 1)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack Id"]);
                    str = "0000" + str;
                    dtImport.Rows[i]["Pack Id1"] = str;
                }
               else if (dtImport.Rows[i]["Pack Id"].ToString().Length == 2)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack Id"]);
                    str = "000" + str;
                    dtImport.Rows[i]["Pack Id1"] = str;
                }
               else if (dtImport.Rows[i]["Pack Id"].ToString().Length == 3)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack Id"]);
                    str = "00" + str;
                    dtImport.Rows[i]["Pack Id1"] = str;
                }
                else if (dtImport.Rows[i]["Pack Id"].ToString().Length == 4)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack Id"]);
                    str = "0" + str;
                    dtImport.Rows[i]["Pack Id1"] = str;
                }
                else
                {
                    dtImport.Rows[i]["Pack Id1"] = Convert.ToString(dtImport.Rows[i]["Pack Id"]);
                }


                if (dtImport.Rows[i]["Linked Pack ID"].ToString().Length == 1)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Linked Pack ID"]);
                    str = "0000" + str;
                    dtImport.Rows[i]["Linked Pack ID1"] = str;
                }
                else if (dtImport.Rows[i]["Linked Pack ID"].ToString().Length == 2)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Linked Pack ID"]);
                    str = "000" + str;
                    dtImport.Rows[i]["Linked Pack ID1"] = str;
                }
                else if (dtImport.Rows[i]["Linked Pack ID"].ToString().Length == 3)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Linked Pack ID"]);
                    str = "00" + str;
                    dtImport.Rows[i]["Linked Pack ID1"] = str;
                }
                else if (dtImport.Rows[i]["Linked Pack ID"].ToString().Length == 4)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Linked Pack ID"]);
                    str = "0" + str;
                    dtImport.Rows[i]["Linked Pack ID1"] = str;
                }
                else
                {
                    dtImport.Rows[i]["Linked Pack ID1"] = Convert.ToString(dtImport.Rows[i]["Linked Pack ID"]);
                }
            }


            TatiBAL ObjImport = new TatiBAL();
            //ObjImport.DtSource = dtImport;
            Boolean Result = false;
            int Count = 0;

            DataTable dt = new DataTable();

            dt.Columns.Add("LineNo");
            dt.Columns.Add("PackBarcode");
            dt.Columns.Add("LineCode12");
            dt.Columns.Add("Remarks");

            for (int i=0;i<dtImport.Rows.Count;i++)
            {

                ObjImport.PackBarcode = dtImport.Rows[i]["Pack Barcode"].ToString();
                ObjImport.LineCode = dtImport.Rows[i]["Linecode(7)"].ToString();
                ObjImport.PackId = dtImport.Rows[i]["Pack Id1"].ToString();
                ObjImport.PackType = dtImport.Rows[i]["Pack Type"].ToString();

                ObjImport.PackOuter = Convert.ToDecimal(dtImport.Rows[i]["Pack Outer"].ToString());
                ObjImport.LineCode12 = dtImport.Rows[i]["Linecode(12)"].ToString();
                ObjImport.Ratio = Convert.ToDecimal(dtImport.Rows[i]["Ratio"].ToString());
                ObjImport.AllSizesInPack = dtImport.Rows[i]["All Sizes in Pack?"].ToString();

                ObjImport.PackLevel = dtImport.Rows[i]["Pack Level"].ToString();
                ObjImport.LinkedPackId = dtImport.Rows[i]["Linked Pack ID1"].ToString();

                Result = ObjImport.InsertPackExtract();

                if(Result)
                {
                    Count++;
                    dt.Rows.Add();

                    dt.Rows[i]["LineNo"] = i.ToString();
                    dt.Rows[i]["PackBarcode"] = dtImport.Rows[i]["Pack Barcode"].ToString();
                    dt.Rows[i]["LineCode12"] = dtImport.Rows[i]["Linecode(12)"].ToString();
                    dt.Rows[i]["Remarks"] = "Success";
                }
                else
                {
                    dt.Rows.Add();
                    dt.Rows[i]["LineNo"] = i.ToString();
                    dt.Rows[i]["PackBarcode"] = dtImport.Rows[i]["Pack Barcode"].ToString();
                    dt.Rows[i]["LineCode12"] = dtImport.Rows[i]["Linecode(12)"].ToString();
                    dt.Rows[i]["Remarks"] = "Failed";
                }
            }
            
            if (Result)
            {
                String filePath = Server.MapPath(".") + "\\Reports\\PackExtract_Log" + "_" + rnd.Next() + ".csv";

                if (dt != null && dt.Rows.Count > 0)
                {
                    ViewState["FileName"] = filePath;
                    StreamWriter sw = new StreamWriter(@filePath, false);

                    ExportToCsv(dt, sw);
                    sw.Close();
                    btnDwdLog.Visible = true;
                }

                lblMessage.Text = Count.ToString()+" Rows Imported!";
                lblMessage.ForeColor = Color.Green;
            }
            else
            {
                String filePath = Server.MapPath(".") + "\\Reports\\PackExtract_Log" + "_" + rnd.Next() + ".csv";

                if (dt != null && dt.Rows.Count > 0)
                {
                    ViewState["FileName"] = filePath;
                    StreamWriter sw = new StreamWriter(@filePath, false);

                    ExportToCsv(dt, sw);
                    sw.Close();
                    btnDwdLog.Visible = true;
                }

                lblMessage.Text = Count.ToString() + " Rows Imported!"; ;
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion ImportPackExtract


        #region ImportContainerExtract
        /// <summary>
        /// ImportContainerExtract
        /// </summary>
        private void ImportContainerExtract()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "Container_Extract" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            dtImport.Columns.Add("Pack ID1");

            DataTable dtCheck = new DataTable();

            dtCheck.Columns.Add("LineCode7");
            dtCheck.Columns.Add("PackId");
            dtCheck.Columns.Add("PackType");
            dtCheck.Columns.Add("Remarks");

            TatiBAL ObjImport = new TatiBAL();
            DataTable dt = new DataTable();
            int j = 0;

            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                if (dtImport.Rows[i]["Pack ID"].ToString().Length == 1)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack ID"]);
                    str = "0000" + str;
                    dtImport.Rows[i]["Pack ID1"] = str;
                }
                else if (dtImport.Rows[i]["Pack ID"].ToString().Length == 2)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack ID"]);
                    str = "000" + str;
                    dtImport.Rows[i]["Pack ID1"] = str;
                }
                else if (dtImport.Rows[i]["Pack ID"].ToString().Length == 3)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack ID"]);
                    str = "00" + str;
                    dtImport.Rows[i]["Pack ID1"] = str;
                }
                else if (dtImport.Rows[i]["Pack ID"].ToString().Length == 4)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["Pack ID"]);
                    str = "0" + str;
                    dtImport.Rows[i]["Pack ID1"] = str;
                }

                else
                {
                    dtImport.Rows[i]["Pack ID1"] = Convert.ToString(dtImport.Rows[i]["Pack ID"]);
                }

                ObjImport.LineCode = dtImport.Rows[i]["Linecode"].ToString();
                ObjImport.PackId = dtImport.Rows[i]["Pack ID1"].ToString();
                ObjImport.PackType = dtImport.Rows[i]["Pack Type"].ToString();

                dt =ObjImport.CheckContainerExtract();

                if(dt.Rows.Count==0)
                {
                    dtCheck.Rows.Add();
                    dtCheck.Rows[j]["LineCode7"] = dtImport.Rows[i]["Linecode"].ToString();
                    dtCheck.Rows[j]["PackId"] = dtImport.Rows[i]["Pack ID1"].ToString();
                    dtCheck.Rows[j]["PackType"] = dtImport.Rows[i]["Pack Type"].ToString();
                    dtCheck.Rows[j]["Remarks"] = "PackBarcode Not Available";

                    j++;
                }

            }

            if (dtCheck.Rows.Count > 0)
            {
                String filePath = Server.MapPath(".") + "\\Reports\\GRN_Log" + txtGRNNo.Text.Trim() + "_" + rnd.Next() + ".csv";
                ViewState["FileName"] = filePath;
                StreamWriter sw = new StreamWriter(@filePath, false);

                ExportToCsv(dtCheck, sw);
                sw.Close();

                btnDwdLog.Visible = true;
                lblMessage.Text = "Log File Generated";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;

            }

            else
            {
                ObjImport.ContainerNo = txtContainerReference.Text.Trim();
                ObjImport.DtSource = dtImport;
                Boolean Result = ObjImport.ImportContainerExtract();

                ObjImport.PONumber = txtPONumber.Text.Trim();

                if (Result)
                {
                    Result = ObjImport.InsertPODetail();
                }

                if (Result)
                {
                    lblMessage.Text = "Successfully Imported!";
                    lblMessage.ForeColor = Color.Green;
                }
                else
                {
                    lblMessage.Text = "Faild!";
                    lblMessage.ForeColor = Color.Red;
                }
            }
        }
        #endregion ImportContainerExtract


        #region ImportExtraContainerExtract
        /// <summary>
        /// ImportExtraContainerExtract
        /// </summary>
        private void ImportExtraContainerExtract()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "Extra_Container_Extract" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            if (dtImport.Rows.Count > 0)
            {

                if (dtImport.Rows[0]["ContainerNo"].ToString() == txtContainerReference.Text.Trim())
                {

                    TatiBAL ObjImport = new TatiBAL();

                    ObjImport.DtSource = dtImport;
                    Boolean Result = ObjImport.ImportExtraContainerExtract();

                    ObjImport.ContainerNo = txtContainerReference.Text.Trim();
                    ObjImport.PONumber = txtPONumber.Text.Trim();

                    if (Result)
                    {
                        Result = ObjImport.InsertExtraContainerExtract();
                    }

                    if (Result)
                    {
                        Result = ObjImport.InsertPODetail();
                    }

                    if (Result)
                    {
                        lblMessage.Text = "Successfully Imported!";
                        lblMessage.ForeColor = Color.Green;
                    }
                    else
                    {
                        lblMessage.Text = "Faild!";
                        lblMessage.ForeColor = Color.Red;
                    }
                }
                else
                {

                    lblMessage.Text = "Put Proper Conatainer No.";
                    lblMessage.ForeColor = Color.Red;

                }
            }
          
        }
        #endregion ImportExtraContainerExtract


        #region ImportPOGrn
        /// <summary>
        /// ImportPOGrn
        /// </summary>
        private void ImportPOGrn()
        {
            TatiBAL ObjImport = new TatiBAL();
            try
            {
                String path = Server.MapPath("~/FileImport/");
                Random rnd = new Random();
                String fileName = "PO_GRN" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
                fileuploadExcel.PostedFile.SaveAs(path
                    + fileName);

                //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

                DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

                dtImport.Columns.Add("PackID1");
                bool Flag = false;

                for (int i = 0; i < dtImport.Rows.Count; i++)
                {
                    if (dtImport.Rows[i]["PackID"].ToString().Length == 1)
                    {

                        string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                        str = "0000" + str;
                        dtImport.Rows[i]["PackID1"] = str;
                    }
                   else if (dtImport.Rows[i]["PackID"].ToString().Length == 2)
                    {

                        string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                        str = "000" + str;
                        dtImport.Rows[i]["PackID1"] = str;
                    }
                    else if (dtImport.Rows[i]["PackID"].ToString().Length == 3)
                    {

                        string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                        str = "00" + str;
                        dtImport.Rows[i]["PackID1"] = str;
                    }
                    else
                    {
                        dtImport.Rows[i]["PackID1"] = Convert.ToString(dtImport.Rows[i]["PackID"]);
                    }

                    if (dtImport.Rows[i]["PONUMBER"].ToString() != CmbPONumber.Text.Trim() )
                    {
                        Flag = true;
                        break;
                    }

                    if (dtImport.Rows[i]["PAGRNNO"].ToString() != txtGRNNo.Text.Trim())
                    {
                        Flag = true;
                        break;
                    }

                }

                if (!Flag)
                {

                    
                    ObjImport.GrnNo = txtGRNNo.Text.Trim();
                    ObjImport.FileName = fileName;
                    ObjImport.PONumber = CmbPONumber.Text.Trim();
                    //ObjImport.DtSource = dtImport;

                    Boolean Result = ObjImport.InsertPOGRNHeader();
                    int Count = 0;

                    if (Result)
                    {
                        //Result = ObjImport.ImportPOGRN();

                        for (int i = 0; i < dtImport.Rows.Count; i++)
                        {

                            // PAGRNNO PAGRNDATE   PAGRNLINENO PONUMBER    CONTAINERNO POLINENO    LINECODE7 PACKID  PACKBARCODE PACKTYPE    PAGRNQTY
                            if (dtImport.Rows[i]["PONUMBER"].ToString() == CmbPONumber.Text && dtImport.Rows[i]["PAGRNNO"].ToString() == txtGRNNo.Text.Trim())
                            {
                                ObjImport.GrnNo = dtImport.Rows[i]["PAGRNNO"].ToString();
                                ObjImport.GrnDate = DateTime.ParseExact(dtImport.Rows[i]["PAGRNDATE"].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                ObjImport.GrnLineNo = Convert.ToInt32(dtImport.Rows[i]["PAGRNLINENO"].ToString());
                                ObjImport.PONumber = dtImport.Rows[i]["PONUMBER"].ToString();

                                ObjImport.ContainerNo = dtImport.Rows[i]["CONTAINERNO"].ToString();
                                ObjImport.POLineNo = Convert.ToInt32(dtImport.Rows[i]["POLINENO"].ToString());
                                ObjImport.LineCode = dtImport.Rows[i]["LINECODE7"].ToString();
                                ObjImport.PackId = dtImport.Rows[i]["PackID1"].ToString();

                                ObjImport.PackBarcode = dtImport.Rows[i]["PACKBARCODE"].ToString();
                                ObjImport.PackType = dtImport.Rows[i]["PACKTYPE"].ToString();
                                ObjImport.GrnQty = Convert.ToDecimal(dtImport.Rows[i]["PAGRNQTY"].ToString());

                                Result = ObjImport.InsertPOGRN();

                                if (Result)
                                {
                                    Count++;
                                }
                            }
                            else
                            {
                                lblMessage.Text = "LineNo. " + i.ToString() + " PO Number or Grn No. Not Proper";
                                lblMessage.ForeColor = Color.Red;
                            }
                        }


                    }
                    else
                    {
                        lblMessage.Text = "This GRN No. Already Existing !";
                        lblMessage.ForeColor = Color.Red;
                        btnSaveStockLedger.Visible = false;
                    }

                    if (Result)
                    {
                        String filePath = Server.MapPath(".") + "\\Reports\\GRN_Log" + txtGRNNo.Text.Trim() + "_" + rnd.Next() + ".csv";

                        DataTable dt = ObjImport.GetGRNLog();
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            ViewState["FileName"] = filePath;
                            StreamWriter sw = new StreamWriter(@filePath, false);

                            ExportToCsv(dt, sw);
                            sw.Close();

                            btnDwdLog.Visible = true;
                            lblMessage.Text = "Log File Generated";
                            lblMessage.ForeColor = Color.Red;
                            btnSaveStockLedger.Visible = false;
                        }
                        else
                        {
                            lblMessage.Text = Count.ToString() + " Rows Successfully Imported!";
                            lblMessage.ForeColor = Color.Green;
                            btnDwdLog.Visible = false;
                            btnSaveStockLedger.Visible = true;
                        }
                    }
                    else
                    {
                        lblMessage.Text = "Failed,Try Again";
                        lblMessage.ForeColor = Color.Red;
                        btnSaveStockLedger.Visible = false;
                        bool Res=ObjImport.DeletePOGrnHeader();
                    }
                }
                else
                {
                    lblMessage.Text = "PO Number or Grn No. Not Proper";
                    lblMessage.ForeColor = Color.Red;
                    btnSaveStockLedger.Visible = false;
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = ex.ToString()+ "----------TRY TO IMPORT AGAIN----------";
                lblMessage.ForeColor = Color.Red;
                bool Res = ObjImport.DeletePOGrnHeader();
            }
        }
        #endregion ImportPOGrn


        #region ImportAllocation
        /// <summary>
        /// ImportAllocation
        /// </summary>
        private void ImportAllocation()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "Allocation" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            dtImport.Columns.Add("PackId1");
            dtImport.Columns.Add("AllocationNo");
            dtImport.Columns.Add("StoreNo");

            for (int i = 0; i < dtImport.Rows.Count; i++)
            {

                dtImport.Rows[i]["AllocationNo"] =txtAllocationNo.Text.Trim();
                dtImport.Rows[i]["StoreNo"] = cmbStoreNo.Text.Trim();

                if (dtImport.Rows[i]["PackId"].ToString().Length == 1)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackId"]);
                    str = "0000" + str;
                    dtImport.Rows[i]["PackId1"] = str;
                }
                else if (dtImport.Rows[i]["PackId"].ToString().Length == 2)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackId"]);
                    str = "000" + str;
                    dtImport.Rows[i]["PackId1"] = str;
                }
                else if (dtImport.Rows[i]["PackId"].ToString().Length == 3)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackId"]);
                    str = "00" + str;
                    dtImport.Rows[i]["PackId1"] = str;
                }
                else if (dtImport.Rows[i]["PackId"].ToString().Length == 4)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackId"]);
                    str = "0" + str;
                    dtImport.Rows[i]["PackId1"] = str;
                }
                else
                {
                    dtImport.Rows[i]["PackId1"] = Convert.ToString(dtImport.Rows[i]["PackId"]);
                }
            }

            TatiBAL ObjImport = new TatiBAL();
            ObjImport.AllocationNo = txtAllocationNo.Text.Trim();
            ObjImport.FileName = fileName;
            ObjImport.Location = cmbStoreNo.Text.Trim();
            ObjImport.DtSource = dtImport;

            Boolean Result = ObjImport.InsertAllocationHeader();

            if (Result)
            {
                Result = ObjImport.ImportAllocation();
            }

            if (Result)
            {
                lblMessage.Text = "Successfully Imported!";
                lblMessage.ForeColor = Color.Green;
                btnDwdLog.Visible = false;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;
            }
        }
        #endregion ImportAllocation


        #region Export To Csv
        /// <summary>
        /// Export To Csv
        /// </summary>
        /// <param name="dt"></param>
        private void ExportToCsv(DataTable dt, StreamWriter sw)
        {

            int iColCount = dt.Columns.Count;
            for (int i = 0; i < iColCount; i++)
            {
                sw.Write(dt.Columns[i]);
                if (i < iColCount - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                        sw.Write(dr[i].ToString());
                    if (i < iColCount - 1)
                        sw.Write(",");
                }
                sw.Write(sw.NewLine);
            }
            sw.Write(sw.NewLine);
        }

        #endregion Export To Csv

        #region FileDownload
        /// <summary>
        /// File Download
        /// </summary>

        private void FileDownload(string filename)
        {


            //string FolderPath = HttpContext.Current.Server.MapPath(".");
            //FolderPath = FolderPath + "\\Reports\\";
            //string FullFilePath = FolderPath + filename;
            FileInfo file = new FileInfo(filename);

            if (!file.Exists) return;
            if ((file.Extension == ".DAT") || (file.Extension == ".dat"))
            {
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=\"" + file.Name + "\"");
                Response.AddHeader("Content-Length", file.Length.ToString());
                Response.TransmitFile(file.FullName);
                Response.Flush();
                Response.End();

            }

            if ((file.Extension == ".CSV") || (file.Extension == ".csv"))
            {
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=\"" + file.Name + "\"");
                Response.AddHeader("Content-Length", file.Length.ToString());
                Response.TransmitFile(file.FullName);
                Response.Flush();
                Response.End();
            }

        }


        #endregion FileDownload


        #region BindPONumber
        /// <summary>
        /// BindPONumber
        /// </summary>
        private void BindPONumber()
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt=objDC.GetPOHeader();
            CmbPONumber.DataSource = dt;
            CmbPONumber.DataMember = "PONumber";
            CmbPONumber.DataValueField = "PONumber";
           
            CmbPONumber.DataBind();
          

        }
        
        #endregion BindPONumber

        #region BindStore
        /// <summary>
        /// BindStore
        /// </summary>
        private void BindStore()
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt = objDC.GetStore();
            cmbStoreNo.DataSource = dt;
            cmbStoreNo.DataMember = "StoreNo";
            cmbStoreNo.DataValueField = "StoreNo";
            //cmbStoreNo.DataTextField = "StoreNo";
            cmbStoreNo.DataBind();
            //cmbStoreNo.Items.Insert(0, "<-- Select -->");


        }
        #endregion BindStore

        #region BindSONumber
        /// <summary>
        /// BindSONumber
        /// </summary>
        private void BindSONumber()
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt = objDC.GetSOHeader();
            cmbSONumber.DataSource = dt;
            cmbSONumber.DataMember = "SONumber";
            cmbSONumber.DataValueField = "SONumber";
            
            cmbSONumber.DataBind();
            

        }
        #endregion BindSONumber

        #region ImportSOIssueNote
        /// <summary>
        /// ImportSOIssueNote
        /// </summary>
        private void ImportSOIssueNote()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "SO_IssueNote" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            dtImport.Columns.Add("PackID1");
            dtImport.Columns.Add("PAISSUEDATE1");

            bool Flag = false;

            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                dtImport.Rows[i]["PAISSUEDATE1"] =DateTime.ParseExact(dtImport.Rows[i]["PAISSUEDATE"].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                if (dtImport.Rows[i]["PackID"].ToString().Length == 1)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                    str = "0000" + str;
                    dtImport.Rows[i]["PackID1"] = str;
                }
                else if (dtImport.Rows[i]["PackID"].ToString().Length == 2)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                    str = "000" + str;
                    dtImport.Rows[i]["PackID1"] = str;
                }
                else if (dtImport.Rows[i]["PackID"].ToString().Length == 3)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                    str = "00" + str;
                    dtImport.Rows[i]["PackID1"] = str;
                }
                else if (dtImport.Rows[i]["PackID"].ToString().Length == 4)
                {
                    string str = Convert.ToString(dtImport.Rows[i]["PackID"]);
                    str = "0" + str;
                    dtImport.Rows[i]["PackID1"] = str;
                }
                else
                {
                    dtImport.Rows[i]["PackID1"] = Convert.ToString(dtImport.Rows[i]["PackID"]);
                }

                if(dtImport.Rows[i]["SONUMBER"].ToString()!=cmbSONumber.Text.Trim() )
                {
                    Flag = true;
                    break;
                }

                if (dtImport.Rows[i]["PAISSUENO"].ToString() != txtIssueNoteNo.Text.Trim())
                {
                    Flag = true;
                    break;
                }
            }

            if (!Flag && dtImport.Rows.Count>0)
            {

                TatiBAL ObjImport = new TatiBAL();
                ObjImport.IssueNo = txtIssueNoteNo.Text.Trim();
                ObjImport.FileName = fileName;
                ObjImport.SONumber = cmbSONumber.Text.Trim();
                ObjImport.DtSource = dtImport;

                Boolean Result = false;

                Result = ObjImport.InsertIssueNoteHeader();

                if (Result)
                {
                    Result = ObjImport.ImportSOIssueNote();
                }

                if (Result)
                {
                    String filePath = Server.MapPath(".") + "\\Reports\\IssueNote_Log" + txtIssueNoteNo.Text.Trim() + "_" + rnd.Next() + ".csv";

                    DataTable dt = ObjImport.GetSOIssueNoteLog();
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        ViewState["FileName"] = filePath;
                        StreamWriter sw = new StreamWriter(@filePath, false);

                        ExportToCsv(dt, sw);
                        sw.Close();

                        btnDwdLog.Visible = true;
                        lblMessage.Text = "Log File Generated";
                        lblMessage.ForeColor = Color.Red;
                        btnSaveStockLedger.Visible = false;
                    }
                    else
                    {
                        lblMessage.Text = "Successfully Imported!";
                        lblMessage.ForeColor = Color.Green;
                        btnDwdLog.Visible = false;
                        btnSaveStockLedger.Visible = true;
                    }
                }
                else
                {
                    lblMessage.Text = "Faild!";
                    lblMessage.ForeColor = Color.Red;
                    btnSaveStockLedger.Visible = false;
                }
            }
            else
            {
                lblMessage.Text = "SO Number or Issue Note No. Not Proper / Incorrect File Format";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;
            }

        }
        #endregion ImportSOIssueNote


        #region Validate
        /// <summary>
        /// Validate
        /// </summary>
        /// <returns></returns>
        private bool Validate()
        {
            bool Result = true;
            try
            {
                switch (ddlType.SelectedItem.Value)
                {
                    case "2":
                        if (txtPONumber.Text.Trim().Length == 0 || txtContainerReference.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper PO Number or Container Reference ";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "3":
                        if (CmbPONumber.Text.Trim().Length == 0 || txtGRNNo.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper PO Number or GRN No.";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "4":
                        if (txtAllocationNo.Text.Trim().Length == 0 || cmbStoreNo.Text.Length == 0)
                        {
                            lblMessage.Text = "Put Proper Allocation Number Or Store No";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;

                    case "5":
                        if (txtIssueNoteNo.Text.Trim().Length == 0 || cmbSONumber.Text.Length == 0)
                        {
                            lblMessage.Text = "Put Proper Issue Note Number Or SO Number.";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "6":
                        if (cmbStoreNo.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper Store No.";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "7":
                        if (cmbStoreNo.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper Store No.";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "8":
                        if (cmbStoreNo.Text.Trim().Length == 0 || txtSONumber.Text.Trim().Length==0 || txtGRNNo.Text.Trim().Length== 0 || txtAsOfDate.Text.Trim().Length == 0 || ddlCompany.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper Store No. or SO Number or Grn No Or Date";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "9":
                        if (txtPONumber.Text.Trim().Length == 0 )
                        {
                            lblMessage.Text = "Put Proper PO Number";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "10":
                        if (txtSONumber.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper SO Number";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "11":
                        if (txtAllocationNo.Text.Trim().Length == 0)
                        {
                            lblMessage.Text = "Put Proper Allocation No.";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "12":
                        if (txtPONumber.Text.Trim().Length == 0 ||txtContainerReference.Text.Trim().Length==0)
                        {
                            lblMessage.Text = "Put Proper PONumber or Container No.";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                    case "13":
                        if (txtAdjustmentNo.Text.Trim().Length == 0 ||txtAdjustmentDate.Text.Trim().Length==0 )
                        {
                            lblMessage.Text = "Put Proper Adjustment No. or Adjustment Date";
                            lblMessage.ForeColor = Color.Red;
                            Result = false;
                        }
                        break;
                }
            }
            catch(Exception ex)
            {
                lblMessage.Text = ex.ToString();
                lblMessage.ForeColor = Color.Red;
            }
            return Result;
        }
        #endregion Validate


        #region ImportProductGroupListing
        /// <summary>
        /// ImportProductGroupListing
        /// </summary>
        private void ImportProductGroupListing()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "ProductGroupListing" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            dtImport.Columns.Add("StoreNo1");

            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                if (dtImport.Rows[i]["StoreNo"].ToString().Length ==3)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["StoreNo"]);
                    str = "0" + str;
                    dtImport.Rows[i]["StoreNo1"] = str;
                }
                else
                {
                    dtImport.Rows[i]["StoreNo1"] = Convert.ToString(dtImport.Rows[i]["StoreNo"]);
                }
            }


            TatiBAL ObjImport = new TatiBAL();
            ObjImport.Location = cmbStoreNo.Text.Trim();
            ObjImport.DtSource = dtImport;

            Boolean Result = ObjImport.ImportProductGroupListing();

         
            if (Result)
            {
                lblMessage.Text = "Successfully Imported!";
                lblMessage.ForeColor = Color.Green;
                btnDwdLog.Visible = false;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;
            }
        }
        #endregion ImportProductGroupListing

        #region ImportFamilyListing
        /// <summary>
        /// ImportFamilyListing
        /// </summary>
        private void ImportFamilyListing()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "FamilyListing" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            dtImport.Columns.Add("StoreCode1");

            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                if (dtImport.Rows[i]["StoreCode"].ToString().Length == 3)
                {

                    string str = Convert.ToString(dtImport.Rows[i]["StoreCode"]);
                    str = "0" + str;
                    dtImport.Rows[i]["StoreCode1"] = str;
                }
                else
                {
                    dtImport.Rows[i]["StoreCode1"] = Convert.ToString(dtImport.Rows[i]["StoreCode"]);
                }
            }

            TatiBAL ObjImport = new TatiBAL();
            ObjImport.Location = cmbStoreNo.Text.Trim();
            ObjImport.DtSource = dtImport;
            Boolean Result = ObjImport.ImportFamilyListing();

            if (Result)
            {
                lblMessage.Text = "Successfully Imported!";
                lblMessage.ForeColor = Color.Green;
                btnDwdLog.Visible = false;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;
            }
        }
        #endregion ImportFamilyListing

        #region CreateSOFromPO
        /// <summary>
        /// CreateSOFromPO
        /// </summary>
        private void CreateSOFromPO()
        {

            TatiBAL ObjImport = new TatiBAL();
            ObjImport.CompanyName = ddlCompany.SelectedItem.Value;
            ObjImport.Location = cmbStoreNo.Text.Trim();
            ObjImport.GrnNo= txtGRNNo.Text.Trim();
            ObjImport.SONumber = txtSONumber.Text.Trim();

            ObjImport.AsOfDate = DateTime.ParseExact(txtAsOfDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

            Boolean Result = ObjImport.CreateSOFromPO();

            if (Result)
            {
                lblMessage.Text = "Successfully Posted!";
                lblMessage.ForeColor = Color.Green;
                btnDwdLog.Visible = false;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;
            }
        }
        #endregion CreateSOFromPO


        #region DeleteDocs
        /// <summary>
        /// DeleteDocs
        /// </summary>
        private void DeleteDocs(string DocNo,string Type)
        {
           
            TatiBAL ObjImport = new TatiBAL();
            ObjImport.DocNo = DocNo;
            ObjImport.ReportType = Type;
            Boolean Result = ObjImport.DeleteDocs();

            if (Result)
            {
                lblMessage.Text = "Successfully Deleted!";
                lblMessage.ForeColor = Color.Green;
                btnDwdLog.Visible = false;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
                btnSaveStockLedger.Visible = false;
            }
        }
        #endregion DeleteDocs

        #region ImportAdjustment
        /// <summary>
        /// ImportAdjustment
        /// </summary>
        private void ImportAdjustment()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "Adjustment_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            dtImport.Columns.Add("AdjustmentNo");
            dtImport.Columns.Add("AdjustmentDate");
            dtImport.Columns.Add("OrgDocNo");
            dtImport.Columns.Add("LineNo");
            

            for (int i=0;i<dtImport.Rows.Count;i++)
            {
                dtImport.Rows[i]["AdjustmentNo"] = txtAdjustmentNo.Text.Trim();
                dtImport.Rows[i]["AdjustmentDate"] = txtAdjustmentDate.Text;
                dtImport.Rows[i]["LineNo"] = i+1;

                if (rdlSelectDoc.SelectedItem.Value=="3")
                    dtImport.Rows[i]["OrgDocNo"] = txtAdjustmentNo.Text.Trim();
                else
                    dtImport.Rows[i]["OrgDocNo"] = cmbDocNumber.Text;
            }


            if (dtImport.Rows.Count > 0)
            {
                TatiBAL ObjImport = new TatiBAL();
                ObjImport.DtSource = dtImport;
                Boolean Result = ObjImport.ImportAdjustment();
                ObjImport.DocNo= txtAdjustmentNo.Text.Trim();
                ObjImport.AdjustmentNo = txtAdjustmentNo.Text.Trim();

                if (Result)
                    {
                       
                        ObjImport.DocNo = cmbDocNumber.Text;
                        ObjImport.FileName = fileName;
                        ObjImport.InsertAdjustmentHeader();
                    }

                    if (Result)
                    {

                        String filePath = Server.MapPath(".") + "\\Reports\\Adjustment_Log" + txtAdjustmentNo.Text.Trim() + "_" + rnd.Next() + ".csv";
                        DataTable dt = ObjImport.GetCheckAdjustmentDetails();
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                ViewState["FileName"] = filePath;
                                StreamWriter sw = new StreamWriter(@filePath, false);

                                ExportToCsv(dt, sw);
                                sw.Close();

                                btnDwdLog.Visible = true;
                                lblMessage.Text = "Log File Generated";
                                lblMessage.ForeColor = Color.Red;
                                btnSaveStockLedger.Visible = false;
                            }
                            else
                            {

                                lblMessage.Text = "Successfully Imported!";
                                lblMessage.ForeColor = Color.Green;
                                btnDwdLog.Visible = false;
                                btnSaveStockLedger.Visible = true;
                            }
                    }
                    else
                    {
                        lblMessage.Text = "Faild!";
                        lblMessage.ForeColor = Color.Red;
                    }
                    
                }
            
                
        }

        #endregion ImportAdjustment


        #region ImportHSCodes
        /// <summary>
        /// ImportHSCodes
        /// </summary>
        private void ImportHSCodes()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "HSCodes" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            if (dtImport.Rows.Count > 0)
            {
                TatiBAL ObjImport = new TatiBAL();
                ObjImport.DtSource = dtImport;
                Boolean Result = ObjImport.ImportHSCode();

                if (Result)
                {
                    lblMessage.Text = "Successfully Imported!";
                    lblMessage.ForeColor = Color.Green;
                }
                else
                {
                    lblMessage.Text = "Faild!";
                    lblMessage.ForeColor = Color.Red;
                }
            }
            else
            {

                lblMessage.Text = "Please Check The File Formats";
                lblMessage.ForeColor = Color.Red;
            }
            

        }
        #endregion ImportHSCodes

        #endregion Methods


    }

}