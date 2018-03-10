
#region NameSpace
        using System;
        using System.Collections.Generic;
        using System.Linq;
        using System.Web;
        using System.Web.UI;
        using System.Web.UI.WebControls;
        using System.Data;
        using Test.BAL;
        using Excel1 = Microsoft.Office.Interop.Excel;
        using System.Globalization;
        using System.IO;

        using Microsoft.Office.Core;
        using System.Diagnostics;
        using System.Data.OleDb;
        using System.Data.SqlClient;
        using System.Configuration;
        using Excel;
        using System.Drawing;
        //using Microsoft.Office.Interop.Excel;
        using Test.DAL;
        using ReportingTool.BAL;

#endregion NameSpace

namespace ReportingTool
{
    public partial class TransferOrder : System.Web.UI.Page
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
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "TransferSender" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadSender.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);
            TatiBAL ObjImport = new TatiBAL();
            Boolean Result = false;

            dtImport.Columns.Add("Country");
            dtImport.Columns.Add("Location");
            dtImport.Columns.Add("DocNo");

            for (int i=0;i<dtImport.Rows.Count;i++)
            {
                dtImport.Rows[i]["DocNo"] = txtDocumentNo.Text.Trim().ToString();
                dtImport.Rows[i]["Quantity"] = Common.Base64Decode(dtImport.Rows[i]["Quantity"].ToString());
                dtImport.Rows[i]["Barcode"] = Common.Base64Decode(dtImport.Rows[i]["Barcode"].ToString());
                dtImport.Rows[i]["Location"] = ddlLocation.Text.ToString();

                dtImport.Rows[i]["R"] = Common.Base64Decode(dtImport.Rows[i]["R"].ToString());
                dtImport.Rows[i]["Country"] =  ddlCountry.Text.ToString();
            }

            ObjImport.FileName = fileName;
            ObjImport.DocNo = txtDocumentNo.Text.Trim().ToString();
            ObjImport.Location = ddlLocation.Text;

            ObjImport.DtSource = dtImport;

            Result = ObjImport.InsertTransferHeader();

            if (Result)
            {
                Result = ObjImport.ImportTransferOrder();

                if(Result)
                   ObjImport.UpdateTransferOrder();
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
        #endregion btnImport_Click


        #region btnDownload_Click
        /// <summary>
        /// btnDownload_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownload_Click(object sender, EventArgs e)
        {
            string fileName = ViewState["FileName"].ToString();
            FileDownload(fileName);
        }
        #endregion btnDownload_Click

        #region ddlCountry_SelectedIndexChanged
        /// <summary>
        /// ddlCountry_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.Country = ddlCountry.SelectedItem.Text;
            DataTable dt = ObjStock.GetStoreByCountry();
            ddlLocation.DataSource = dt;
            ddlLocation.DataMember = "LocationCode";
            ddlLocation.DataValueField = "LocationCode";
            ddlLocation.DataBind();
        }

        #endregion ddlCountry_SelectedIndexChanged

        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            TatiBAL ObjStock = new TatiBAL();
            ObjStock.DocNo = txtDocumentNo.Text.Trim();
            DataTable dt = ObjStock.GetTransferOrder();

            if (dt != null && dt.Rows.Count > 0)
            {

                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\Transfer_" + txtDocumentNo.Text + "_"+ rnd.Next() + ".csv";
                ViewState["FileName"] = filePath;
                StreamWriter sw = new StreamWriter(@filePath, false);

                ExportToCsv(dt, sw);
                sw.Close();

                btnDownload.Visible = true;
                lblMessage.Text = "Report Generated";
                lblMessage.ForeColor = Color.Green;
            }
        }
        #endregion btnGenerate_Click


        #region btnImportReceiver_Click
        /// <summary>
        /// btnImportReceiver_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnImportReceiver_Click(object sender, EventArgs e)
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "TransferReceiver" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadReceiver.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);
            TatiBAL ObjImport = new TatiBAL();
            Boolean Result = false;


            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                ObjImport.DocNo = txtDocumentNo.Text.Trim().ToString();
                ObjImport.Quantity = Convert.ToDecimal(Common.Base64Decode(dtImport.Rows[i]["Quantity"].ToString()));
                ObjImport.PackBarcode = Common.Base64Decode(dtImport.Rows[i]["Barcode"].ToString());
                ObjImport.FileName = fileName;
                Result = ObjImport.UpdateReceivedQty();
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
        #endregion btnImportReceiver_Click


        #region btnPost_Click
        /// <summary>
        /// btnPost_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnPost_Click(object sender, EventArgs e)
        {
            TatiBAL ObjImport = new TatiBAL();
            Boolean Result = false;

            ObjImport.DocNo = txtDocumentNo.Text.Trim().ToString();
            ObjImport.CompanyName = ddlCountry.SelectedItem.Value;
            ObjImport.Location = ddlLocation.SelectedItem.Value;
            Result = ObjImport.InsertInventoryNAV();
          
            if (Result)
            {
                lblMessage.Text = "Successfully Posted!";
                lblMessage.ForeColor = Color.Green;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
            }

        }
        #endregion btnPost_Click

        #region btnImportAdjustment_Click
        /// <summary>
        /// btnImportAdjustment_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnImportAdjustment_Click(object sender, EventArgs e)
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "TransferAdjustment" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadAdjustment.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);
            TatiBAL ObjImport = new TatiBAL();
            Boolean Result = false;


            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                ObjImport.DocNo = txtDocumentNo.Text.Trim().ToString();
                ObjImport.Quantity = Convert.ToDecimal(dtImport.Rows[i]["VarianceAdjustment"].ToString());
                ObjImport.Remarks = dtImport.Rows[i]["Remarks"].ToString();
                ObjImport.Id = Convert.ToInt32(dtImport.Rows[i]["Id"].ToString());
                Result = ObjImport.UpdateTransferAdjustment();
            }

            if (Result)
            {
                lblMessage.Text = "Adjustment Successfully Updated!";
                lblMessage.ForeColor = Color.Green;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion btnImportAdjustment_Click


        #endregion Events

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

        
    }
}