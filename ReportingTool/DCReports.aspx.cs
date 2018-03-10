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

#endregion NameSpace
namespace ReportingTool
{
    public partial class DCReports : System.Web.UI.Page
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

        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            GetStockDetails ObjStock = new GetStockDetails();

            ObjStock.PackID = txtPackId.Text.Trim();
            ObjStock.CevaIssueNo = txtCevaNo.Text.Trim();
            ObjStock.IntType = Convert.ToInt32(ddlType.SelectedItem.Value);
            ObjStock.FromDate = txtFromDate.Text.Length > 0 ? Convert.ToDateTime(txtFromDate.Text) : default(DateTime);
            ObjStock.ToDate = txtToDate.Text.Length > 0 ? Convert.ToDateTime(txtToDate.Text) : default(DateTime);

            ObjStock.PackBarcode = txtPackBarcode.Text.Trim();
            ObjStock.PONumber = txtPONumber.Text.Trim();
            ObjStock.LineCode7 = txtLinecode7.Text.Trim();

            DataTable dt=null;
            if (ddlType.SelectedItem.Value=="1")
            {
                if(txtLinecode7.Text.Trim().Length==0 && txtPackBarcode.Text.Trim().Length == 0 && txtPackId.Text.Trim().Length == 0)
                {
                    lblMessage.Text="Please Put Proper Filters(Linecode7 or Packbarcode or PackId)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetDCReports();
                }
            }
            else if(ddlType.SelectedItem.Value == "2")
            {
                if (txtCevaNo.Text.Trim().Length == 0 )
                {
                    lblMessage.Text = "Please Put Proper Filters(CEVA Issue No)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetDCReports();
                }

            }
            else if (ddlType.SelectedItem.Value == "3")
            {
                if (txtPackBarcode.Text.Trim().Length == 0)
                {
                    lblMessage.Text = "Please Put Proper Filters(PackBarcode)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetDCReports();
                }

            }
            else if (ddlType.SelectedItem.Value == "4")
            {
                  lblMessage.Text = "";
                  dt = ObjStock.GetDCReports();
            }

            else if (ddlType.SelectedItem.Value == "5")
            {
                if (txtFromDate.Text.Trim().Length == 0 || txtToDate.Text.Trim().Length == 0)
                {
                    lblMessage.Text = "Please Put Proper Filters(FromDate and ToDate)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetDCReports();
                }

            }

            else if (ddlType.SelectedItem.Value == "6")
            {
                if (txtPONumber.Text.Trim().Length == 0)
                {
                    lblMessage.Text = "Please Put Proper Filters(PO Number)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetDCReports();
                }

            }

            if (dt != null && dt.Rows.Count > 0 )
            {

                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\" + ddlType.SelectedItem.Text + "_" + rnd.Next() + ".csv";
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

        #endregion Events

        #region Methods

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

        #endregion Methods

    }
}