﻿#region NameSpace
    using ReportingTool.BAL;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Web;
    using System.Web.UI;
    using System.Web.UI.WebControls;
#endregion NameSpace
namespace ReportingTool
{
    public partial class TATIMISReports : System.Web.UI.Page
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

        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            TatiBAL ObjStock = new TatiBAL();

            ObjStock.ReportType = ddlType.SelectedItem.Value;
            ObjStock.Country = ddlCountry.SelectedItem.Value;
            ObjStock.Location = ddlLocation.SelectedItem.Value;

            ObjStock.FromDate = txtFromDate.Text.Trim().Length > 0 ? Convert.ToDateTime(txtFromDate.Text) : default(DateTime);
            ObjStock.ToDate = txtToDate.Text.Trim().Length > 0 ? Convert.ToDateTime(txtToDate.Text) : default(DateTime);

            //ObjStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            //ObjStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);


            ObjStock.JordanOffer = txtPromoNumbers.Text.Trim();
            //ObjStock.LineCode7 = txtLinecode7.Text.ToString().Trim();
            //ObjStock.DivisionCode = ddlDivision.SelectedItem.Value.Trim();
            DataTable dt = ObjStock.GetTATIMISReports();

            if (dt.Rows.Count > 0)
            {
                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\" + ddlType.SelectedItem.Text + "_" + rnd.Next() + ".csv";
                ViewState["FileName"] = filePath;
                StreamWriter sw = new StreamWriter(@filePath, false);

                ExportToCsv(dt, sw);
                sw.Close();

                btnDownload.Visible = true;
            }

            
            lblMessage.Text = "Report Generated";
            lblMessage.ForeColor = Color.Green;
        }
        #endregion btnGenerate_Click


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

        #region ddlType_SelectedIndexChanged
        /// <summary>
        /// ddlType_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlType_SelectedIndexChanged(object sender, EventArgs e)
        {
            trCountry.Visible = false;
            trLocation.Visible = false;
            trPromotion.Visible = false;
            trFromDate.Visible = false;
            trToDate.Visible = false;

            ViewState["FileName"] = null;
            btnDownload.Visible = false;

            switch (ddlType.SelectedItem.Value)
            {
                case "1":trCountry.Visible = true;
                         trLocation.Visible = true;
                         trToDate.Visible = true;
                    break;
                case "3":
                        trFromDate.Visible = true;
                        trToDate.Visible = true;
                        break;
                case "4":
                        trPromotion.Visible = true;
                        break;
                case "5":
                        trCountry.Visible = true;
                        trLocation.Visible = true;
                        trFromDate.Visible = true;
                        trToDate.Visible = true;
                        break;
                case "6":
                        trCountry.Visible = true;
                        trLocation.Visible = true;
                        trFromDate.Visible = true;
                        trToDate.Visible = true;
                        break;
                case "7":
                        trCountry.Visible = true;
                        break;
            }
        }
        #endregion ddlType_SelectedIndexChanged

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