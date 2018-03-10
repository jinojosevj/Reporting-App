#region NameSpace
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Test.BAL;
#endregion NameSpace

namespace ReportingTool
{
    public partial class ItemInfo : System.Web.UI.Page
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

        #region ddlCompany_SelectedIndexChanged
        /// <summary>
        /// ddlCompany_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.Country = ddlCompany.SelectedItem.Text;
            DataTable dt = ObjStock.GetStoreByCountry();
            ddlLocation.DataSource = dt;
            ddlLocation.DataMember = "LocationCode";
            ddlLocation.DataValueField = "LocationCode";
            ddlLocation.DataBind();
        }
        #endregion ddlCompany_SelectedIndexChanged

        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            ExportItemInfo();
        }
        #endregion btnGenerate_Click

        protected void btnDownload_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileName"].ToString();
            FileDownload(filename);
        }

        #endregion Events

        #region Methods

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
            if ((file.Extension == ".CSV") || (file.Extension == ".csv") )
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

        #region ExportItemInfo
        /// <summary>
        /// ExportItemInfo
        /// </summary>
        private void ExportItemInfo()
        {
            GetStockDetails objStock = new GetStockDetails();
            objStock.Country = ddlCompany.SelectedItem.Text;
            objStock.Location = ddlLocation.SelectedItem.Text;
            objStock.LineCode = txtLineCode.Text.Trim();
            DataTable dt=objStock.GetItemInfo();
            ExportToCsv(dt);
            btnDownload.Visible = true;
        }
        #endregion ExportItemInfo

        #region Export To Csv
        /// <summary>
        /// Export To Csv
        /// </summary>
        /// <param name="dt"></param>
        private void ExportToCsv(DataTable dt)
        {
            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\ItemInfo_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            ViewState["FileName"] = filePath;
            StreamWriter sw = new StreamWriter(@filePath, false);

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
            sw.Close();
        }

        #endregion Export To Csv

        #endregion Methods

    }
}