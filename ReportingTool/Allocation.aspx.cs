using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ReportingTool.BAL;
using System.Data;
using System.IO;
using System.Drawing;

namespace ReportingTool
{
    public partial class Allocation : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //ViewState["FileNameDCS"] = null;
        }

        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            TatiBAL objDC = new TatiBAL();
            objDC.Location = txtStoreCode.Text.Trim().ToString();
            DataTable dt = null;
            ViewState["FileNameDCP"] = null;
            ViewState["FileNameDCS"] = null;

               dt = objDC.GetDCAllocationSingle();

                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\" + txtStoreCode.Text+"_"+DateTime.Now.ToString("ddMMyyyy")+"_S_" + rnd.Next() + ".csv";
                ViewState["FileNameDCS"] = filePath;
                StreamWriter sw = new StreamWriter(@filePath, false);

                ExportToCsv(dt, sw);
                sw.Close();

            //dt = objDC.GetDCStockSingle();

            //string filePathDCS = Server.MapPath(".") + "\\Reports\\" + "DCStock_Single_" + rnd.Next() + ".csv";
            //ViewState["FileNameDCS"] = filePathDCS;
            //StreamWriter sw = new StreamWriter(@filePathDCS, false);

            //ExportToCsv(dt, sw);
            //sw.Close();


                objDC.SellThrough = Convert.ToDecimal(txtSellThrough.Text.Trim());
                dt = objDC.GetDCAllocation();

                rnd = new Random();
                filePath = Server.MapPath(".") + "\\Reports\\" + txtStoreCode.Text + "_" + DateTime.Now.ToString("ddMMyyyy") + "_P_" + rnd.Next() + ".csv";
                ViewState["FileNameDCP"] = filePath;
                sw = new StreamWriter(@filePath, false);

                ExportToCsv(dt, sw);
                sw.Close();

                //dt = objDC.GetDCStockProcess();

                //string filePathDCS = Server.MapPath(".") + "\\Reports\\" + "DCStock" + rnd.Next() + ".csv";
                //ViewState["FileNameDCS"] = filePathDCS;
                //StreamWriter sw = new StreamWriter(@filePathDCS, false);

                //ExportToCsv(dt, sw);
                //sw.Close();
           
           // btnDownload.Visible = true;
            btnDownloadDC.Visible = true;
            btnDownloadDCS.Visible = true;

            lblMessage.Text = "Report Generated";
            lblMessage.ForeColor = Color.Green;

        }

       
        protected void btnDownloadDC_Click(object sender, EventArgs e)
        {
            string fileName = ViewState["FileNameDCP"].ToString();
            FileDownload(fileName);
        }

        protected void btnDownloadDCS_Click(object sender, EventArgs e)
        {
            string fileName = ViewState["FileNameDCS"].ToString();
            FileDownload(fileName);
        }

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