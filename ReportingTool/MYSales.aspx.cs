using ReportingTool.BAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ReportingTool
{
    public partial class MYSales : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            TatiBAL objMY = new TatiBAL();
            objMY.Location = "F004";
            objMY.PostingDate = DateTime.ParseExact(txtPostingDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objMY.ReportType = "All";
            objMY.ReceiptNo = "0";
            DataTable dt = objMY.GetMYSalesFile();
            DataTable details = null;
            string filePath = Server.MapPath(".") + "\\Reports\\" + "TVFR004_QATAR_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-") + ".PAQ";
            ViewState["FileNameF004" ] = filePath;
            StreamWriter sw = new StreamWriter(@filePath, false);

            sw.Write("<ENTETE>EXP  002.000"+ DateTime.Now.ToString("dd/MM/yyyyHH:mm:ss") + "LCFR FFRA");
            sw.Write(sw.NewLine);

            for (int i=0;i<dt.Rows.Count;i++)
            {
                objMY.ReportType = "H";
                objMY.ReceiptNo = dt.Rows[i]["Receipt No_"].ToString();
                details= objMY.GetMYSalesFile();
                ExportToDAT(details, sw);


                objMY.ReportType = "D";
                details = objMY.GetMYSalesFile();
                ExportToDAT(details, sw);


                objMY.ReportType = "P";
                details = objMY.GetMYSalesFile();
                ExportToDAT(details, sw);

                btnSales.Visible = true;
            }
            sw.Close();


        }

        protected void btnSales_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameF004"].ToString();
            FileDownload(filename);
        }


        #region ExportToDAT
        /// <summary>
        /// ExportToDAT
        /// </summary>
        /// <param name="dt"></param>
        private void ExportToDAT(DataTable dt, StreamWriter sw)
        {
            Random rnd = new Random();
           

            int iColCount = dt.Columns.Count;
            //for (int i = 0; i < iColCount; i++)
            //{
            //    sw.Write(dt.Columns[i]);
            //    if (i < iColCount - 1)
            //    {
            //        sw.Write(",");
            //    }
            //}
            //sw.Write(sw.NewLine);

            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < iColCount; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                        sw.Write(dr[i].ToString());
                    if (i < iColCount - 1)
                        sw.Write("\t");
                }
                sw.Write(sw.NewLine);
            }
           
        }

        #endregion ExportToDAT

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
            if ((file.Extension == ".PAQ") || (file.Extension == ".paq"))
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