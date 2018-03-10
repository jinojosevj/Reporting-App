
#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ReportingTool.BAL;
using System.Globalization;
using System.Data;
using System.IO;
#endregion NameSpace
namespace ReportingTool
{
    public partial class TatiStockAndSales : System.Web.UI.Page
    {
        #region Events


        protected void Page_Load(object sender, EventArgs e)
        {

        }


        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            TatiBAL objTati = new TatiBAL();
            objTati.Location = "4728";
            objTati.PostingDate=DateTime.ParseExact(txtPostingDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DataTable dt=  objTati.GetTatiSalesFile();
            ExportToDAT(dt,"4728");

            objTati.Location = "4729";
            dt = objTati.GetTatiSalesFile();
            ExportToDAT(dt, "4729");

            objTati.Location = "4731";
            dt = objTati.GetTatiSalesFile();
            ExportToDAT(dt, "4731");

            btnSales.Visible = true;
            btnDwdSales4729.Visible = true;
            btnDwdSales4731.Visible = true;
        }
        #endregion btnGenerate_Click


        #region btnSales_Click
        /// <summary>
        /// btnSales_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnSales_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileName4728"].ToString();
            FileDownload(filename);
        }
        #endregion btnSales_Click


        #region btnGenerateStock_Click
        /// <summary>
        /// btnGenerateStock_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerateStock_Click(object sender, EventArgs e)
        {
            TatiBAL objTati = new TatiBAL();
            objTati.PostingDate = DateTime.ParseExact(txtPostingDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            
            objTati.Location = "4728";
            DataTable dt = objTati.GetTatiStockFile();
            ExportToCSV(dt,"4728");

            objTati.Location = "4729";
            dt = objTati.GetTatiStockFile();
            ExportToCSV(dt,"4729");

            objTati.Location = "4731";
            dt = objTati.GetTatiStockFile();
            ExportToCSV(dt, "4731");
            
            btnStock.Visible = true;
            btnDwdStock4729.Visible = true;
            btnDwdStock4731.Visible = true;

        }
        #endregion btnGenerateStock_Click

        #region btnStock_Click
        /// <summary>
        /// btnStock_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnStock_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameStock4728"].ToString();
            FileDownload(filename);
        }
        #endregion btnStock_Click


        #region btnDwdSales4729_Click
        /// <summary>
        /// btnDwdSales4729_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDwdSales4729_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileName4729"].ToString();
            FileDownload(filename);
        }
        #endregion btnDwdSales4729_Click

        #region btnDwdStock4729_Click
        /// <summary>
        /// btnDwdStock4729_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDwdStock4729_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameStock4729"].ToString();
            FileDownload(filename);
        }
        #endregion btnDwdStock4729_Click

        #region btnDwdStock4731_Click
        /// <summary>
        /// btnDwdSales4731_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDwdSales4731_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileName4731"].ToString();
            FileDownload(filename);
        }

        #endregion btnDwdStock4731_Click

        #region btnDwdStock4731_Click
        /// <summary>
        /// btnDwdStock4731_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDwdStock4731_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameStock4731"].ToString();
            FileDownload(filename);
        }
        #endregion btnDwdStock4731_Click

        #endregion Events


        #region Methods

        #region ExportToDAT
        /// <summary>
        /// ExportToDAT
        /// </summary>
        /// <param name="dt"></param>
        private void ExportToDAT(DataTable dt,string location)
        {
            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\"+location+"FVENTE" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".DAT";
            ViewState["FileName"+location] = filePath;
            StreamWriter sw = new StreamWriter(@filePath, false);

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
                    //if (i < iColCount - 1)
                    //    sw.Write(",");
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
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



        #region ExportToCSV
        /// <summary>
        /// ExportToCSV
        /// </summary>
        /// <param name="dt"></param>
        private void ExportToCSV(DataTable dt,string Location)
        {
            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\000000"+Location+".LCV_IMGSTK_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            ViewState["FileNameStock"+Location] = filePath;
            StreamWriter sw = new StreamWriter(@filePath, false);

            string rowCount = (dt.Rows.Count+1).ToString();
            string strRowCount = "";

            switch(rowCount.Length)
            {
                case 1: strRowCount = "00000"+rowCount;
                    break;
                case 2: strRowCount = "0000" + rowCount;
                    break;
                case 3: strRowCount = "000" + rowCount;
                    break;
                case 4: strRowCount = "00" + rowCount;
                    break;
                case 5: strRowCount = "0" + rowCount;
                    break;

            }


            int iColCount = dt.Columns.Count;
            for (int i = 0; i < iColCount; i++)
            {
                switch(i)
                { 
                    case 0:  sw.Write("000001");
                        break;
                    case 1: sw.Write(DateTime.Now.ToString("yyyyMMddHHmmss"));
                        break;
                    case 2: sw.Write("000000"+Location);
                        break;
                    case 3: sw.Write(strRowCount);
                        break;
                }
                if (i < iColCount - 1)
                {
                    sw.Write(";");
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
                        sw.Write(";");
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        #endregion ExportToDAT

       
      

       

      

        

       


        #endregion Methods

    }
}