#region Name Space

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Web.UI;
using System.Web.UI.WebControls;
using Test.BAL;
using System.Data;

using Microsoft.Office.Core;
using Excel1 = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.IO;

#endregion Name Space
namespace ReportingTool
{
    public partial class CustomerCount : System.Web.UI.Page
    {
        public DataTable dtStock = null;

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
            ViewState["FileName"] = null;

            GenerateReport();
            Page.ClientScript.RegisterStartupScript(this.GetType(), "CallMyFunction", "$('#btnGenerate').Show();", true);
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
            string filename = ViewState["FileName"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownload_Click

        #endregion Events

        #region Methods

        #region GenerateReport
        /// <summary>
        /// To generate excel report for stock values
        /// </summary>
        private void GenerateReport()
        {

            //try
            //{

            Excel1.Application myExcelApp;

            Excel1.Workbooks myExcelWorkbooks;

            Excel1.Workbook myExcelWorkbook;


            object misValue = System.Reflection.Missing.Value;

            myExcelApp = new Excel1.Application();

            myExcelApp.Visible = false;

            myExcelWorkbooks = myExcelApp.Workbooks;


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\CustomerCount.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            GetStockDetails objStock = new GetStockDetails();

            objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);


            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\CustomerCount.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;

                dtStock = objStock.GetCustomerCount();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                    WriteToExcel(dtStock, xlSheet, location);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnDownload.Visible = true;
                }
                else
                {
                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "No Data Found";
                }

            }

            else
            {


                Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;

                Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet0400.Name = "0400";
                objStock.Location = "0400";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0400, "0400");

                Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheet0401.Name = "0401";
                objStock.Location = "0401";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0401, "0401");

                Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheet0402.Name = "0402";
                objStock.Location = "0402";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0402, "0402");

                Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet0403.Name = "0403";
                objStock.Location = "0403";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0403, "0403");

                Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheet0404.Name = "0404";
                objStock.Location = "0404";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0404, "0404");

                Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheet0405.Name = "0405";
                objStock.Location = "0405";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0405, "0405");

                Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                xlSheet0406.Name = "0406";
                objStock.Location = "0406";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0406, "0406");

                Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                xlSheet0407.Name = "0407";
                objStock.Location = "0407";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0407, "0407");


                Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                xlSheet0409.Name = "0409";
                objStock.Location = "0409";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0409, "0409");

                Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                xlSheet0410.Name = "0410";
                objStock.Location = "0410";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0410, "0410");

                Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                xlSheet0411.Name = "0411";
                objStock.Location = "0411";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0411, "0411");

                Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                xlSheet0414.Name = "0414";
                objStock.Location = "0414";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0414, "0414");


                Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                xlSheet0415.Name = "0415";
                objStock.Location = "0415";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0415, "0415");

                Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                xlSheet0416.Name = "0416";
                objStock.Location = "0416";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0416, "0416");

                Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                xlSheet0417.Name = "0417";
                objStock.Location = "0417";
                dtStock = objStock.GetCustomerCount();
                WriteToExcel(dtStock, xlSheet0417, "0417");

                Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                xlSheet0412.Name = "0412";
                objStock.Location = "0412";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0412, "0412");
                }

                Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                xlSheet0418.Name = "0418";
                objStock.Location = "0418";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0418, "0418");
                }

                Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                xlSheet0419.Name = "0419";
                objStock.Location = "0419";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0419, "0419");
                }

                Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                xlSheet0421.Name = "0421";
                objStock.Location = "0421";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0421, "0421");
                }

                Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                xlSheet0422.Name = "0422";
                objStock.Location = "0422";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0422, "0422");
                }

                Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                xlSheet0423.Name = "0423";
                objStock.Location = "0423";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0423, "0423");
                }

                Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                xlSheet0424.Name = "0424";
                objStock.Location = "0424";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0424, "0424");
                }

                Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                xlSheet0425.Name = "0425";
                objStock.Location = "0425";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0425, "0425");
                }

                Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                xlSheet0426.Name = "0426";
                objStock.Location = "0426";
                dtStock = objStock.GetCustomerCount();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0426, "0426");
                }


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownload.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\CustomerCount" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileName"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);


            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();



            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateReport

        #region WriteToExcel
        /// <summary>
        /// Write To Excel
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcel(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string location)
        {
            object misValue = System.Reflection.Missing.Value;

            string storeName = "";

            switch(location)
            {
                case "0400": storeName = "Matalan Middle East Al Baraka Mall (" + location+")";
                             break;
                case "0401": storeName = "Matalan Middle East Arabian Centre (" + location + ")";
                             break;
                case "0402": storeName = "Matalan Middle East Dalma Mall (" + location + ")";
                             break;
                case "0403": storeName = "Matalan Middle East Lamcy Plaza (" + location + ")";
                             break;
                case "0404": storeName = "Matalan Middle East Arabella Mall (" + location + ")";
                             break;
                case "0405": storeName = "Matalan Middle East Mushrif Mall (" + location + ")";
                             break;
                case "0406": storeName = "Matalan Middle East Century Mall (" + location + ")";
                             break;
                case "0407": storeName = "Matalan Middle East Markaz Al Bahja Mall (" + location + ")";
                             break;
                case "0409": storeName = "Matalan Middle East Mirdiff City Centre (" + location + ")";
                             break;
                case "0410": storeName = "Matalan Middle East Sahara Centre (" + location + ")";
                             break;
                case "0411": storeName = "Matalan Middle East Galleria Mall (" + location + ")";
                             break;
                case "0412": storeName = "Matalan Middle East Gulf Mall (" + location + ")";
                             break;
                case "0414": storeName = "Matalan Middle East Al Ghurair Centre (" + location + ")";
                             break;
                case "0415": storeName = "Matalan Middle East Khalidiya Mall (" + location + ")";
                             break;
                case "0416": storeName = "Matalan Middle East Bahrain City Centre (" + location + ")";
                             break;
                case "0417": storeName = "Matalan Middle East RAK Mall (" + location + ")";
                             break;
                case "0418": storeName = "Matalan Middle East Al Foah Mall (" + location + ")";
                             break;
                case "0419": storeName = "Matalan Middle East Wafi Mall (" + location + ")";
                             break;
                case "0421": storeName = "Matalan Middle East Avenue Mall (" + location + ")";
                             break;
                case "0422": storeName = "Matalan Middle East Mecca Mall (Ladies) (" + location + ")";
                             break;
                case "0423": storeName = "Matalan Middle East Mecca Mall (Kids) (" + location + ")";
                             break;
                case "0424": storeName = "Matalan Middle East Al Qasar Mall (" + location + ")";
                             break;

            }

            myExcelWorksheet.get_Range("A1", misValue).Formula = storeName;
            myExcelWorksheet.get_Range("A2", misValue).Formula = "Customer Count Data From " + txtFromDate.Text.ToString()+" To "+txtToDate.Text.ToString();
            //BorderAround(myExcelWorksheet.get_Range("A2", misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("A1", misValue).Font.Bold = true;

            int j = 4;

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["StaffId"]) ? dtStock.Rows[i]["StaffId"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
               
                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["NewPhoneAndEmail"]) ? dtStock.Rows[i]["NewPhoneAndEmail"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["NewPhoneNoOnly"]) ? dtStock.Rows[i]["NewPhoneNoOnly"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["Existing"]) ? dtStock.Rows[i]["Existing"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["Nodata"]) ? dtStock.Rows[i]["Nodata"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalTransaction"]) ? dtStock.Rows[i]["TotalTransaction"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=F" + j + "/F" + (4 + dtStock.Rows.Count).ToString();
                BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            }

            myExcelWorksheet.get_Range("A" + j, "G" + j).Interior.Color = System.Drawing.Color.Yellow;
            myExcelWorksheet.get_Range("A" + j, misValue).Formula = "Total";
            BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("A" + j, misValue).EntireRow.Font.Bold = true;

            myExcelWorksheet.get_Range("B" + j, misValue).Formula = "=SUM(B3" + ":B" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("C" + j, misValue).Formula = "=SUM(C3" + ":C" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("D" + j, misValue).Formula = "=SUM(D3" + ":D" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E3" + ":E" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=SUM(F3" + ":F" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G3" + ":G" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            // Mix %

            myExcelWorksheet.get_Range("A" + (j + 1), "G" + (j + 1)).Interior.Color = System.Drawing.Color.Yellow;
            myExcelWorksheet.get_Range("A" + (j+1), misValue).Formula = "Mix %";
            BorderAround(myExcelWorksheet.get_Range("A" + (j + 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("A" +( j+1), misValue).EntireRow.Font.Bold = true;

            myExcelWorksheet.get_Range("B" + (j + 1), misValue).Formula = "=Round(B" + j + "/F" + j + ",2)*100";
            BorderAround(myExcelWorksheet.get_Range("B" + (j + 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("C" + (j + 1), misValue).Formula = "=Round(C" + j + "/F" + j + ",2)*100";
            BorderAround(myExcelWorksheet.get_Range("C" + (j + 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("D" + (j + 1), misValue).Formula = "=Round(D" + j + "/F" + j + ",2)*100";
            BorderAround(myExcelWorksheet.get_Range("D" + (j + 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("E" + (j + 1), misValue).Formula = "=Round(E" + j + "/F" + j + ",2)*100";
            BorderAround(myExcelWorksheet.get_Range("E" + (j + 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("F" + (j + 1), misValue).Formula = "=Round(F" + j + "/F" + j + ",2)*100";
            BorderAround(myExcelWorksheet.get_Range("F" + (j + 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

           // myExcelWorksheet.get_Range("G" + j+1, misValue).Formula = "=SUM(G3" + ":G" + (j - 1) + ")";
           // BorderAround(myExcelWorksheet.get_Range("G" + j+1, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

        }

        #endregion WriteToExcel

        #region BorderAround
        /// <summary>
        /// Border Around
        /// </summary>
        /// <param name="range"></param>
        /// <param name="colour"></param>
        private void BorderAround(Excel1.Range range, int colour)
        {
            Excel1.Borders borders = range.Borders;
            borders[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
            borders[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
            borders[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;
            borders[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
            borders.Color = colour;
            borders[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlLineStyleNone;
            borders[Excel1.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel1.XlLineStyle.xlLineStyleNone;
            borders[Excel1.XlBordersIndex.xlDiagonalUp].LineStyle = Excel1.XlLineStyle.xlLineStyleNone;
            borders[Excel1.XlBordersIndex.xlDiagonalDown].LineStyle = Excel1.XlLineStyle.xlLineStyleNone;
            borders = null;
        }

        #endregion BorderAround

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
            if ((file.Extension == ".xlsx") || (file.Extension == ".XLSX") || (file.Extension == ".xls") || (file.Extension == ".XLS"))
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