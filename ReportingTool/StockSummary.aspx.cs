#region NameSpace
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
    
#endregion NameSpace

namespace ReportingTool
{
    public partial class StockSummary : System.Web.UI.Page
    {
        public DataTable dtStock = null;
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

            //String fileName = "C:\\book1.xlsx";
            // myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.ActiveSheet;

            //myExcelWorkbooks.Close();




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\Stock_Summary.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objStock = new GetStockDetails();
           
            objStock.AsOfDate = DateTime.ParseExact(txtAsofDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\Stock_Summary.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;
                
                switch(location)
                {
                    case "0400": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;
                    case "0401": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0402": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0403": objStock.CompanyName = "Matalan-UAE HO";
                                 break;

                    case "0404": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;
                    case "0405": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0406": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0407": objStock.CompanyName = "Matalan-Oman HO";
                                 break;

                    case "0408": objStock.CompanyName = "Matalan DC-JAFZA";
                                 break;
                    case "0409": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0410": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0411": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;

                    case "0414": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0415": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0416": objStock.CompanyName = "Matalan-BAHRAIN HO";
                                 break;
                    case "0417": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    
                    case "0412": objStock.CompanyName = "Matalan-QATAR HO";
                                 break;
                    case "0418": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0419": objStock.CompanyName = "Matalan-UAE HO";
                                 break;
                    case "0421": objStock.CompanyName = "Matalan-Oman HO";
                                 break;
                    
                    case "0422": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;
                    case "0423": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;
                    case "0424": objStock.CompanyName = "MATALAN-KSA HO";
                                 break;
                    case "0425": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;
                   
                    case "0426": objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;

                    default: objStock.CompanyName = "MATALAN-JORDAN HO";
                                 break;
                }

                dtStock = objStock.GetStockSummary();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;

                    //xlSheet.get_Range("A1", misValue).Formula = "Store No: " + location + " Stock Summary As Of "+txtAsofDate.Text;
                    WriteToExcel(dtStock, xlSheet, location);


                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnDownload.Visible = true;
                    //btnDownloadSSRStore.Visible = true;
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
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0400, "0400");

                Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheet0401.Name = "0401";
                objStock.Location = "0401";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0401, "0401");

                Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheet0402.Name = "0402";
                objStock.Location = "0402";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0402, "0402");

                Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet0403.Name = "0403";
                objStock.Location = "0403";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0403, "0403");

                Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheet0404.Name = "0404";
                objStock.Location = "0404";
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0404, "0404");

                Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheet0405.Name = "0405";
                objStock.Location = "0405";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0405, "0405");

                Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                xlSheet0406.Name = "0406";
                objStock.Location = "0406";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0406, "0406");

                Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                xlSheet0407.Name = "0407";
                objStock.Location = "0407";
                objStock.CompanyName = "Matalan-Oman HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0407, "0407");

                Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                xlSheet0408.Name = "0408";
                objStock.Location = "0408";
                objStock.CompanyName = "Matalan DC-JAFZA";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0408, "0408");


                Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                xlSheet0409.Name = "0409";
                objStock.Location = "0409";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0409, "0409");

                Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                xlSheet0410.Name = "0410";
                objStock.Location = "0410";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0410, "0410");

                Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                xlSheet0411.Name = "0411";
                objStock.Location = "0411";
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0411, "0411");

                Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                xlSheet0414.Name = "0414";
                objStock.Location = "0414";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0414, "0414");


                Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                xlSheet0415.Name = "0415";
                objStock.Location = "0415";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0415, "0415");

                Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                xlSheet0416.Name = "0416";
                objStock.Location = "0416";
                objStock.CompanyName = "Matalan-BAHRAIN HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0416, "0416");

                Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                xlSheet0417.Name = "0417";
                objStock.Location = "0417";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                WriteToExcel(dtStock, xlSheet0417, "0417");

                Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                xlSheet0412.Name = "0412";
                objStock.Location = "0412";
                objStock.CompanyName = "Matalan-QATAR HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0412, "0412");
                }

                Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                xlSheet0418.Name = "0418";
                objStock.Location = "0418";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0418, "0418");
                }

                Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                xlSheet0419.Name = "0419";
                objStock.Location = "0419";
                objStock.CompanyName = "Matalan-UAE HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0419, "0419");
                }

                Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                xlSheet0421.Name = "0421";
                objStock.Location = "0421";
                objStock.CompanyName = "Matalan-Oman HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0421, "0421");
                }

                Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                xlSheet0422.Name = "0422";
                objStock.Location = "0422";
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0422, "0422");
                }

                Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                xlSheet0423.Name = "0423";
                objStock.Location = "0423";
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0423, "0423");
                }

                Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                xlSheet0424.Name = "0424";
                objStock.Location = "0424";
                objStock.CompanyName = "MATALAN-KSA HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0424, "0424");
                }


                Excel1.Worksheet xlSheet0425= (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                xlSheet0425.Name = "0425";
                objStock.Location = "0425";
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0425, "0425");
                }

                Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                xlSheet0426.Name = "0426";
                objStock.Location = "0426";
                objStock.CompanyName = "MATALAN-JORDAN HO";
                dtStock = objStock.GetStockSummary();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcel(dtStock, xlSheet0426, "0426");
                }


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownload.Visible = true;
                // btnDownloadSSRStore.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\Stock_Summery" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

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

            myExcelWorksheet.get_Range("A1", misValue).Formula = "Store No: " + location + " Stock Summary As Of " + txtAsofDate.Text;
            BorderAround(myExcelWorksheet.get_Range("A1", misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("A1", misValue).Font.Bold = true;
           
            for (int i = 0, j = 3; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Location"]) ? dtStock.Rows[i]["Location"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["MajorDept"]) ? dtStock.Rows[i]["MajorDept"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["DivisionDescription"]) ? dtStock.Rows[i]["DivisionDescription"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["Qty"]) ? dtStock.Rows[i]["Qty"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesValue"]) ? dtStock.Rows[i]["SalesValue"].ToString()  : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
               
            }

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