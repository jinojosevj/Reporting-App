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
namespace ReportingTool
{
    public partial class DcStockReport : System.Web.UI.Page
    {
        public DataTable dtStock = null;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            GenerateReport();
        }


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




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\DcStockReport.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objStock = new GetStockDetails();

            objStock.AsOfDate = DateTime.ParseExact(txtAsofDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtLocation.Text.Trim().Length > 0)
            {
                //string location = txtLocation.Text.Trim();

                //fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\DcStockReport.xlsx";
                //myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                //Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                //objStock.Location = location;

                //dtStock = objStock.GetDCStock();

                //if (dtStock.Rows.Count > 0)
                //{
                //    xlSheet.Name = location;

                //    //xlSheet.get_Range("A1", misValue).Formula = "Store No: " + location + " Stock Summary As Of "+txtAsofDate.Text;
                //    WriteToExcel(dtStock, xlSheet, location);


                //    lblMessage.Visible = true;
                //    lblMessage.ForeColor = System.Drawing.Color.Green;
                //    lblMessage.Text = "Report Generation Complete";

                //    btnDownload.Visible = true;
                //    //btnDownloadSSRStore.Visible = true;
                //}
                //else
                //{
                //    lblMessage.Visible = true;
                //    lblMessage.ForeColor = System.Drawing.Color.Red;
                //    lblMessage.Text = "No Data Found";
                //}

                //Random rnd = new Random();
                //string filePath = Server.MapPath(".") + "\\Reports\\Stock_Summery" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

                //ViewState["FileName"] = filePath;
                //myExcelWorkbook.SaveAs(@filePath);


                //myExcelWorkbook.Close();
                //myExcelWorkbooks.Close();
            }

            else
            {


                Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;

                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0400";
                objStock.Location = "0400";
                objStock.Type = true;

                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);

                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\0400_StoreStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

                //ViewState["FileName"] = filePath;
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();

                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0400";
                objStock.Location = "0400";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0400_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);

                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0401";
                objStock.Location = "0401";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0401_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0401";
                objStock.Location = "0401";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0401_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0402";
                objStock.Location = "0402";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0402_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0402";
                objStock.Location = "0402";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0402_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0403";
                objStock.Location = "0403";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0403_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0403";
                objStock.Location = "0403";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0403_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0404";
                objStock.Location = "0404";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0404_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0404";
                objStock.Location = "0404";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0404_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0405";
                objStock.Location = "0405";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0405_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0405";
                objStock.Location = "0405";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0405_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0406";
                objStock.Location = "0406";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0406_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0406";
                objStock.Location = "0406";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0406_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0407";
                objStock.Location = "0407";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0407_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0407";
                objStock.Location = "0407";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0407_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0409";
                objStock.Location = "0409";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0409_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0409";
                objStock.Location = "0409";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0409_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);

                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0410";
                objStock.Location = "0410";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0410_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0410";
                objStock.Location = "0410";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0410_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0411";
                objStock.Location = "0411";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0411_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0411";
                objStock.Location = "0411";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0411_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0412";
                objStock.Location = "0412";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0412_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0412";
                objStock.Location = "0412";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0412_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);

                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0414";
                objStock.Location = "0414";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0414_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0414";
                objStock.Location = "0414";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0414_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0415";
                objStock.Location = "0415";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0415_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0415";
                objStock.Location = "0415";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0415_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0416";
                objStock.Location = "0416";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0416_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0416";
                objStock.Location = "0416";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0416_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0417";
                objStock.Location = "0417";
                objStock.Type = true;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, true);
                filePath = Server.MapPath(".") + "\\Reports\\0417_StockStock_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet.Name = "0417";
                objStock.Location = "0417";
                objStock.Type = false;
                dtStock = objStock.GetDCStock();
                WriteToExcel(dtStock, xlSheet, false);
                filePath = Server.MapPath(".") + "\\Reports\\0417_StockCover_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                myExcelWorkbook.SaveAs(@filePath);
                myExcelWorkbook.Close();
                myExcelWorkbooks.Close();
                
                
                
                
                
                
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownload.Visible = true;
                // btnDownloadSSRStore.Visible = true;

            }

            



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
        private void WriteToExcel(DataTable dtStock, Excel1.Worksheet myExcelWorksheet ,bool type)
        {
            object misValue = System.Reflection.Missing.Value;

            if (type == false)
            {

                myExcelWorksheet.get_Range("A1", misValue).Formula = "ID";

                myExcelWorksheet.get_Range("B1", misValue).Formula = "BatchID";

                myExcelWorksheet.get_Range("C1", misValue).Formula = "BatchDate";

                myExcelWorksheet.get_Range("D1", misValue).Formula = "ProductGroup";

                myExcelWorksheet.get_Range("E1", misValue).Formula = "SoldQtyLast7Days";

                myExcelWorksheet.get_Range("F1", misValue).Formula = "ClosingQty";

                myExcelWorksheet.get_Range("G1", misValue).Formula = "Cover";

                myExcelWorksheet.get_Range("H1", misValue).Formula = "StoreCode";

                for (int i = 0, j = 2; i < dtStock.Rows.Count; i++, j++)
                {
                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["ID"]) ? dtStock.Rows[i]["ID"].ToString() : "0";

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["BatchID"]) ? dtStock.Rows[i]["BatchID"].ToString() : "0";

                    myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["BatchDate"]) ? dtStock.Rows[i]["BatchDate"].ToString() : "0";

                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroup"]) ? dtStock.Rows[i]["ProductGroup"].ToString() : "0";

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQtyLast7Days"]) ? dtStock.Rows[i]["SoldQtyLast7Days"].ToString() : "0";

                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingQty"]) ? dtStock.Rows[i]["ClosingQty"].ToString() : "0";

                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["StoreCode"]) ? dtStock.Rows[i]["StoreCode"].ToString() : "0";

                }
            }
            else
            {
                myExcelWorksheet.get_Range("A1", misValue).Formula = "ID";

                myExcelWorksheet.get_Range("B1", misValue).Formula = "BatchID";

                myExcelWorksheet.get_Range("C1", misValue).Formula = "BatchDate";

                myExcelWorksheet.get_Range("D1", misValue).Formula = "ProductGroup";

                myExcelWorksheet.get_Range("E1", misValue).Formula = "LineCode7";

                myExcelWorksheet.get_Range("F1", misValue).Formula = "SeasonCode";

                myExcelWorksheet.get_Range("G1", misValue).Formula = "LineCode7Qty";		

                myExcelWorksheet.get_Range("H1", misValue).Formula = "StoreCode";

                myExcelWorksheet.get_Range("I1", misValue).Formula = "SoldQtyLast7Days";

                myExcelWorksheet.get_Range("J1", misValue).Formula = "DeptCode";


                for (int i = 0, j = 2; i < dtStock.Rows.Count; i++, j++)
                {
                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["ID"]) ? dtStock.Rows[i]["ID"].ToString() : "0";

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["BatchID"]) ? dtStock.Rows[i]["BatchID"].ToString() : "0";

                    myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["BatchDate"]) ? dtStock.Rows[i]["BatchDate"].ToString() : "0";

                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroup"]) ? dtStock.Rows[i]["ProductGroup"].ToString() : "0";

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["LineCode7"]) ? dtStock.Rows[i]["LineCode7"].ToString() : "0";

                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["SeasonCode"]) ? dtStock.Rows[i]["SeasonCode"].ToString() : "0";

                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["LineCode7Qty"]) ? dtStock.Rows[i]["LineCode7Qty"].ToString() : "0";

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["StoreCode"]) ? dtStock.Rows[i]["StoreCode"].ToString() : "0";

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQtyLast7Days"]) ? dtStock.Rows[i]["SoldQtyLast7Days"].ToString() : "0";

                    myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["DeptCode"]) ? dtStock.Rows[i]["DeptCode"].ToString() : "0";

                }

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