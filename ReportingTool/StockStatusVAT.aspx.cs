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
using System.Diagnostics;
using System.Threading;
using System.IO;

#endregion NameSpace

namespace ReportingTool
{
    public partial class StockStatusVAT : System.Web.UI.Page
    {

        public DataTable dtStock = null;
        public const int StockStatusProcessId = 2;

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
             GenerateStockStatusReport();

            GenerateReportLCP();
        }
        #endregion btnGenerate_Click

        #endregion Events


        #region Methods

        #region GenerateStockStatusReport
        /// <summary>
        /// Generate Stock Status
        /// </summary>
        private void GenerateStockStatusReport()
        {
            //if(rdlStockStatus.SelectedItem.Value=="1")
            //   InsertStockStatus();

            GenerateReport();
        }
        #endregion GenerateStockStatusReport


        #region InsertStockStatus
        /// <summary>
        /// InsertStockStatus
        /// </summary>
        private void InsertStockStatus()
        {
            // tdLocation.Visible = false;


            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            ObjStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);



            ObjStock.SSOperationType = false;
            ObjStock.SSReportOperationType = false;
            ObjStock.SSWeeklyOperationType = false;

            ObjStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
            ObjStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            ObjStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            ObjStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());

            ObjStock.KsaRate = Convert.ToDecimal(txtKSARate.Text.Trim());

            bool Result = ObjStock.InsertStockStatusVAT();

            //Timer1.Enabled = false;

            /*if(Result==true)
            {
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Successfuly Completed.";
                

                tdLocation.Visible = true;
               
            }
            else
            {
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Stock Status Report Generaion Failed.";
              
                tdLocation.Visible = false;
                
            }*/

        }

        #endregion InsertStockStatus

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

            
            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatusVAT.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            
            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatusVAT.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];

                dtStock = objStock.GetStockStatusReportVAT(location);

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;

                    xlSheet.get_Range("A1", misValue).Formula = "Store No: " + location + " Stock Status As Of ";
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

                objStock.IntType=2;
                DataTable dtStore = objStock.GetStoreDetails();

                Excel1.Worksheet xlSht = myExcelWorkbook.Sheets[1];

                for (int i = 0; i < dtStore.Rows.Count; i++)
                {
                    xlSht.Copy(Type.Missing, myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count]); // copy
                    myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count].Name = "NEW SHEET";        // rename

                    xlSht = myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count];
                    string Location = dtStore.Rows[i]["LocationCode"].ToString();

                    xlSht.Name = Location;
                    objStock.Location = Location;
                    dtStock = objStock.GetStockStatusReportVAT(Location);
                    WriteToExcel(dtStock, xlSht, Location);
                }
                xlSht = myExcelWorkbook.Sheets[1];
                xlSht.Visible = 0;


                //Excel1.Worksheet xlSheetSummeryJordan = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                //xlSheetSummeryJordan.Name = "Jordan Summary";
                //dtStock = objStock.GetAllStockValues("JORDAN");
                //WriteToExcel(dtStock, xlSheetSummeryJordan, "JORDAN");

                //Excel1.Worksheet xlSheetSummeryUAE = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                //xlSheetSummeryUAE.Name = "UAE Summary";
                //dtStock = objStock.GetAllStockValues("UAE");
                //WriteToExcel(dtStock, xlSheetSummeryUAE, "UAE");

                //Excel1.Worksheet xlSheetSummeryOman = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                //xlSheetSummeryOman.Name = "Oman Summary";
                //dtStock = objStock.GetAllStockValues("OMAN");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheetSummeryOman, "OMAN");

                //Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                //xlSheet0408.Name = "DC";
                //dtStock = objStock.GetAllStockValues("0408");
                //WriteToExcel(dtStock, xlSheet0408, "0408");

                ////Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                //Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                //xlSheet0400.Name = "0400";
                //dtStock = objStock.GetAllStockValues("0400");
                //WriteToExcel(dtStock, xlSheet0400, "0400");

                //Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                //xlSheet0401.Name = "0401";
                //dtStock = objStock.GetAllStockValues("0401");
                //WriteToExcel(dtStock, xlSheet0401, "0401");

                //Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                //xlSheet0402.Name = "0402";
                //dtStock = objStock.GetAllStockValues("0402");
                //WriteToExcel(dtStock, xlSheet0402, "0402");

                //Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                //xlSheet0403.Name = "0403";
                //dtStock = objStock.GetAllStockValues("0403");
                //WriteToExcel(dtStock, xlSheet0403, "0403");

                //Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                //xlSheet0404.Name = "0404";
                //dtStock = objStock.GetAllStockValues("0404");
                //WriteToExcel(dtStock, xlSheet0404, "0404");

                //Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                //xlSheet0405.Name = "0405";
                //dtStock = objStock.GetAllStockValues("0405");
                //WriteToExcel(dtStock, xlSheet0405, "0405");

                //Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                //xlSheet0406.Name = "0406";
                //dtStock = objStock.GetAllStockValues("0406");
                //WriteToExcel(dtStock, xlSheet0406, "0406");

                //Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                //xlSheet0407.Name = "0407";
                //dtStock = objStock.GetAllStockValues("0407");
                //WriteToExcel(dtStock, xlSheet0407, "0407");



                //Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                //xlSheet0409.Name = "0409";
                //dtStock = objStock.GetAllStockValues("0409");
                //WriteToExcel(dtStock, xlSheet0409, "0409");

                //Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                //xlSheet0410.Name = "0410";
                //dtStock = objStock.GetAllStockValues("0410");
                //WriteToExcel(dtStock, xlSheet0410, "0410");

                //Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                //xlSheet0411.Name = "0411";
                //dtStock = objStock.GetAllStockValues("0411");
                //WriteToExcel(dtStock, xlSheet0411, "0411");

                //Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                //xlSheet0412.Name = "0412";
                //dtStock = objStock.GetAllStockValues("0412");
                //if (dtStock.Rows.Count > 0)
                //{
                //    WriteToExcel(dtStock, xlSheet0412, "0412");
                //}

                //Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                //xlSheet0414.Name = "0414";
                //dtStock = objStock.GetAllStockValues("0414");
                //WriteToExcel(dtStock, xlSheet0414, "0414");

                //Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                //xlSheet0415.Name = "0415";
                //dtStock = objStock.GetAllStockValues("0415");
                //WriteToExcel(dtStock, xlSheet0415, "0415");


                //Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                //xlSheet0416.Name = "0416";
                //dtStock = objStock.GetAllStockValues("0416");
                //WriteToExcel(dtStock, xlSheet0416, "0416");

                //Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                //xlSheet0417.Name = "0417";
                //dtStock = objStock.GetAllStockValues("0417");
                //WriteToExcel(dtStock, xlSheet0417, "0417");

                //Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                //xlSheet0418.Name = "0418";
                //dtStock = objStock.GetAllStockValues("0418");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0418, "0418");


                //Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                //xlSheet0419.Name = "0419";
                //dtStock = objStock.GetAllStockValues("0419");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0419, "0419");

                //Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                //xlSheet0421.Name = "0421";
                //dtStock = objStock.GetAllStockValues("0421");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0421, "0421");

                //Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[26];
                //xlSheet0422.Name = "0422";
                //dtStock = objStock.GetAllStockValues("0422");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0422, "0422");

                //Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[27];
                //xlSheet0423.Name = "0423";
                //dtStock = objStock.GetAllStockValues("0423");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0423, "0423");

                //Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[28];
                //xlSheet0424.Name = "0424";
                //dtStock = objStock.GetAllStockValues("0424");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0424, "0424");

                //Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[29];
                //xlSheet0425.Name = "0425";
                //dtStock = objStock.GetAllStockValues("0425");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0425, "0425");

                //Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[30];
                //xlSheet0426.Name = "0426";
                //dtStock = objStock.GetAllStockValues("0426");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcel(dtStock, xlSheet0426, "0426");


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownload.Visible = true;
                // btnDownloadSSRStore.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\SSR_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileName"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            //if (txtLocation.Text.Trim().Length > 0)
            //{
            //    DeleteColumnSSR(myExcelWorkbook, false);
            //}
            //else
            //{
            //    DeleteColumnSSR(myExcelWorkbook, true);
            //}


            //string filePathSSR1 = Server.MapPath(".") + "\\Reports\\SSR_Store" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            //ViewState["FileNameSSR1"] = filePathSSR1;
            //myExcelWorkbook.SaveAs(@filePathSSR1);

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

            string Heading = myExcelWorksheet.get_Range("A1", misValue).Formula;
            Heading = Heading + " " + Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("A1", misValue).Formula = Heading.ToString();

            string weeklyHeading = Convert.ToDateTime(dtStock.Rows[0]["FromDate"]).ToString("MMMM dd, yyyy") + " To " + Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("K3", misValue).Formula = weeklyHeading.ToString();

            string TotalsHeading = myExcelWorksheet.get_Range("AB3", misValue).Formula;
            TotalsHeading = TotalsHeading + " " + Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("AB3", misValue).Formula = TotalsHeading.ToString();

            int flag = 0;
            for (int i = 0, j = 0; i < dtStock.Rows.Count; i++, j++)
            {


                if (dtStock.Rows[i]["ReportLevel"].ToString() == "LD" && flag != 1)
                {
                    flag = 1;

                    Excel1.Range RngToCopy = myExcelWorksheet.get_Range("A2", "AL3").EntireRow;
                    Excel1.Range RngToInsert = myExcelWorksheet.get_Range("A" + (j + 6), Type.Missing).EntireRow;
                    RngToInsert.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy(Type.Missing));


                    myExcelWorksheet.get_Range("A" + (j + 5), misValue).Formula = "Divisionwise";
                    myExcelWorksheet.get_Range("A" + (j + 5), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("A" + (j + 5), misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    j += 3;
                    i -= 1;
                }

                else if (dtStock.Rows[i]["ReportLevel"].ToString() == "LSD" && flag != 2)
                {
                    flag = 2;

                    Excel1.Range RngToCopy = myExcelWorksheet.get_Range("A2", "AL3").EntireRow;
                    Excel1.Range RngToInsert = myExcelWorksheet.get_Range("A" + (j + 6), Type.Missing).EntireRow;
                    RngToInsert.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy(Type.Missing));

                    myExcelWorksheet.get_Range("A" + (j + 5), misValue).Formula = "Season Divisionwise";
                    myExcelWorksheet.get_Range("A" + (j + 5), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("A" + (j + 5), misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    j += 3;
                    i -= 1;
                }
                else if (dtStock.Rows[i]["ReportLevel"].ToString() == "Total")
                {
                    myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "Grand Total";
                    myExcelWorksheet.get_Range("A" + (j + 4), misValue).EntireRow.Font.Bold = true;
                    myExcelWorksheet.get_Range("A" + (j + 4), misValue).EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("A" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                 
                    BorderAround(myExcelWorksheet.get_Range("A" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SellThru%"]) ? dtStock.Rows[i]["SellThru%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Cls"]) ? dtStock.Rows[i]["Total Cls"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total GRN"]) ? dtStock.Rows[i]["Total GRN"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("F" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Intake Margin%"]) ? dtStock.Rows[i]["Intake Margin%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Disc % Existing"]) ? dtStock.Rows[i]["Disc % Existing"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("H" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin % Existing"]) ? dtStock.Rows[i]["Margin % Existing"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["GRN Contri%"]) ? dtStock.Rows[i]["GRN Contri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Cls Contri%"]) ? dtStock.Rows[i]["Cls Contri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Qty(Week)"]) ? dtStock.Rows[i]["Sold Qty(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("L" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Qty Contri%(Week)"]) ? dtStock.Rows[i]["Sold Qty Contri%(Week)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("L" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Cost Value(Week)"]) ? dtStock.Rows[i]["Cost Value(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("N" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Retail Value(Week)"]) ? dtStock.Rows[i]["Retail Value(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Value(Week)"]) ? dtStock.Rows[i]["Sold Value(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("P" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["VAT(Week)"]) ? dtStock.Rows[i]["VAT(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NetSoldValue(Week)"]) ? dtStock.Rows[i]["NetSoldValue(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("R" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin(Week)"]) ? dtStock.Rows[i]["Margin(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin%(Week)"]) ? dtStock.Rows[i]["Earned Margin%(Week)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("T" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Qty"]) ? dtStock.Rows[i]["WC@Qty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Cost"]) ? dtStock.Rows[i]["WC@Cost"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("U" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("V" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Retail"]) ? dtStock.Rows[i]["WC@Retail"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@NetSalesValue"]) ? dtStock.Rows[i]["WC@NetSalesValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgCP@ClsQty"]) ? dtStock.Rows[i]["AvgCP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgRP@ClsQty"]) ? dtStock.Rows[i]["AvgRP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Y" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Z" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgSP@ClsQty"]) ? dtStock.Rows[i]["AvgSP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Z" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AA" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CV@ClsQty"]) ? dtStock.Rows[i]["CV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AA" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AB" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RV@ClsQty"]) ? dtStock.Rows[i]["RV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AB" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AC" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NRV@ClsQty"]) ? dtStock.Rows[i]["NRV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AC" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AD" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SV@ClsQty"]) ? dtStock.Rows[i]["SV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AD" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AE" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NSV@ClsQty"]) ? dtStock.Rows[i]["NSV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AE" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AF" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AF" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AG" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldQtyContri%"]) ? dtStock.Rows[i]["SoldQtyContri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AG" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AH" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AH" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AI" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RetailValue"]) ? dtStock.Rows[i]["RetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AI" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AJ" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["VAT(RV)"]) ? dtStock.Rows[i]["VAT(RV)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AJ" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AK" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NetRetailValue"]) ? dtStock.Rows[i]["NetRetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AK" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AL" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AL" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AM" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["VAT(SV)"]) ? dtStock.Rows[i]["VAT(SV)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AM" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AN" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NetSoldValue"]) ? dtStock.Rows[i]["NetSoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AN" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));



                    myExcelWorksheet.get_Range("AO" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin"]) ? dtStock.Rows[i]["Total Intake Margin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AO" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AP" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin%"]) ? dtStock.Rows[i]["Total Intake Margin%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AP" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AQ" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin"]) ? dtStock.Rows[i]["Earned Margin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AQ" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AR" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin %"]) ? dtStock.Rows[i]["Earned Margin %"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AR" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AS" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance"]) ? dtStock.Rows[i]["Variance"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AS" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AT" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance Over Intake %"]) ? dtStock.Rows[i]["Variance Over Intake %"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AT" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                }

                else
                {

                    string SeasonCode = (null != dtStock.Rows[i]["Season"]) ? dtStock.Rows[i]["Season"].ToString() : "0";


                    switch (SeasonCode)
                    {
                        case "C":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "CHILDRENSWEAR";
                            break;
                        case "F":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "FOOTWEAR AND ACCESSORIES";
                            break;
                        case "H":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "HOMEWARE";
                            break;
                        case "L":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "LADIESWEAR";
                            break;
                        case "M":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "MENSWEAR";
                            break;
                        case "P":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "PROMOTIONAL";
                            break;
                        case "S":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "OWN BRAND SPORTS";
                            break;
                        case "Z":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "OTHERS";
                            break;
                        case "R":
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "SPORTS";
                            break;

                        default:
                            myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = SeasonCode;
                            break;
                    }



                    //myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Season"]) ? dtStock.Rows[i]["Season"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SellThru%"]) ? dtStock.Rows[i]["SellThru%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Cls"]) ? dtStock.Rows[i]["Total Cls"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total GRN"]) ? dtStock.Rows[i]["Total GRN"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("F" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Intake Margin%"]) ? dtStock.Rows[i]["Intake Margin%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Disc % Existing"]) ? dtStock.Rows[i]["Disc % Existing"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("H" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin % Existing"]) ? dtStock.Rows[i]["Margin % Existing"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["GRN Contri%"]) ? dtStock.Rows[i]["GRN Contri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Cls Contri%"]) ? dtStock.Rows[i]["Cls Contri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Qty(Week)"]) ? dtStock.Rows[i]["Sold Qty(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("L" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Qty Contri%(Week)"]) ? dtStock.Rows[i]["Sold Qty Contri%(Week)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("L" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Cost Value(Week)"]) ? dtStock.Rows[i]["Cost Value(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("N" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Retail Value(Week)"]) ? dtStock.Rows[i]["Retail Value(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Value(Week)"]) ? dtStock.Rows[i]["Sold Value(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                                        
                    myExcelWorksheet.get_Range("P" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["VAT(Week)"]) ? dtStock.Rows[i]["VAT(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NetSoldValue(Week)"]) ? dtStock.Rows[i]["NetSoldValue(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    

                    myExcelWorksheet.get_Range("R" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin(Week)"]) ? dtStock.Rows[i]["Margin(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin%(Week)"]) ? dtStock.Rows[i]["Earned Margin%(Week)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("T" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Qty"]) ? dtStock.Rows[i]["WC@Qty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Cost"]) ? dtStock.Rows[i]["WC@Cost"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("U" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("V" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Retail"]) ? dtStock.Rows[i]["WC@Retail"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@NetSalesValue"]) ? dtStock.Rows[i]["WC@NetSalesValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgCP@ClsQty"]) ? dtStock.Rows[i]["AvgCP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgRP@ClsQty"]) ? dtStock.Rows[i]["AvgRP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Y" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Z" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgSP@ClsQty"]) ? dtStock.Rows[i]["AvgSP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Z" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AA" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CV@ClsQty"]) ? dtStock.Rows[i]["CV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AA" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AB" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RV@ClsQty"]) ? dtStock.Rows[i]["RV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AB" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AC" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NRV@ClsQty"]) ? dtStock.Rows[i]["NRV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AC" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AD" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SV@ClsQty"]) ? dtStock.Rows[i]["SV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AD" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("AE" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NSV@ClsQty"]) ? dtStock.Rows[i]["NSV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AE" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("AF" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AF" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AG" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldQtyContri%"]) ? dtStock.Rows[i]["SoldQtyContri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AG" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AH" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AH" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AI" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RetailValue"]) ? dtStock.Rows[i]["RetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AI" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AJ" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["VAT(RV)"]) ? dtStock.Rows[i]["VAT(RV)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AJ" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AK" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NetRetailValue"]) ? dtStock.Rows[i]["NetRetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AK" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AL" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AL" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AM" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["VAT(SV)"]) ? dtStock.Rows[i]["VAT(SV)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AM" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AN" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["NetSoldValue"]) ? dtStock.Rows[i]["NetSoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AN" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));



                    myExcelWorksheet.get_Range("AO" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin"]) ? dtStock.Rows[i]["Total Intake Margin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AO" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AP" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin%"]) ? dtStock.Rows[i]["Total Intake Margin%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AP" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AQ" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin"]) ? dtStock.Rows[i]["Earned Margin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AQ" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AR" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin %"]) ? dtStock.Rows[i]["Earned Margin %"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AR" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AS" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance"]) ? dtStock.Rows[i]["Variance"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AS" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AT" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance Over Intake %"]) ? dtStock.Rows[i]["Variance Over Intake %"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AT" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }

            }

        }

        #endregion WriteToExcel


        #region GenerateReportLCP
        /// <summary>
        /// To generate excel report for stock values LCP
        /// </summary>
        private void GenerateReportLCP()
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

            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatusVATLCP.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatusVATLCP.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];

                dtStock = objStock.GetAllStockValuesLCP(location);

                if (dtStock.Rows.Count > 0)
                {

                    xlSheet.Name = location;
                    xlSheet.get_Range("A1", misValue).Formula = "Store No: " + location + " Stock Status As Of ";
                    WriteToExcelLCP(dtStock, xlSheet, location);

                    //lblMessage.Visible = true;
                    //lblMessage.ForeColor = System.Drawing.Color.Green;
                    //lblMessage.Text = "Report Generation Complete";

                    btnDownloadLCP.Visible = true;
                    //btnDownloadLCPStore.Visible = true;
                }
                else
                {
                    //lblMessage.Visible = true;
                    //lblMessage.ForeColor = System.Drawing.Color.Red;
                    //lblMessage.Text = "No Data Found";
                }

            }

            else
            {


                //Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;
                                
                //Excel1.Worksheet xlSheetSummery = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                //xlSheetSummery.Name = "Summary";
                //dtStock = objStock.GetAllStockValuesLCP("Summery");
                //WriteToExcelLCP(dtStock, xlSheetSummery, "Summery");


                objStock.IntType = 2;
                DataTable dtStore = objStock.GetStoreDetails();

                Excel1.Worksheet xlSht = myExcelWorkbook.Sheets[1];

                for (int i = 0; i < dtStore.Rows.Count; i++)
                {
                    xlSht.Copy(Type.Missing, myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count]); // copy
                    myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count].Name = "NEW SHEET";        // rename

                    xlSht = myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count];
                    string Location = dtStore.Rows[i]["LocationCode"].ToString();

                    xlSht.Name = Location;
                    objStock.Location = Location;
                    dtStock = objStock.GetAllStockValuesLCPVAT(Location);
                    WriteToExcelLCP(dtStock, xlSht, Location);
                }
                xlSht = myExcelWorkbook.Sheets[1];
                xlSht.Visible = 0;





                //Excel1.Worksheet xlSheetJor = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                //xlSheetJor.Name = "Jordan Summary";
                //dtStock = objStock.GetAllStockValuesLCP("JORDAN");
                //WriteToExcelLCP(dtStock, xlSheetJor, "JORDAN");


                //Excel1.Worksheet xlSheetUae = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                //xlSheetUae.Name = "UAE Summary";
                //dtStock = objStock.GetAllStockValuesLCP("UAE");
                //WriteToExcelLCP(dtStock, xlSheetUae, "UAE");

                //Excel1.Worksheet xlSheetOman = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                //xlSheetOman.Name = "Oman Summary";

                //dtStock = objStock.GetAllStockValuesLCP("OMAN");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheetOman, "OMAN");

                //Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                //xlSheet0408.Name = "DC";
                //dtStock = objStock.GetAllStockValuesLCP("0408");
                //WriteToExcelLCP(dtStock, xlSheet0408, "0408");

                //Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                //xlSheet0400.Name = "0400";
                //dtStock = objStock.GetAllStockValuesLCP("0400");
                //WriteToExcelLCP(dtStock, xlSheet0400, "0400");

                //Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                //xlSheet0401.Name = "0401";
                //dtStock = objStock.GetAllStockValuesLCP("0401");
                //WriteToExcelLCP(dtStock, xlSheet0401, "0401");

                //Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                //xlSheet0402.Name = "0402";
                //dtStock = objStock.GetAllStockValuesLCP("0402");
                //WriteToExcelLCP(dtStock, xlSheet0402, "0402");

                //Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                //xlSheet0403.Name = "0403";
                //dtStock = objStock.GetAllStockValuesLCP("0403");
                //WriteToExcelLCP(dtStock, xlSheet0403, "0403");

                //Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                //xlSheet0404.Name = "0404";
                //dtStock = objStock.GetAllStockValuesLCP("0404");
                //WriteToExcelLCP(dtStock, xlSheet0404, "0404");

                //Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                //xlSheet0405.Name = "0405";
                //dtStock = objStock.GetAllStockValuesLCP("0405");
                //WriteToExcelLCP(dtStock, xlSheet0405, "0405");

                //Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                //xlSheet0406.Name = "0406";
                //dtStock = objStock.GetAllStockValuesLCP("0406");
                //WriteToExcelLCP(dtStock, xlSheet0406, "0406");

                //Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                //xlSheet0407.Name = "0407";
                //dtStock = objStock.GetAllStockValuesLCP("0407");
                //WriteToExcelLCP(dtStock, xlSheet0407, "0407");


                //Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                //xlSheet0409.Name = "0409";
                //dtStock = objStock.GetAllStockValuesLCP("0409");
                //WriteToExcelLCP(dtStock, xlSheet0409, "0409");

                //Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                //xlSheet0410.Name = "0410";
                //dtStock = objStock.GetAllStockValuesLCP("0410");
                //WriteToExcelLCP(dtStock, xlSheet0410, "0410");

                //Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                //xlSheet0411.Name = "0411";
                //dtStock = objStock.GetAllStockValuesLCP("0411");
                //WriteToExcelLCP(dtStock, xlSheet0411, "0411");

                //Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                //xlSheet0412.Name = "0412";
                //dtStock = objStock.GetAllStockValuesLCP("0412");
                //if (dtStock.Rows.Count > 0)
                //{
                //    WriteToExcelLCP(dtStock, xlSheet0412, "0412");
                //}

                //Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                //xlSheet0414.Name = "0414";
                //dtStock = objStock.GetAllStockValuesLCP("0414");
                //WriteToExcelLCP(dtStock, xlSheet0414, "0414");

                //Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                //xlSheet0415.Name = "0415";
                //dtStock = objStock.GetAllStockValuesLCP("0415");
                //WriteToExcelLCP(dtStock, xlSheet0415, "0415");


                //Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                //xlSheet0416.Name = "0416";
                //dtStock = objStock.GetAllStockValuesLCP("0416");
                //WriteToExcelLCP(dtStock, xlSheet0416, "0416");

                //Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                //xlSheet0417.Name = "0417";
                //dtStock = objStock.GetAllStockValuesLCP("0417");
                //WriteToExcelLCP(dtStock, xlSheet0417, "0417");

                //Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                //xlSheet0418.Name = "0418";
                //dtStock = objStock.GetAllStockValuesLCP("0418");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0418, "0418");

                //Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                //xlSheet0419.Name = "0419";
                //dtStock = objStock.GetAllStockValuesLCP("0419");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0419, "0419");

                //Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                //xlSheet0421.Name = "0421";
                //dtStock = objStock.GetAllStockValuesLCP("0421");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0421, "0421");

                //Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[26];
                //xlSheet0422.Name = "0422";
                //dtStock = objStock.GetAllStockValuesLCP("0422");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0422, "0422");


                //Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[27];
                //xlSheet0423.Name = "0423";
                //dtStock = objStock.GetAllStockValuesLCP("0423");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0423, "0423");

                //Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[28];
                //xlSheet0424.Name = "0424";
                //dtStock = objStock.GetAllStockValuesLCP("0424");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0424, "0424");

                //Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[29];
                //xlSheet0425.Name = "0425";
                //dtStock = objStock.GetAllStockValuesLCP("0425");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0425, "0425");

                //Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[30];
                //xlSheet0426.Name = "0426";
                //dtStock = objStock.GetAllStockValuesLCP("0426");
                //if (dtStock.Rows.Count > 0)
                //    WriteToExcelLCP(dtStock, xlSheet0426, "0426");


                //objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                //objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                //dtStock = objStock.GetStockStatusLCPSummery();
                //WriteToExcelLCPSummery(dtStock, xlSheetSummery);

                //lblMessage.Visible = true;
                //lblMessage.ForeColor = System.Drawing.Color.Green;
                //lblMessage.Text = "Report Generation Complete";
                btnDownloadLCP.Visible = true;
                // btnDownloadLCPStore.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\SSR_LCP_VAT" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameLCP"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);


            //if (txtLocation.Text.Trim().Length > 0)
            //{
            //    DeleteColumnLCP(myExcelWorkbook, false);
            //}
            //else
            //{
            //    DeleteColumnLCP(myExcelWorkbook, true);
            //}


            //string filePathLCP1 = Server.MapPath(".") + "\\Reports\\SSR_LCP_Store" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            //ViewState["FileNameLCP1"] = filePathLCP1;
            //myExcelWorkbook.SaveAs(@filePathLCP1);


            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();
            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateReportLCP

        #region WriteToExcelLCP
        /// <summary>
        /// Write To Excel LCP
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelLCP(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string location)
        {
            object misValue = System.Reflection.Missing.Value;
                        

            string Heading = myExcelWorksheet.get_Range("A1", misValue).Formula;
            Heading = Heading + " From " + Convert.ToDateTime(dtStock.Rows[0]["FromDate"]).ToString("MMMM dd, yyyy") + " To " + Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("A1", misValue).Formula = Heading.ToString();


            for (int i = 0, j = 0; i < dtStock.Rows.Count; i++, j++)
            {

                if (dtStock.Rows[i]["ReportLevel"].ToString() == "Total")
                {
                    myExcelWorksheet.get_Range("A" + (j + 3), misValue).Formula = "Total";
                    myExcelWorksheet.get_Range("A" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("A" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("A" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("B" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["TotalClsQty"]) ? dtStock.Rows[i]["TotalClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("D" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("D" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("D" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["TotalSoldQty"]) ? dtStock.Rows[i]["TotalSoldQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("E" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("E" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("E" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("F" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["WeeklySoldQty"]) ? dtStock.Rows[i]["WeeklySoldQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("F" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("F" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("F" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SellThru(%)"]) ? dtStock.Rows[i]["SellThru(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("G" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("G" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("G" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("H" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldAvg/Day"]) ? dtStock.Rows[i]["SoldAvg/Day"].ToString() : "0";
                    myExcelWorksheet.get_Range("H" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("H" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("H" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["WeekCover"]) ? dtStock.Rows[i]["WeekCover"].ToString() : "0";
                    myExcelWorksheet.get_Range("I" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("I" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("I" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["TotalSoldQty"]) ? dtStock.Rows[i]["TotalSoldQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("J" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("J" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("J" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldQtyContribution(%)"]) ? dtStock.Rows[i]["SoldQtyContribution(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("K" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("K" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("K" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("L" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("L" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("L" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("L" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["RetailValue"]) ? dtStock.Rows[i]["RetailValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("M" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("M" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("M" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["VAT(RV)"]) ? dtStock.Rows[i]["VAT(RV)"].ToString() : "0";
                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["NetRetailValue"]) ? dtStock.Rows[i]["NetRetailValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["VAT(SV)"]) ? dtStock.Rows[i]["VAT(SV)"].ToString() : "0";
                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["NetSoldValue"]) ? dtStock.Rows[i]["NetSoldValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin"]) ? dtStock.Rows[i]["IntakeMargin"].ToString() : "0";
                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin(%)"]) ? dtStock.Rows[i]["IntakeMargin(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin"]) ? dtStock.Rows[i]["EarnedMargin"].ToString() : "0";
                    myExcelWorksheet.get_Range("U" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("U" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("U" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("V" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin(%)"]) ? dtStock.Rows[i]["EarnedMargin(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("V" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("V" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("V" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(Intake Vs Earned)"]) ? dtStock.Rows[i]["Variance(Intake Vs Earned)"].ToString() : "0";
                    myExcelWorksheet.get_Range("W" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("W" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(%)"]) ? dtStock.Rows[i]["Variance(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("X" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("X" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }

                else
                {

                    myExcelWorksheet.get_Range("A" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["CategoryCode"]) ? dtStock.Rows[i]["CategoryCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["ProductGroupCode"]) ? dtStock.Rows[i]["ProductGroupCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["ProductDescription"]) ? dtStock.Rows[i]["ProductDescription"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["TotalClsQty"]) ? dtStock.Rows[i]["TotalClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["TotalSoldQty"]) ? dtStock.Rows[i]["TotalSoldQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("F" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["WeeklySoldQty"]) ? dtStock.Rows[i]["WeeklySoldQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SellThru(%)"]) ? dtStock.Rows[i]["SellThru(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("H" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldAvg/Day"]) ? dtStock.Rows[i]["SoldAvg/Day"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["WeekCover"]) ? dtStock.Rows[i]["WeekCover"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["TotalSoldQty"]) ? dtStock.Rows[i]["TotalSoldQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldQtyContribution(%)"]) ? dtStock.Rows[i]["SoldQtyContribution(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("L" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("L" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["RetailValue"]) ? dtStock.Rows[i]["RetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["VAT(RV)"]) ? dtStock.Rows[i]["VAT(RV)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["NetRetailValue"]) ? dtStock.Rows[i]["NetRetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["VAT(SV)"]) ? dtStock.Rows[i]["VAT(SV)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["NetSoldValue"]) ? dtStock.Rows[i]["NetSoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin"]) ? dtStock.Rows[i]["IntakeMargin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin(%)"]) ? dtStock.Rows[i]["IntakeMargin(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin"]) ? dtStock.Rows[i]["EarnedMargin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("U" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("V" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin(%)"]) ? dtStock.Rows[i]["EarnedMargin(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("V" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(Intake Vs Earned)"]) ? dtStock.Rows[i]["Variance(Intake Vs Earned)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(%)"]) ? dtStock.Rows[i]["Variance(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                }

            }
        }

        #endregion WriteToExcelLCP

        #region WriteToExcelLCPSummery
        /// <summary>
        /// Write To Excel LCP Summery
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>

        private void WriteToExcelLCPSummery(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
                      
            for (int i = 0, j = 0; i < dtStock.Rows.Count; i++, j++)
            {

                if (dtStock.Rows[i]["CategoryCode"].ToString() == "Total")
                {
                    //myExcelWorksheet.get_Range("A" + (j + 3), misValue).Formula = "Total";
                    // myExcelWorksheet.get_Range("A" + (j + 3), misValue).Font.Bold = true;


                    if (Convert.ToInt32(dtStock.Rows[i]["0400"]) != 0)
                    {
                        myExcelWorksheet.get_Range("V" + (j + 3), misValue).Formula = dtStock.Rows[i]["0400"].ToString();
                        myExcelWorksheet.get_Range("V" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("V" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("V" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0401"]) != 0)
                    {
                        myExcelWorksheet.get_Range("W" + (j + 3), misValue).Formula = dtStock.Rows[i]["0401"].ToString();
                        myExcelWorksheet.get_Range("W" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("W" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("W" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0402"]) != 0)
                    {
                        myExcelWorksheet.get_Range("X" + (j + 3), misValue).Formula = dtStock.Rows[i]["0402"].ToString();
                        myExcelWorksheet.get_Range("X" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("X" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("X" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0403"]) != 0)
                    {
                        myExcelWorksheet.get_Range("Y" + (j + 3), misValue).Formula = dtStock.Rows[i]["0403"].ToString();
                        myExcelWorksheet.get_Range("Y" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("Y" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("Y" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    }
                    else if (Convert.ToInt32(dtStock.Rows[i]["0404"]) != 0)
                    {
                        myExcelWorksheet.get_Range("Z" + (j + 3), misValue).Formula = dtStock.Rows[i]["0404"].ToString();
                        myExcelWorksheet.get_Range("Z" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("Z" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("Z" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0405"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AA" + (j + 3), misValue).Formula = dtStock.Rows[i]["0405"].ToString();
                        myExcelWorksheet.get_Range("AA" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AA" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AA" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0406"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AB" + (j + 3), misValue).Formula = dtStock.Rows[i]["0406"].ToString();
                        myExcelWorksheet.get_Range("AB" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AB" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AB" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0407"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AC" + (j + 3), misValue).Formula = dtStock.Rows[i]["0407"].ToString();
                        myExcelWorksheet.get_Range("AC" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AC" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AC" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0409"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AD" + (j + 3), misValue).Formula = dtStock.Rows[i]["0409"].ToString();
                        myExcelWorksheet.get_Range("AD" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AD" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AD" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0410"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AE" + (j + 3), misValue).Formula = dtStock.Rows[i]["0410"].ToString();
                        myExcelWorksheet.get_Range("AE" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AE" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AE" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }


                    else if (Convert.ToInt32(dtStock.Rows[i]["0411"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AF" + (j + 3), misValue).Formula = dtStock.Rows[i]["0411"].ToString();
                        myExcelWorksheet.get_Range("AF" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AF" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AF" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }


                    else if (Convert.ToInt32(dtStock.Rows[i]["0414"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AG" + (j + 3), misValue).Formula = dtStock.Rows[i]["0414"].ToString();
                        myExcelWorksheet.get_Range("AG" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AG" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AG" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0415"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AH" + (j + 3), misValue).Formula = dtStock.Rows[i]["0415"].ToString();
                        myExcelWorksheet.get_Range("AH" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AH" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AH" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0416"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AI" + (j + 3), misValue).Formula = dtStock.Rows[i]["0416"].ToString();
                        myExcelWorksheet.get_Range("AI" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AI" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AI" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0417"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AJ" + (j + 3), misValue).Formula = dtStock.Rows[i]["0417"].ToString();
                        myExcelWorksheet.get_Range("AJ" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AJ" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AJ" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0408"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AK" + (j + 3), misValue).Formula = dtStock.Rows[i]["0408"].ToString();
                        myExcelWorksheet.get_Range("AK" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AK" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AK" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0412"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AL" + (j + 3), misValue).Formula = dtStock.Rows[i]["0412"].ToString();
                        myExcelWorksheet.get_Range("AL" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AL" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AL" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0418"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AM" + (j + 3), misValue).Formula = dtStock.Rows[i]["0418"].ToString();
                        myExcelWorksheet.get_Range("AM" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AM" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AM" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0419"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AN" + (j + 3), misValue).Formula = dtStock.Rows[i]["0419"].ToString();
                        myExcelWorksheet.get_Range("AN" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AN" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AN" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0421"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AO" + (j + 3), misValue).Formula = dtStock.Rows[i]["0421"].ToString();
                        myExcelWorksheet.get_Range("AO" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AO" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AO" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0422"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AP" + (j + 3), misValue).Formula = dtStock.Rows[i]["0422"].ToString();
                        myExcelWorksheet.get_Range("AP" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AP" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AP" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }
                    else if (Convert.ToInt32(dtStock.Rows[i]["0423"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AQ" + (j + 3), misValue).Formula = dtStock.Rows[i]["0423"].ToString();
                        myExcelWorksheet.get_Range("AQ" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AQ" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AQ" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0424"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AR" + (j + 3), misValue).Formula = dtStock.Rows[i]["0424"].ToString();
                        myExcelWorksheet.get_Range("AR" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AR" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AR" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0425"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AS" + (j + 3), misValue).Formula = dtStock.Rows[i]["0425"].ToString();
                        myExcelWorksheet.get_Range("AS" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AS" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AS" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0426"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AT" + (j + 3), misValue).Formula = dtStock.Rows[i]["0426"].ToString();
                        myExcelWorksheet.get_Range("AT" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AT" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AT" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    j = j - 1;
                }

                else
                {


                    myExcelWorksheet.get_Range("V" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0400"]) ? dtStock.Rows[i]["0400"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("V" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0401"]) ? dtStock.Rows[i]["0401"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0402"]) ? dtStock.Rows[i]["0402"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0403"]) ? dtStock.Rows[i]["0403"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Y" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Z" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0404"]) ? dtStock.Rows[i]["0404"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Z" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AA" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0405"]) ? dtStock.Rows[i]["0405"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AA" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AB" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0406"]) ? dtStock.Rows[i]["0406"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AB" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AC" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0407"]) ? dtStock.Rows[i]["0407"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AC" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("AD" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0409"]) ? dtStock.Rows[i]["0409"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AD" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AE" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0410"]) ? dtStock.Rows[i]["0410"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AE" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AF" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0411"]) ? dtStock.Rows[i]["0411"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AF" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AG" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0414"]) ? dtStock.Rows[i]["0414"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AG" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AH" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0415"]) ? dtStock.Rows[i]["0415"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AH" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AI" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0416"]) ? dtStock.Rows[i]["0416"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AI" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AJ" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0417"]) ? dtStock.Rows[i]["0417"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AJ" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AK" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0408"]) ? dtStock.Rows[i]["0408"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AK" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AL" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0412"]) ? dtStock.Rows[i]["0412"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AL" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AM" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0418"]) ? dtStock.Rows[i]["0418"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AM" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AN" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0419"]) ? dtStock.Rows[i]["0419"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AN" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AO" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0421"]) ? dtStock.Rows[i]["0421"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AO" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AP" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0422"]) ? dtStock.Rows[i]["0422"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AP" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AQ" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0423"]) ? dtStock.Rows[i]["0423"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AQ" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AR" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0424"]) ? dtStock.Rows[i]["0424"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AR" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AS" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0425"]) ? dtStock.Rows[i]["0425"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AS" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AT" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0426"]) ? dtStock.Rows[i]["0426"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AT" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }

            }
        }

        #endregion WriteToExcelLCP




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


        #endregion Methods
    }
}