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

namespace Test
{
    public partial class StockStatus : System.Web.UI.Page
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
            if(!IsPostBack)
            {
             
            }
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
           
            lblMessage.Visible = false;
            btnDownload.Visible = false;
            //udpReport.Update();
            if (rdlStockStatus.SelectedValue == "1")
            {
                if (GetProcessStatus())
                {
                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "Tables Are Locked By Another User,Try Again Later";
                }
                else
                {
                    GetStockDetails objStock = new GetStockDetails();
                    objStock.ProcessStatusFlag = true;
                    objStock.ProcessStatusId = StockStatusProcessId;
                    objStock.UpdateProcessStatus();

                    InsertStockStatus();
                    InsertWssiReport();
                    InsertProductGroupCmpReport();
                    //--Not Using---InsertWssiDivisionReport();

                    InsertWssiProductGroupReport();

                    objStock.ProcessStatusFlag = false;
                    objStock.ProcessStatusId = StockStatusProcessId;
                    objStock.UpdateProcessStatus();
                }
            }
           
            ViewState["FileName"] = null;

            ViewState["FileNameLCP"] = null;

            ViewState["FileNameSSR1"] = null;
            ViewState["FileNameLCP1"] = null;

            ViewState["FileNameWSSI"] = null;
            ViewState["FileNamePgCmp"] = null;

            ViewState["FileNameWssiDivision"] = null;

            ViewState["FileNameWssiProductGroup"] = null;


            if (!GetProcessStatus())
            {
                GenerateReportLCP();

                GenerateReport();

                GenerateWSSIReport();

                GeneratePgCmpReport();

                //--Not Using---GenerateWSSIDivisionReport();

                GenerateWSSIProductReport();
            }
            //udpReport.Update();

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
                
        #region btnDownloadLCP_Click
        /// <summary>
        /// btnDownloadLCP_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadLCP_Click(object sender, EventArgs e)
        {
            FileDownloadLCP();
        }
        #endregion btnDownloadLCP_Click


        #region btnDownloadSSRStore_Click
        /// <summary>
        /// btnDownloadSSRStore_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadSSRStore_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameSSR1"].ToString();
            FileDownload(filename);

        }
        #endregion btnDownloadSSRStore_Click

        #region btnDownloadLCPStore_Click
        /// <summary>
        /// btnDownloadLCPStore_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadLCPStore_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameLCP1"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadLCPStore_Click

        #region btnDownloadWssi_Click
        /// <summary>
        /// btnDownloadWssi_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadWssi_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameWSSI"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadWssi_Click

        #region btnPGCompare_Click
        /// <summary>
        /// btnPGCompare_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnPGCompare_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNamePgCmp"].ToString();
            FileDownload(filename);
        }
        #endregion btnPGCompare_Click


        #region btnWssiDivision_Click
        /// <summary>
        /// btnWssiDivision_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnWssiDivision_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameWssiDivision"].ToString();
            FileDownload(filename);
        }
        #endregion btnWssiDivision_Click


        #region btnWssiPg_Click
        /// <summary>
        /// btnWssiPg_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnWssiPg_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameWssiProductGroup"].ToString();
            FileDownload(filename);
        }
        #endregion btnWssiPg_Click

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

                //String fileName = "C:\\book1.xlsx";
               // myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                //Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.ActiveSheet;
                
                //myExcelWorkbooks.Close();

             


                string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatusReport6.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                


                //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
                //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

                //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);
          
               

                GetStockDetails objStock = new GetStockDetails();

                //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
                if (txtLocation.Text.Trim().Length > 0)
                {
                    string location = txtLocation.Text.Trim();
                    
                    fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatusOneLocation1.xlsx";
                    myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                   
                    dtStock = objStock.GetAllStockValues(location);

                    if (dtStock.Rows.Count > 0)
                    {
                        xlSheet.Name = location;

                        xlSheet.get_Range("A1", misValue).Formula = "Store No: "+location+" Stock Status As Of ";
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


                    Excel1.Worksheet xlSheetSummery = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheetSummery.Name = "Summary";
                    dtStock = objStock.GetAllStockValues("Summery");
                    WriteToExcel(dtStock, xlSheetSummery, "Summery");
                   

                    Excel1.Worksheet xlSheetSummeryJordan = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheetSummeryJordan.Name = "Jordan Summary";
                    dtStock = objStock.GetAllStockValues("JORDAN");
                    WriteToExcel(dtStock, xlSheetSummeryJordan, "JORDAN");

                    Excel1.Worksheet xlSheetSummeryUAE = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheetSummeryUAE.Name = "UAE Summary";
                    dtStock = objStock.GetAllStockValues("UAE");
                    WriteToExcel(dtStock, xlSheetSummeryUAE, "UAE");

                    Excel1.Worksheet xlSheetSummeryOman = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheetSummeryOman.Name = "Oman Summary";
                    dtStock = objStock.GetAllStockValues("OMAN");
                    if(dtStock.Rows.Count>0)
                        WriteToExcel(dtStock, xlSheetSummeryOman, "OMAN");

                    Excel1.Worksheet xlSheetSummeryBahrain = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheetSummeryBahrain.Name = "Bahrain Summary";
                    dtStock = objStock.GetAllStockValues("BAHRAIN");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheetSummeryBahrain, "BAHRAIN");

                    Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheet0408.Name = "DC";
                    dtStock = objStock.GetAllStockValues("DC");
                    WriteToExcel(dtStock, xlSheet0408, "DC");

                    //Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                    Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                    xlSheet0400.Name = "0400";
                    dtStock = objStock.GetAllStockValues("0400");
                    WriteToExcel(dtStock, xlSheet0400, "0400");

                    Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                    xlSheet0401.Name = "0401";
                    dtStock = objStock.GetAllStockValues("0401");
                    WriteToExcel(dtStock, xlSheet0401, "0401");

                    Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                    xlSheet0402.Name = "0402";
                    dtStock = objStock.GetAllStockValues("0402");
                    WriteToExcel(dtStock, xlSheet0402, "0402");

                    Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                    xlSheet0403.Name = "0403";
                    dtStock = objStock.GetAllStockValues("0403");
                    WriteToExcel(dtStock, xlSheet0403, "0403");

                    Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                    xlSheet0404.Name = "0404";
                    dtStock = objStock.GetAllStockValues("0404");
                    WriteToExcel(dtStock, xlSheet0404, "0404");

                    Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                    xlSheet0405.Name = "0405";
                    dtStock = objStock.GetAllStockValues("0405");
                    WriteToExcel(dtStock, xlSheet0405, "0405");

                    Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                    xlSheet0406.Name = "0406";
                    dtStock = objStock.GetAllStockValues("0406");
                    WriteToExcel(dtStock, xlSheet0406, "0406");

                    Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                    xlSheet0407.Name = "0407";
                    dtStock = objStock.GetAllStockValues("0407");
                    WriteToExcel(dtStock, xlSheet0407, "0407");

                   

                    Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                    xlSheet0409.Name = "0409";
                    dtStock = objStock.GetAllStockValues("0409");
                    WriteToExcel(dtStock, xlSheet0409, "0409");

                    Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                    xlSheet0410.Name = "0410";
                    dtStock = objStock.GetAllStockValues("0410");
                    WriteToExcel(dtStock, xlSheet0410, "0410");

                    Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                    xlSheet0411.Name = "0411";
                    dtStock = objStock.GetAllStockValues("0411");
                    WriteToExcel(dtStock, xlSheet0411, "0411");

                    Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                    xlSheet0412.Name = "0412";
                    dtStock = objStock.GetAllStockValues("0412");
                    if (dtStock.Rows.Count > 0)
                    {
                        WriteToExcel(dtStock, xlSheet0412, "0412");
                    }

                    Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                    xlSheet0414.Name = "0414";
                    dtStock = objStock.GetAllStockValues("0414");
                    WriteToExcel(dtStock, xlSheet0414, "0414");

                    Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                    xlSheet0415.Name = "0415";
                    dtStock = objStock.GetAllStockValues("0415");
                    WriteToExcel(dtStock, xlSheet0415, "0415");


                    Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                    xlSheet0416.Name = "0416";
                    dtStock = objStock.GetAllStockValues("0416");
                    WriteToExcel(dtStock, xlSheet0416, "0416");

                    Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                    xlSheet0417.Name = "0417";
                    dtStock = objStock.GetAllStockValues("0417");
                    WriteToExcel(dtStock, xlSheet0417, "0417");

                    Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                    xlSheet0418.Name = "0418";
                    dtStock = objStock.GetAllStockValues("0418");
                    if(dtStock.Rows.Count>0)
                        WriteToExcel(dtStock, xlSheet0418, "0418");


                    Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                    xlSheet0419.Name = "0419";
                    dtStock = objStock.GetAllStockValues("0419");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0419, "0419");

                    Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[26];
                    xlSheet0421.Name = "0421";
                    dtStock = objStock.GetAllStockValues("0421");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0421, "0421");

                    Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[27];
                    xlSheet0422.Name = "0422";
                    dtStock = objStock.GetAllStockValues("0422");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0422, "0422");

                    Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[28];
                    xlSheet0423.Name = "0423";
                    dtStock = objStock.GetAllStockValues("0423");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0423, "0423");

                    Excel1.Worksheet xlSheet0424= (Excel1.Worksheet)myExcelWorkbook.Sheets[29];
                    xlSheet0424.Name = "0424";
                    dtStock = objStock.GetAllStockValues("0424");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0424, "0424");

                    Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[30];
                    xlSheet0425.Name = "0425";
                    dtStock = objStock.GetAllStockValues("0425");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0425, "0425");

                    Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[31];
                    xlSheet0426.Name = "0426";
                    dtStock = objStock.GetAllStockValues("0426");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0426, "0426");

                    Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[32];
                    xlSheet0427.Name = "0427";
                    dtStock = objStock.GetAllStockValues("0427");
                    if (dtStock.Rows.Count > 0)
                        WriteToExcel(dtStock, xlSheet0427, "0427");


                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";
                    btnDownload.Visible = true;
                   // btnDownloadSSRStore.Visible = true;

                }

                Random rnd = new Random();
                string filePath = Server.MapPath(".") +"\\Reports\\SSR_"+DateTime.Now.Day+"-"+DateTime.Now.Month+"-"+DateTime.Now.Year+"_"+ rnd.Next()+".xlsx";
                
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
        private void WriteToExcel(DataTable dtStock, Excel1.Worksheet myExcelWorksheet,string location)
        {
            object misValue = System.Reflection.Missing.Value;

           // myExcelWorksheet.get_Range("A2", "AL20")
            //System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(79, 129, 189)
           
            // BorderAround(myExcelWorksheet.get_Range("A2", "AL20").Cells, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            //Range RngToCopy = ws.get_Range(StartCell, EndCell).EntireRow;
            //Range RngToInsert = ws.get_Range(StartCell, Type.Missing).EntireRow;
            //oRngToInsert.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, oRngToCopy.Copy(Type.Missing));


            //if(location=="JORDAN")
            //    myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Summery Report For Jordan  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();
            //else if (location == "UAE")
            //    myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Summery Report For UAE  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();
            //else if (location == "Summery")
            //    myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Summery Report from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();
          
            //else
            //    myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Report For Store No : " + location + "  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();

            //myExcelWorksheet.get_Range("A1", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("A1", misValue).Font.Size = 14;


            //myExcelWorksheet.get_Range("A2", misValue).Formula = "Season Code";
            //myExcelWorksheet.get_Range("A2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("A2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("A2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0,51,102));

          
            //myExcelWorksheet.get_Range("B2", misValue).Formula = "Sell Thru(%)";
            //myExcelWorksheet.get_Range("B2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("B2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("B2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown);

            //myExcelWorksheet.get_Range("C2", misValue).Formula = "TotalCls";
            //myExcelWorksheet.get_Range("C2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("C2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("C2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            //myExcelWorksheet.get_Range("D2", misValue).Formula = "Total Sold";
            //myExcelWorksheet.get_Range("D2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("D2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("D2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            //myExcelWorksheet.get_Range("E2", misValue).Formula = "Total GRN";
            //myExcelWorksheet.get_Range("E2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("E2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("E2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            //myExcelWorksheet.get_Range("F2", misValue).Formula = "In Take Margin(%)";
            //myExcelWorksheet.get_Range("F2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("F2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("F2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Magenta);

            //myExcelWorksheet.get_Range("G2", misValue).Formula = "Disc % Existing";
            //myExcelWorksheet.get_Range("G2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("G2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("G2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Magenta);

            //myExcelWorksheet.get_Range("H2", misValue).Formula = "Margin % Existing";
            //myExcelWorksheet.get_Range("H2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("H2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("H2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Magenta);


            //myExcelWorksheet.get_Range("I2", misValue).Formula = "GRN Contri(%)";
            //myExcelWorksheet.get_Range("I2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("I2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("I2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            //myExcelWorksheet.get_Range("J2", misValue).Formula = "Cls Contri(%)";
            //myExcelWorksheet.get_Range("J2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("J2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("J2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);


            //myExcelWorksheet.get_Range("K2", misValue).Formula = "Sold Qty(Week)";
            //myExcelWorksheet.get_Range("K2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("K2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("K2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            //myExcelWorksheet.get_Range("L2", misValue).Formula = "Sold Qty Contri(%) (Week)";
            //myExcelWorksheet.get_Range("L2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("L2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("L2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);


            //myExcelWorksheet.get_Range("M2", misValue).Formula = "Cost Value (Week)";
            //myExcelWorksheet.get_Range("M2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("M2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("M2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);


            //myExcelWorksheet.get_Range("N2", misValue).Formula = "Retail Value (Week)";
            //myExcelWorksheet.get_Range("N2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("N2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("N2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);


            //myExcelWorksheet.get_Range("O2", misValue).Formula = "Sold Value (Week)";
            //myExcelWorksheet.get_Range("O2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("O2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("O2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);


            //myExcelWorksheet.get_Range("P2", misValue).Formula = "Margin (Week)";
            //myExcelWorksheet.get_Range("P2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("P2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("P2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            //myExcelWorksheet.get_Range("Q2", misValue).Formula = "Earned Margin% (Week)";
            //myExcelWorksheet.get_Range("Q2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("Q2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("Q2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            //myExcelWorksheet.get_Range("R2", misValue).Formula = "Wc @ Qty";
            //myExcelWorksheet.get_Range("R2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("R2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("R2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

            //myExcelWorksheet.get_Range("S2", misValue).Formula = "Wc @ Cost";
            //myExcelWorksheet.get_Range("S2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("S2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("S2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

            //myExcelWorksheet.get_Range("T2", misValue).Formula = "Wc @ Retail";
            //myExcelWorksheet.get_Range("T2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("T2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("T2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

            //myExcelWorksheet.get_Range("U2", misValue).Formula = "Wc @ Sales Value";
            //myExcelWorksheet.get_Range("U2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("U2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("U2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

            //myExcelWorksheet.get_Range("V2", misValue).Formula = "Avg CP";
            //myExcelWorksheet.get_Range("V2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("V2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("V2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown);

            //myExcelWorksheet.get_Range("W2", misValue).Formula = "Avg RP";
            //myExcelWorksheet.get_Range("W2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("W2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("W2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown);

            //myExcelWorksheet.get_Range("X2", misValue).Formula = "Avg SP";
            //myExcelWorksheet.get_Range("X2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("X2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("X2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown);

            //myExcelWorksheet.get_Range("Y2", misValue).Formula = "CV @ Cls Qty";
            //myExcelWorksheet.get_Range("Y2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("Y2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("Y2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("Z2", misValue).Formula = "RV @ Cls Qty";
            //myExcelWorksheet.get_Range("Z2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("Z2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("Z2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("AA2", misValue).Formula = "SV @ Cls Qty";
            //myExcelWorksheet.get_Range("AA2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AA2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AA2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


            //myExcelWorksheet.get_Range("AB2", misValue).Formula = "Sold Qty";
            //myExcelWorksheet.get_Range("AB2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AB2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AB2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AC2", misValue).Formula = "Sold Qty Contri %";
            //myExcelWorksheet.get_Range("AC2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AC2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AC2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AD2", misValue).Formula = "Cost Value";
            //myExcelWorksheet.get_Range("AD2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AD2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AD2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AE2", misValue).Formula = "Retail Value";
            //myExcelWorksheet.get_Range("AE2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AE2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AE2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AF2", misValue).Formula = "Sold Value";
            //myExcelWorksheet.get_Range("AF2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AF2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AF2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AG2", misValue).Formula = "Intake Margin";
            //myExcelWorksheet.get_Range("AG2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AG2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AG2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AH2", misValue).Formula = "Intake Margin%";
            //myExcelWorksheet.get_Range("AH2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AH2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AH2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AI2", misValue).Formula = "Earned Margin";
            //myExcelWorksheet.get_Range("AI2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AI2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AI2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AJ2", misValue).Formula = "Earned Margin %";
            //myExcelWorksheet.get_Range("AJ2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AJ2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AJ2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AK2", misValue).Formula = "Variance";
            //myExcelWorksheet.get_Range("AK2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AK2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AK2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //myExcelWorksheet.get_Range("AL2", misValue).Formula = "% (Var. Over Intake)";
            //myExcelWorksheet.get_Range("AL2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AL2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AL2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);


            string Heading = myExcelWorksheet.get_Range("A1", misValue).Formula;
            Heading= Heading+" "+ Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("A1", misValue).Formula = Heading.ToString();

            string weeklyHeading = Convert.ToDateTime(dtStock.Rows[0]["FromDate"]).ToString("MMMM dd, yyyy") + " To " + Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("K3", misValue).Formula = weeklyHeading.ToString();

            string TotalsHeading = myExcelWorksheet.get_Range("AB3", misValue).Formula;
            TotalsHeading = TotalsHeading +" "+Convert.ToDateTime(dtStock.Rows[0]["ToDate"]).ToString("MMMM dd, yyyy");
            myExcelWorksheet.get_Range("AB3", misValue).Formula = TotalsHeading.ToString();

            int flag = 0;
            for (int i = 0,j=0; i < dtStock.Rows.Count; i++,j++)
            {
               

                if (dtStock.Rows[i]["ReportLevel"].ToString() == "LD" && flag!=1)
                {
                    flag = 1;

                    Excel1.Range RngToCopy = myExcelWorksheet.get_Range("A2","AL3").EntireRow;
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
                    myExcelWorksheet.get_Range("A" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("A" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("A" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SellThru%"]) ? dtStock.Rows[i]["SellThru%"].ToString()+"%" : "0";
                    myExcelWorksheet.get_Range("B" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("B" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("B" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Cls"]) ? dtStock.Rows[i]["Total Cls"].ToString() : "0";
                    myExcelWorksheet.get_Range("C" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("C" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("C" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    myExcelWorksheet.get_Range("D" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("D" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("D" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total GRN"]) ? dtStock.Rows[i]["Total GRN"].ToString() : "0";
                    myExcelWorksheet.get_Range("E" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("E" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("E" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("F" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Intake Margin%"]) ? dtStock.Rows[i]["Intake Margin%"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("F" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("F" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("F" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Disc % Existing"]) ? dtStock.Rows[i]["Disc % Existing"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("G" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("G" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("G" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("H" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin % Existing"]) ? dtStock.Rows[i]["Margin % Existing"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("H" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("H" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("H" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["GRN Contri%"]) ? dtStock.Rows[i]["GRN Contri%"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("I" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("I" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("I" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Cls Contri%"]) ? dtStock.Rows[i]["Cls Contri%"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("J" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("J" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("J" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Qty(Week)"]) ? dtStock.Rows[i]["Sold Qty(Week)"].ToString() : "0";
                    myExcelWorksheet.get_Range("K" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("K" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("K" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("L" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Qty Contri%(Week)"]) ? dtStock.Rows[i]["Sold Qty Contri%(Week)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("L" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("L" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("L" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Cost Value(Week)"]) ? dtStock.Rows[i]["Cost Value(Week)"].ToString() : "0";
                    myExcelWorksheet.get_Range("M" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("M" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("M" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("N" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Retail Value(Week)"]) ? dtStock.Rows[i]["Retail Value(Week)"].ToString() : "0";
                    myExcelWorksheet.get_Range("N" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("N" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                                        
                    myExcelWorksheet.get_Range("O" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Sold Value(Week)"]) ? dtStock.Rows[i]["Sold Value(Week)"].ToString() : "0";
                    myExcelWorksheet.get_Range("O" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("O" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    
                    myExcelWorksheet.get_Range("P" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin(Week)"]) ? dtStock.Rows[i]["Margin(Week)"].ToString() : "0";
                    myExcelWorksheet.get_Range("P" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("P" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin%(Week)"]) ? dtStock.Rows[i]["Earned Margin%(Week)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("Q" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("Q" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Qty"]) ? dtStock.Rows[i]["WC@Qty"].ToString() : "0";
                    myExcelWorksheet.get_Range("R" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("R" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Cost"]) ? dtStock.Rows[i]["WC@Cost"].ToString() : "0";
                    myExcelWorksheet.get_Range("S" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("S" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("T" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Retail"]) ? dtStock.Rows[i]["WC@Retail"].ToString() : "0";
                    myExcelWorksheet.get_Range("T" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("T" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@SalesValue"]) ? dtStock.Rows[i]["WC@SalesValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("U" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("U" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("U" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("V" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgCP@ClsQty"]) ? dtStock.Rows[i]["AvgCP@ClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("V" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("V" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("V" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                                        
                    myExcelWorksheet.get_Range("W" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgRP@ClsQty"]) ? dtStock.Rows[i]["AvgRP@ClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("W" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("W" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgSP@ClsQty"]) ? dtStock.Rows[i]["AvgSP@ClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("X" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("X" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CV@ClsQty"]) ? dtStock.Rows[i]["CV@ClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("Y" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("Y" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("Y" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Z" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RV@ClsQty"]) ? dtStock.Rows[i]["RV@ClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("Z" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("Z" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("Z" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AA" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SV@ClsQty"]) ? dtStock.Rows[i]["SV@ClsQty"].ToString() : "0";
                    myExcelWorksheet.get_Range("AA" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AA" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AA" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AB" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    myExcelWorksheet.get_Range("AB" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AB" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AB" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AC" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldQtyContri%"]) ? dtStock.Rows[i]["SoldQtyContri%"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("AC" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AC" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AC" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AD" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("AD" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AD" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AD" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AE" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RetailValue"]) ? dtStock.Rows[i]["RetailValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("AE" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AE" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AE" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AF" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("AF" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AF" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AF" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AG" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin"]) ? dtStock.Rows[i]["Total Intake Margin"].ToString() : "0";
                    myExcelWorksheet.get_Range("AG" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AG" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AG" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AH" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin%"]) ? dtStock.Rows[i]["Total Intake Margin%"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("AH" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AH" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AH" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AI" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin"]) ? dtStock.Rows[i]["Earned Margin"].ToString() : "0";
                    myExcelWorksheet.get_Range("AI" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AI" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AI" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AJ" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin %"]) ? dtStock.Rows[i]["Earned Margin %"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("AJ" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AJ" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AJ" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AK" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance"]) ? dtStock.Rows[i]["Variance"].ToString() : "0";
                    myExcelWorksheet.get_Range("AK" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AK" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AK" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AL" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance Over Intake %"]) ? dtStock.Rows[i]["Variance Over Intake %"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("AL" + (j + 4), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("AL" + (j + 4), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("AL" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }

                else
                {

                    string SeasonCode=(null != dtStock.Rows[i]["Season"]) ? dtStock.Rows[i]["Season"].ToString() : "0";
                    

                    switch(SeasonCode)
                    {
                        case "C": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "CHILDRENSWEAR";
                                  break;
                        case "F": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "FOOTWEAR AND ACCESSORIES";
                                  break;
                        case "H": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "HOMEWARE";
                                  break;
                        case "L": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "LADIESWEAR";
                                  break;
                        case "M": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "MENSWEAR";
                                  break;
                        case "P": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "PROMOTIONAL";
                                  break;
                        case "S": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "OWN BRAND SPORTS";
                                  break;
                        case "Z": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "OTHERS";
                                  break;
                        case "R": myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = "SPORTS";
                                  break;
                        
                        default: myExcelWorksheet.get_Range("A" + (j + 4), misValue).Formula = SeasonCode;
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

                    myExcelWorksheet.get_Range("P" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Margin(Week)"]) ? dtStock.Rows[i]["Margin(Week)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin%(Week)"]) ? dtStock.Rows[i]["Earned Margin%(Week)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Qty"]) ? dtStock.Rows[i]["WC@Qty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Cost"]) ? dtStock.Rows[i]["WC@Cost"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("T" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@Retail"]) ? dtStock.Rows[i]["WC@Retail"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["WC@SalesValue"]) ? dtStock.Rows[i]["WC@SalesValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("U" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("V" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgCP@ClsQty"]) ? dtStock.Rows[i]["AvgCP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("V" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgRP@ClsQty"]) ? dtStock.Rows[i]["AvgRP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["AvgSP@ClsQty"]) ? dtStock.Rows[i]["AvgSP@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("X" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CV@ClsQty"]) ? dtStock.Rows[i]["CV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Y" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Z" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RV@ClsQty"]) ? dtStock.Rows[i]["RV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Z" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AA" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SV@ClsQty"]) ? dtStock.Rows[i]["SV@ClsQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AA" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AB" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Sold"]) ? dtStock.Rows[i]["Total Sold"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AB" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AC" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldQtyContri%"]) ? dtStock.Rows[i]["SoldQtyContri%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AC" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AD" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["CostValue"]) ? dtStock.Rows[i]["CostValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AD" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AE" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["RetailValue"]) ? dtStock.Rows[i]["RetailValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AE" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AF" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AF" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AG" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin"]) ? dtStock.Rows[i]["Total Intake Margin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AG" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AH" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Total Intake Margin%"]) ? dtStock.Rows[i]["Total Intake Margin%"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AH" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AI" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin"]) ? dtStock.Rows[i]["Earned Margin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AI" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AJ" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Earned Margin %"]) ? dtStock.Rows[i]["Earned Margin %"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AJ" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AK" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance"]) ? dtStock.Rows[i]["Variance"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AK" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AL" + (j + 4), misValue).Formula = (null != dtStock.Rows[i]["Variance Over Intake %"]) ? dtStock.Rows[i]["Variance Over Intake %"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("AL" + (j + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }
               
            }
           
        }

        #endregion WriteToExcel

        #region InsertStockStatus
        /// <summary>
        /// InsertStockStatus
        /// </summary>
        private void InsertStockStatus()
        {
           // tdLocation.Visible = false;
           
                        
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.FromDate=DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            ObjStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                            
            
           
            ObjStock.SSOperationType =false;
            ObjStock.SSReportOperationType =false;
            ObjStock.SSWeeklyOperationType =false;

            ObjStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
            ObjStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            ObjStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            ObjStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());

            ObjStock.KsaRate = Convert.ToDecimal(txtKSARate.Text.Trim());

            bool Result=ObjStock.InsertStockStatus();
            
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
        
        #region FileDownloadLCP
        /// <summary>
        ///File Download LCP
        /// </summary>

        private void FileDownloadLCP()
        {
           
            string fileNameLCP = ViewState["FileNameLCP"].ToString();

            //string FolderPath = HttpContext.Current.Server.MapPath(".");
            //FolderPath = FolderPath + "\\Reports\\";
            //string FullFilePath = FolderPath + filename;
            FileInfo file = new FileInfo(fileNameLCP);

            
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

        #endregion FileDownloadLCP

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

            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatus_LCP1.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                


            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);

           

            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\StockStatus_LCP_One.xlsx";
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


                Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;




                Excel1.Worksheet xlSheetSummery = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetSummery.Name = "Summary";
                dtStock = objStock.GetAllStockValuesLCP("Summery");
                WriteToExcelLCP(dtStock, xlSheetSummery, "Summery");


                Excel1.Worksheet xlSheetJor = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetJor.Name = "Jordan Summary";
                dtStock = objStock.GetAllStockValuesLCP("JORDAN");
                WriteToExcelLCP(dtStock, xlSheetJor, "JORDAN");


                Excel1.Worksheet xlSheetUae = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheetUae.Name = "UAE Summary";
                dtStock = objStock.GetAllStockValuesLCP("UAE");
                WriteToExcelLCP(dtStock, xlSheetUae, "UAE");

                Excel1.Worksheet xlSheetOman = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheetOman.Name = "Oman Summary";
                
                dtStock = objStock.GetAllStockValuesLCP("OMAN");
                if(dtStock.Rows.Count>0)
                   WriteToExcelLCP(dtStock, xlSheetOman, "OMAN");

                Excel1.Worksheet xlSheetBahrain = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheetBahrain.Name = "Bahrain Summary";

                dtStock = objStock.GetAllStockValuesLCP("BAHRAIN");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheetBahrain, "BAHRAIN");

                Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                xlSheet0408.Name = "DC";
                dtStock = objStock.GetAllStockValuesLCP("DC");
                WriteToExcelLCP(dtStock, xlSheet0408, "DC");

                Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                xlSheet0400.Name = "0400";
                dtStock = objStock.GetAllStockValuesLCP("0400");
                WriteToExcelLCP(dtStock, xlSheet0400, "0400");

                Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                xlSheet0401.Name = "0401";
                dtStock = objStock.GetAllStockValuesLCP("0401");
                WriteToExcelLCP(dtStock, xlSheet0401, "0401");

                Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                xlSheet0402.Name = "0402";
                dtStock = objStock.GetAllStockValuesLCP("0402");
                WriteToExcelLCP(dtStock, xlSheet0402, "0402");

                Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                xlSheet0403.Name = "0403";
                dtStock = objStock.GetAllStockValuesLCP("0403");
                WriteToExcelLCP(dtStock, xlSheet0403, "0403");

                Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                xlSheet0404.Name = "0404";
                dtStock = objStock.GetAllStockValuesLCP("0404");
                WriteToExcelLCP(dtStock, xlSheet0404, "0404");

                Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                xlSheet0405.Name = "0405";
                dtStock = objStock.GetAllStockValuesLCP("0405");
                WriteToExcelLCP(dtStock, xlSheet0405, "0405");

                Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                xlSheet0406.Name = "0406";
                dtStock = objStock.GetAllStockValuesLCP("0406");
                WriteToExcelLCP(dtStock, xlSheet0406, "0406");

                Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                xlSheet0407.Name = "0407";
                dtStock = objStock.GetAllStockValuesLCP("0407");
                WriteToExcelLCP(dtStock, xlSheet0407, "0407");


                Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                xlSheet0409.Name = "0409";
                dtStock = objStock.GetAllStockValuesLCP("0409");
                WriteToExcelLCP(dtStock, xlSheet0409, "0409");

                Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                xlSheet0410.Name = "0410";
                dtStock = objStock.GetAllStockValuesLCP("0410");
                WriteToExcelLCP(dtStock, xlSheet0410, "0410");

                Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                xlSheet0411.Name = "0411";
                dtStock = objStock.GetAllStockValuesLCP("0411");
                WriteToExcelLCP(dtStock, xlSheet0411, "0411");

                Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                xlSheet0412.Name = "0412";
                dtStock = objStock.GetAllStockValuesLCP("0412");
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcelLCP(dtStock, xlSheet0412, "0412");
                }

                Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                xlSheet0414.Name = "0414";
                dtStock = objStock.GetAllStockValuesLCP("0414");
                WriteToExcelLCP(dtStock, xlSheet0414, "0414");

                Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                xlSheet0415.Name = "0415";
                dtStock = objStock.GetAllStockValuesLCP("0415");
                WriteToExcelLCP(dtStock, xlSheet0415, "0415");


                Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                xlSheet0416.Name = "0416";
                dtStock = objStock.GetAllStockValuesLCP("0416");
                WriteToExcelLCP(dtStock, xlSheet0416, "0416");

                Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                xlSheet0417.Name = "0417";
                dtStock = objStock.GetAllStockValuesLCP("0417");
                WriteToExcelLCP(dtStock, xlSheet0417, "0417");

                Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                xlSheet0418.Name = "0418";
                dtStock = objStock.GetAllStockValuesLCP("0418");
                if(dtStock.Rows.Count>0)
                    WriteToExcelLCP(dtStock, xlSheet0418, "0418");

                Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                xlSheet0419.Name = "0419";
                dtStock = objStock.GetAllStockValuesLCP("0419");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0419, "0419");

                Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[26];
                xlSheet0421.Name = "0421";
                dtStock = objStock.GetAllStockValuesLCP("0421");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0421, "0421");

                Excel1.Worksheet xlSheet0422= (Excel1.Worksheet)myExcelWorkbook.Sheets[27];
                xlSheet0422.Name = "0422";
                dtStock = objStock.GetAllStockValuesLCP("0422");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0422, "0422");


                Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[28];
                xlSheet0423.Name = "0423";
                dtStock = objStock.GetAllStockValuesLCP("0423");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0423, "0423");

                Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[29];
                xlSheet0424.Name = "0424";
                dtStock = objStock.GetAllStockValuesLCP("0424");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0424, "0424");

                Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[30];
                xlSheet0425.Name = "0425";
                dtStock = objStock.GetAllStockValuesLCP("0425");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0425, "0425");

                Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[31];
                xlSheet0426.Name = "0426";
                dtStock = objStock.GetAllStockValuesLCP("0426");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0426, "0426");

                Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[32];
                xlSheet0427.Name = "0427";
                dtStock = objStock.GetAllStockValuesLCP("0427");
                if (dtStock.Rows.Count > 0)
                    WriteToExcelLCP(dtStock, xlSheet0427, "0427");


                objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                dtStock = objStock.GetStockStatusLCPSummery();
                WriteToExcelLCPSummery(dtStock, xlSheetSummery);

                //lblMessage.Visible = true;
                //lblMessage.ForeColor = System.Drawing.Color.Green;
                //lblMessage.Text = "Report Generation Complete";
                btnDownloadLCP.Visible = true;
               // btnDownloadLCPStore.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\SSR_LCP" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

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

            // if (location == "JORDAN")
            //     myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Summery Report-LCP For Jordan  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();
            // else if (location == "UAE")
            //     myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Summery Report-LCP For UAE  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();
            // else if (location == "Summery")
            //     myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Summery Report-LCP from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();

            // else
            //     myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Report-LCP For Store No : " + location + "  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();


            //// myExcelWorksheet.get_Range("A1", misValue).Formula = "Stock Status Report-LCP For Store No : " + location + "  from  " + dtStock.Rows[0]["FromDate"].ToString() + "  To  " + dtStock.Rows[0]["ToDate"].ToString();
            // myExcelWorksheet.get_Range("A1", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("A1", misValue).Font.Size = 14;

            // myExcelWorksheet.get_Range("A2", misValue).Formula = "Category Code";
            // myExcelWorksheet.get_Range("A2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("A2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("A2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            // myExcelWorksheet.get_Range("B2", misValue).Formula = "Product Group Code";
            // myExcelWorksheet.get_Range("B2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("B2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("B2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            // myExcelWorksheet.get_Range("C2", misValue).Formula = "TotalCls";
            // myExcelWorksheet.get_Range("C2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("C2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("C2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            // myExcelWorksheet.get_Range("D2", misValue).Formula = "Total Sold";
            // myExcelWorksheet.get_Range("D2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("D2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("D2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            // myExcelWorksheet.get_Range("E2", misValue).Formula = "Weekly Sold Qty";
            // myExcelWorksheet.get_Range("E2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("E2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("E2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            // myExcelWorksheet.get_Range("F2", misValue).Formula = "Sell Thru(%)";
            // myExcelWorksheet.get_Range("F2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("F2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("F2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.RosyBrown);

            // myExcelWorksheet.get_Range("G2", misValue).Formula = "Sold Avg/Day";
            // myExcelWorksheet.get_Range("G2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("G2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("G2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

            // myExcelWorksheet.get_Range("H2", misValue).Formula = "Week Cover";
            // myExcelWorksheet.get_Range("H2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("H2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("H2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);


            // myExcelWorksheet.get_Range("I2", misValue).Formula = "Sold Qty";
            // myExcelWorksheet.get_Range("I2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("I2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("I2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            // myExcelWorksheet.get_Range("J2", misValue).Formula = "Sold Qty Contri(%)";
            // myExcelWorksheet.get_Range("J2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("J2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("J2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);


            // myExcelWorksheet.get_Range("K2", misValue).Formula = "Cost Value";
            // myExcelWorksheet.get_Range("K2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("K2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("K2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            // myExcelWorksheet.get_Range("L2", misValue).Formula = "Retail Value";
            // myExcelWorksheet.get_Range("L2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("L2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("L2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);


            // myExcelWorksheet.get_Range("M2", misValue).Formula = "Sold Value";
            // myExcelWorksheet.get_Range("M2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("M2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("M2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);


            // myExcelWorksheet.get_Range("N2", misValue).Formula = "Intake Margin";
            // myExcelWorksheet.get_Range("N2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("N2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("N2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);


            // myExcelWorksheet.get_Range("O2", misValue).Formula = "Intake Margin(%)";
            // myExcelWorksheet.get_Range("O2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("O2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("O2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            // myExcelWorksheet.get_Range("P2", misValue).Formula = "Earned Margin";
            // myExcelWorksheet.get_Range("P2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("P2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("P2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            // myExcelWorksheet.get_Range("Q2", misValue).Formula = "Earned Margin(%)";
            // myExcelWorksheet.get_Range("Q2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("Q2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("Q2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            // myExcelWorksheet.get_Range("R2", misValue).Formula = "Variance (Intake Vs Earned)";
            // myExcelWorksheet.get_Range("R2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("R2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("R2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            // myExcelWorksheet.get_Range("S2", misValue).Formula = "Variance(%)";
            // myExcelWorksheet.get_Range("S2", misValue).Font.Bold = true;
            // myExcelWorksheet.get_Range("S2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            // myExcelWorksheet.get_Range("S2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkViolet);

            //int flag = 0;

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

                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin"]) ? dtStock.Rows[i]["IntakeMargin"].ToString() : "0";
                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin(%)"]) ? dtStock.Rows[i]["IntakeMargin(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin"]) ? dtStock.Rows[i]["EarnedMargin"].ToString() : "0";
                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin(%)"]) ? dtStock.Rows[i]["EarnedMargin(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(Intake Vs Earned)"]) ? dtStock.Rows[i]["Variance(Intake Vs Earned)"].ToString() : "0";
                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(%)"]) ? dtStock.Rows[i]["Variance(%)"].ToString() + "%" : "0";
                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Font.Bold = true;
                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
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

                    myExcelWorksheet.get_Range("N" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("N" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin"]) ? dtStock.Rows[i]["IntakeMargin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("P" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["IntakeMargin(%)"]) ? dtStock.Rows[i]["IntakeMargin(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("P" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin"]) ? dtStock.Rows[i]["EarnedMargin"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["EarnedMargin(%)"]) ? dtStock.Rows[i]["EarnedMargin(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("R" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(Intake Vs Earned)"]) ? dtStock.Rows[i]["Variance(Intake Vs Earned)"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("T" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["Variance(%)"]) ? dtStock.Rows[i]["Variance(%)"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("T" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

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


            //myExcelWorksheet.get_Range("U1", misValue).Formula = "Stores Closing Quantity ";
            //myExcelWorksheet.get_Range("U1", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("U1", misValue).Font.Size = 14;

            //myExcelWorksheet.get_Range("U2", misValue).Formula = "0400";
            //myExcelWorksheet.get_Range("U2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("U2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("U2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("V2", misValue).Formula = "0401";
            //myExcelWorksheet.get_Range("V2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("V2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("V2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("W2", misValue).Formula = "0402";
            //myExcelWorksheet.get_Range("W2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("W2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("W2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("X2", misValue).Formula = "0403";
            //myExcelWorksheet.get_Range("X2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("X2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("X2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("Y2", misValue).Formula = "0404";
            //myExcelWorksheet.get_Range("Y2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("Y2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("Y2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("Z2", misValue).Formula = "0405";
            //myExcelWorksheet.get_Range("Z2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("Z2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("Z2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("AA2", misValue).Formula = "0406";
            //myExcelWorksheet.get_Range("AA2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AA2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AA2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("AB2", misValue).Formula = "0407";
            //myExcelWorksheet.get_Range("AB2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AB2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AB2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);



            //myExcelWorksheet.get_Range("AC2", misValue).Formula = "0409";
            //myExcelWorksheet.get_Range("AC2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AC2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AC2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


            //myExcelWorksheet.get_Range("AE2", misValue).Formula = "0410";
            //myExcelWorksheet.get_Range("AE2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AE2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AE2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("AF2", misValue).Formula = "0411";
            //myExcelWorksheet.get_Range("AF2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AF2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AF2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


            //myExcelWorksheet.get_Range("AG2", misValue).Formula = "0414";
            //myExcelWorksheet.get_Range("AG2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AG2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AG2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


            //myExcelWorksheet.get_Range("AH2", misValue).Formula = "0415";
            //myExcelWorksheet.get_Range("AH2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AH2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AH2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);


            //myExcelWorksheet.get_Range("AI2", misValue).Formula = "0416";
            //myExcelWorksheet.get_Range("AI2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AI2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AI2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("AJ2", misValue).Formula = "0417";
            //myExcelWorksheet.get_Range("AJ2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AJ2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AJ2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //myExcelWorksheet.get_Range("AJ2", misValue).Formula = "0408";
            //myExcelWorksheet.get_Range("AJ2", misValue).Font.Bold = true;
            //myExcelWorksheet.get_Range("AJ2", misValue).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //myExcelWorksheet.get_Range("AJ2", misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            //int flag = 0;
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

                    else if (Convert.ToInt32(dtStock.Rows[i]["0420"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AU" + (j + 3), misValue).Formula = dtStock.Rows[i]["0420"].ToString();
                        myExcelWorksheet.get_Range("AU" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AU" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AU" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    }

                    else if (Convert.ToInt32(dtStock.Rows[i]["0427"]) != 0)
                    {
                        myExcelWorksheet.get_Range("AV" + (j + 3), misValue).Formula = dtStock.Rows[i]["0427"].ToString();
                        myExcelWorksheet.get_Range("AV" + (j + 3), misValue).Font.Bold = true;
                        myExcelWorksheet.get_Range("AV" + (j + 3), misValue).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        BorderAround(myExcelWorksheet.get_Range("AV" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
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

                    myExcelWorksheet.get_Range("AU" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0420"]) ? dtStock.Rows[i]["0420"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AU" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("AV" + (j + 3), misValue).Formula = (null != dtStock.Rows[i]["0427"]) ? dtStock.Rows[i]["0427"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("AV" + (j + 3), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

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

        #region DeleteColumnSSR
        /// <summary>
        /// DeleteColumnSSR
        /// </summary>
        /// <param name="myExcelWorkbook"></param>
        private void DeleteColumnSSR(Excel1.Workbook myExcelWorkbook,bool Type)
        {
            object misValue = System.Reflection.Missing.Value;

            if (Type == true)
            {


                for (int i = 2; i <= 20; i++)
                {

                    Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[i];

                    xlSheet.get_Range("F1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("F1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("F1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("J1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("J1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("M1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("M1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("N1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("N1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("Q1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("Q1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                }
            }

            else
            {

                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];

                xlSheet.get_Range("F1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("F1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("F1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("J1", misValue).EntireColumn.Delete(misValue);

                xlSheet.get_Range("J1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);

                xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("M1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("M1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("N1", misValue).EntireColumn.Delete(misValue);

                xlSheet.get_Range("N1", misValue).EntireColumn.Delete(misValue);

                xlSheet.get_Range("Q1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("Q1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);

                xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
                xlSheet.get_Range("R1", misValue).EntireColumn.Delete(misValue);
            }
        }
        
        #endregion DeleteColumnSSR

        #region DeleteColumnLCP
        /// <summary>
        /// Delete Column LCP
        /// </summary>
        /// <param name="myExcelWorkbook"></param>
        private void DeleteColumnLCP(Excel1.Workbook myExcelWorkbook, bool Type)
        {
            object misValue = System.Reflection.Missing.Value;

            if (Type == true)
            {


                for (int i = 2; i <= 20; i++)
                {

                    Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[i];

                    xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue); 
                }
            }

            else
            {

                    Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];

                    xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("K1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);

                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue);
                    xlSheet.get_Range("L1", misValue).EntireColumn.Delete(misValue); 
            }
        }

        #endregion DeleteColumnLCP

        #region GenerateWSSIReport
        /// <summary>
        /// To generate excel report for stock values
        /// </summary>
        private void GenerateWSSIReport()
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




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\WSSI_Test2.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtWeekNo.Text.Trim().Length > 0)
            {
               // string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\WSSI_Report.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                objStock.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                objStock.Year = ddlYear.SelectedItem.Text;
                objStock.SeasonCode = "S";

                dtStock = objStock.GetWSSIReport();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = "Summery";

                    //xlSheet.get_Range("A1", misValue).Formula = "Store No: " + location + " WSSI Report From " + txtFromDate.Text + " To " + txtToDate.Text + "  (Week No :" + dtStock.Rows[0]["WeekNo"].ToString() + ")";
                    //xlSheet.get_Range("A1", misValue).Font.Bold = true;
                    
                    WriteToExcelWssi(dtStock, xlSheet,"S");


                    objStock.SeasonCode = "W";
                    dtStock = objStock.GetWSSIReport();
                    WriteToExcelWssi(dtStock, xlSheet, "W");

                    objStock.SeasonCode = "Y";
                    dtStock = objStock.GetWSSIReport();
                    int j=WriteToExcelWssi(dtStock, xlSheet, "Y");

                    objStock.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                    objStock.Year = ddlYear.SelectedItem.Text;
                    objStock.Type = true;

                    dtStock = objStock.GetWSSIForcast();
                    WriteToExcelWssiBudget(dtStock, xlSheet);

                    objStock.Type = false;
                    dtStock = objStock.GetWSSIForcast();
                    WriteToExcelWssiForcast(dtStock, xlSheet,j);
                    
                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnDownloadWssi.Visible = true;
                    
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
                //do nothing
            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\WSSI_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameWSSI"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();


            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateWSSIReport

        #region WriteToExcelWssi
        /// <summary>
        ///Write To Excel Wssi
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelWssi(DataTable dtStock, Excel1.Worksheet myExcelWorksheet,string seasonCode)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 8;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                   
                if(seasonCode=="S")
                { 
                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["WeekNo"]) ? dtStock.Rows[i]["WeekNo"].ToString() : "0";

                    myExcelWorksheet.get_Range("E" + j , misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("F" +j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("G" + j , misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";

                    myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["AvgRetailPrice"]) ? dtStock.Rows[i]["AvgRetailPrice"].ToString() : "0";
                
                }

                if (seasonCode == "W")
                {
                    myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                }

                if (seasonCode == "Y")
                {
                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("V" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                }
            
            }


            //if (seasonCode == "S")
            //{
            //    myExcelWorksheet.get_Range("B" + j, misValue).Formula = "Grand Total";

            //    myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E8" + ":E" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("E" + j, misValue).Font.Bold = true;
            //    myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
            //    myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Font.Bold = true;

            //    myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G8" + ":G" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("G" + j, misValue).Font.Bold = true;

            //    myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=SUM(I8" + ":I" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("I" + j, misValue).Font.Bold = true;
            //}

            //if (seasonCode == "W")
            //{
            //    myExcelWorksheet.get_Range("L" + j, misValue).Formula = "=SUM(L8" + ":L" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("L" + j, misValue).Font.Bold = true;

            //    myExcelWorksheet.get_Range("N" + j, misValue).Formula = "=SUM(N8" + ":N" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("N" + j, misValue).Font.Bold = true;

            //    myExcelWorksheet.get_Range("P" + j, misValue).Formula = "=SUM(P8" + ":P" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("P" + j, misValue).Font.Bold = true;
            //}

            //if (seasonCode == "Y")
            //{
            //    myExcelWorksheet.get_Range("S" + j, misValue).Formula = "=SUM(S8" + ":S" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("S" + j, misValue).Font.Bold = true;

            //    myExcelWorksheet.get_Range("U" + j, misValue).Formula = "=SUM(U8" + ":U" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("U" + j, misValue).Font.Bold = true;

            //    myExcelWorksheet.get_Range("W" + j, misValue).Formula = "=SUM(W8" + ":W" + (j - 1) + ")";
            //    myExcelWorksheet.get_Range("W" + j, misValue).Font.Bold = true;
            //}

            return j;
        }

        #endregion WriteToExcelWssi

        #region WriteToExcelWssiForcast
        /// <summary>
        ///Write To Excel Wssi Forcast
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelWssiForcast(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j )
        {
            object misValue = System.Reflection.Missing.Value;
           
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                
                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["WeekNo"]) ? dtStock.Rows[i]["WeekNo"].ToString() : "0";

                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetSSCoded"]) ? dtStock.Rows[i]["BudgetSSCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetSSCoded"]) ? dtStock.Rows[i]["BudgetSSCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=F"+(j-1)+"+G"+j+"-E"+j+"+I"+j;

                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["DcIntakeSSCoded"]) ? dtStock.Rows[i]["DcIntakeSSCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = "=F"+j+"/E"+j;

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = "0";


                    myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetAWCoded"]) ? dtStock.Rows[i]["BudgetAWCoded"].ToString() : "0";
                  
                    myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetAWCoded"]) ? dtStock.Rows[i]["BudgetAWCoded"].ToString() : "0";
                    
                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = "=M"+(j-1)+"+N"+j+"-L"+j+"+P"+j;

                    myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["DcIntakeAWCoded"]) ? dtStock.Rows[i]["DcIntakeAWCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = "=M"+j+"/L"+j;

                    myExcelWorksheet.get_Range("P" + j, misValue).Formula = "0";


                    myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetYCoded"]) ? dtStock.Rows[i]["BudgetYCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetYCoded"]) ? dtStock.Rows[i]["BudgetYCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("T" + j, misValue).Formula = "=T"+(j-1)+"+U"+j+"-S"+j+"+W"+j;

                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["DcIntakeYCoded"]) ? dtStock.Rows[i]["DcIntakeYCoded"].ToString() : "0";

                    myExcelWorksheet.get_Range("V" + j, misValue).Formula = "=T"+j+"/S"+j;

                    myExcelWorksheet.get_Range("W" + j, misValue).Formula = "0";

                    myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["AveRetailPrice"]) ? dtStock.Rows[i]["AveRetailPrice"].ToString() : "0";

            
            }

            
                myExcelWorksheet.get_Range("B" + j, misValue).Formula = "Grand Total";

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = "=SUM(D8" + ":D" + (j - 1) + ")";

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E8" + ":E" + (j - 1) + ")";
                myExcelWorksheet.get_Range("E" + j, misValue).Font.Bold = true;
                myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
                myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Font.Bold = true;

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G8" + ":G" + (j - 1) + ")";
                myExcelWorksheet.get_Range("G" + j, misValue).Font.Bold = true;

                myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=SUM(I8" + ":I" + (j - 1) + ")";
                myExcelWorksheet.get_Range("I" + j, misValue).Font.Bold = true;

               
                myExcelWorksheet.get_Range("K" + j, misValue).Formula = "=SUM(K8" + ":K" + (j - 1) + ")";
            
                myExcelWorksheet.get_Range("L" + j, misValue).Formula = "=SUM(L8" + ":L" + (j - 1) + ")";
                myExcelWorksheet.get_Range("L" + j, misValue).Font.Bold = true;

                myExcelWorksheet.get_Range("N" + j, misValue).Formula = "=SUM(N8" + ":N" + (j - 1) + ")";
                myExcelWorksheet.get_Range("N" + j, misValue).Font.Bold = true;

                myExcelWorksheet.get_Range("P" + j, misValue).Formula = "=SUM(P8" + ":P" + (j - 1) + ")";
                myExcelWorksheet.get_Range("P" + j, misValue).Font.Bold = true;

                myExcelWorksheet.get_Range("R" + j, misValue).Formula = "=SUM(R8" + ":R" + (j - 1) + ")";

                myExcelWorksheet.get_Range("S" + j, misValue).Formula = "=SUM(S8" + ":S" + (j - 1) + ")";
                myExcelWorksheet.get_Range("S" + j, misValue).Font.Bold = true;

                myExcelWorksheet.get_Range("U" + j, misValue).Formula = "=SUM(U8" + ":U" + (j - 1) + ")";
                myExcelWorksheet.get_Range("U" + j, misValue).Font.Bold = true;

                myExcelWorksheet.get_Range("W" + j, misValue).Formula = "=SUM(W8" + ":W" + (j - 1) + ")";
                myExcelWorksheet.get_Range("W" + j, misValue).Font.Bold = true;
            



                //myExcelWorksheet.get_Range("B" + j, misValue).Formula = "Grand Total";

                //myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E8" + ":E" + (j - 1) + ")";
                //myExcelWorksheet.get_Range("E" + j, misValue).Font.Bold = true;
                //myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
                //myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Font.Bold = true;

                //myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G8" + ":G" + (j - 1) + ")";
                //myExcelWorksheet.get_Range("G" + j, misValue).Font.Bold = true;

                //myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=SUM(I8" + ":I" + (j - 1) + ")";
                //myExcelWorksheet.get_Range("I" + j, misValue).Font.Bold = true;
         
        }

        #endregion WriteToExcelWssiForcast

        #region WriteToExcelWssiBudget
        /// <summary>
        ///Write To Excel Wssi Budget
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        private void WriteToExcelWssiBudget(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 8;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetSSCoded"]) ? dtStock.Rows[i]["BudgetSSCoded"].ToString() : "0";

                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetAWCoded"]) ? dtStock.Rows[i]["BudgetAWCoded"].ToString() : "0";

                myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["BudgetYCoded"]) ? dtStock.Rows[i]["BudgetYCoded"].ToString() : "0";
            }
        }

        #endregion WriteToExcelWssiBudget



        #region Insert Wssi Report
        /// <summary>
        /// Insert Wssi Report
        /// </summary>
        private void InsertWssiReport()
        {
            if (txtWeekNo.Text.Trim().Length > 0)
            {
                GetStockDetails objStock = new GetStockDetails();


                objStock.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                objStock.Year = ddlYear.SelectedItem.Text.ToString();

                objStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());

                objStock.KsaRate = Convert.ToDecimal(txtKSARate.Text.Trim());

                objStock.InsertWssiReport();
            }
            else
            {
                lblMessage.Text = "Enter Valid Week No !";
            }
        }
        #endregion Insert Wssi Report

        #region InsertProductGroupCmpReport
        /// <summary>
        /// Insert Product Group Cmp Report
        /// </summary>
        private void InsertProductGroupCmpReport()
        {
            if (txtWeekNo.Text.Trim().Length > 0)
            {
                GetStockDetails objStock = new GetStockDetails();
                
                objStock.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                objStock.IntYear = Convert.ToInt32(ddlYear.SelectedItem.Value);

                objStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objStock.KsaRate = Convert.ToDecimal(txtKSARate.Text.Trim());
                
                objStock.InsertProductGroupCmpReport();
            }
            else
            {
                lblMessage.Text = "Enter Valid Week No !";
            }
        }
        #endregion InsertProductGroupCmpReport
        
        #region GeneratePgCmpReport
        /// <summary>
        /// To generate PgCmpReport
        /// </summary>
        private void GeneratePgCmpReport()
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




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\PgCmp_Report.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtLocation.Text.Trim().Length > 0)
            {
                // string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\PgCmp_Test _One.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                objStock.Location =txtLocation.Text.Trim();
               
                dtStock = objStock.GetPgCmpReport();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = txtLocation.Text.Trim(); 
                    WriteToExcelPgCmp(dtStock, xlSheet);

                                        
                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnPGCompare.Visible = true;

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
                Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet0400.Name = "0400";
                objStock.Location = "0400";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0400);

                Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheet0401.Name = "0401";
                objStock.Location = "0401";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0401);

                Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheet0402.Name = "0402";
                objStock.Location = "0402";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0402);

                Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet0403.Name = "0403";
                objStock.Location = "0403";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0403);

                Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheet0404.Name = "0404";
                objStock.Location = "0404";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0404);

                Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheet0405.Name = "0405";
                objStock.Location = "0405";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0405);

                Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                xlSheet0406.Name = "0406";
                objStock.Location = "0406";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0406);

                Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                xlSheet0407.Name = "0407";
                objStock.Location = "0407";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0407);

                Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                xlSheet0409.Name = "0409";
                objStock.Location = "0409";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0409);

                Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                xlSheet0410.Name = "0410";
                objStock.Location = "0410";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0410);

                Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                xlSheet0411.Name = "0411";
                objStock.Location = "0411";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0411);

                Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                xlSheet0412.Name = "0412";
                objStock.Location = "0412";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcelPgCmp(dtStock, xlSheet0412);
                }

                Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                xlSheet0414.Name = "0414";
                objStock.Location = "0414";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0414);

                Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                xlSheet0415.Name = "0415";
                objStock.Location = "0415";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0415);

                Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                xlSheet0416.Name = "0416";
                objStock.Location = "0416";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0416);

                Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                xlSheet0417.Name = "0417";
                objStock.Location = "0417";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0417);


                Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                xlSheet0418.Name = "0418";
                objStock.Location = "0418";
                dtStock = objStock.GetPgCmpReport();
                if(dtStock.Rows.Count>0)
                    WriteToExcelPgCmp(dtStock, xlSheet0418);

                Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                xlSheet0419.Name = "0419";
                objStock.Location = "0419";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0419);


                Excel1.Worksheet xlSheet0420= (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                xlSheet0420.Name = "0421";
                objStock.Location = "0421";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0420);


                Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                xlSheet0421.Name = "0422";
                objStock.Location = "0422";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0421);

                Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                xlSheet0426.Name = "0423";
                objStock.Location = "0423";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0426);

                Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                xlSheet0427.Name = "0424";
                objStock.Location = "0424";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0427);


                Excel1.Worksheet xlSheet0429 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                xlSheet0429.Name = "0425";
                objStock.Location = "0425";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0429);

                Excel1.Worksheet xlSheet0430 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                xlSheet0430.Name = "0426";
                objStock.Location = "0426";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0430);

                Excel1.Worksheet xlSheet0431 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                xlSheet0431.Name = "0427";
                objStock.Location = "0427";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                    WriteToExcelPgCmp(dtStock, xlSheet0431);



                Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[26];
                xlSheet0422.Name = "JORDAN";
                objStock.Location = "JORDAN";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0422);

                Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[27];
                xlSheet0423.Name = "UAE";
                objStock.Location = "UAE";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0423);

                Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[28];
                xlSheet0424.Name = "OMAN";
                objStock.Location = "OMAN";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcelPgCmp(dtStock, xlSheet0424);
                }

                Excel1.Worksheet xlSheetBah = (Excel1.Worksheet)myExcelWorkbook.Sheets[29];
                xlSheetBah.Name = "BAHRAIN";
                objStock.Location = "BAHRAIN";
                dtStock = objStock.GetPgCmpReport();
                if (dtStock.Rows.Count > 0)
                {
                    WriteToExcelPgCmp(dtStock, xlSheetBah);
                }

                Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[30];
                xlSheet0425.Name = "SUMMARY";
                objStock.Location = "Summery";
                dtStock = objStock.GetPgCmpReport();
                WriteToExcelPgCmp(dtStock, xlSheet0425);

                
                Excel1.Worksheet xlSheet0428 = (Excel1.Worksheet)myExcelWorkbook.Sheets[31];
                xlSheet0428.Name = "SUMMARY-DIVISON";
                dtStock = objStock.GetPgcmpSummaryByDivision();
                WriteToExcelPgCmpDivision(dtStock, xlSheet0428);

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\PgCmp_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNamePgCmp"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();


            lblMessage.Visible = true;
            lblMessage.ForeColor = System.Drawing.Color.Green;
            lblMessage.Text = "Report Generation Complete";

            btnPGCompare.Visible = true;

            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GeneratePgCmpReport

        #region WriteToExcelPgCmp
        /// <summary>
        ///Write To Excel Product Group Compare Report
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelPgCmp(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 3;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["WeekNo"]) ? dtStock.Rows[i]["WeekNo"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroupCode"]) ? dtStock.Rows[i]["ProductGroupCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroupDesc"]) ? dtStock.Rows[i]["ProductGroupDesc"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQty"]) ? dtStock.Rows[i]["SoldQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQtyLastYear"]) ? dtStock.Rows[i]["SoldQtyLastYear"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=D" + (j) + "-" + "E" + (j);
                    BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQtyLastYearPer"]) ? dtStock.Rows[i]["SoldQtyLastYearPer"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldValueLastYear"]) ? dtStock.Rows[i]["SoldValueLastYear"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + j, misValue).Formula = "=H" + (j) + "-" + "I" + (j);
                    BorderAround(myExcelWorksheet.get_Range("J" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldValueLastYearPer"]) ? dtStock.Rows[i]["SoldValueLastYearPer"].ToString() + "%" : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
          }

                           
                myExcelWorksheet.get_Range("A" + (j-1), misValue).Formula = "Grand Total";
                myExcelWorksheet.get_Range("A" + (j-1), misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
                myExcelWorksheet.get_Range("A" + (j-1), misValue).EntireRow.Font.Bold = true;
                BorderAround(myExcelWorksheet.get_Range("A" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                
                myExcelWorksheet.get_Range("B" + (j - 1), misValue).Formula = "";
                BorderAround(myExcelWorksheet.get_Range("B" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + (j - 1), misValue).Formula = "=SUM(F3" + ":F" + (j - 2) + ")";
                BorderAround(myExcelWorksheet.get_Range("F" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("J" + (j - 1), misValue).Formula = "=SUM(J3" + ":J" + (j - 2) + ")";
                BorderAround(myExcelWorksheet.get_Range("J" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                //myExcelWorksheet.get_Range("D" + j, misValue).Formula = "=SUM(D3" + ":D" + (j - 1) + ")";
               
                //myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=SUM(F3" + ":F" + (j - 1) + ")";


                //myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E3" + ":E" + (j - 1) + ")";
                //myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G3" + ":G" + (j - 1) + ")";
              
        }

        #endregion WriteToExcelPgCmp

        #region WriteToExcelPgCmpDivision
        /// <summary>
        /// WriteToExcelPgCmpDivision
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelPgCmpDivision(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 3;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["WeekNo"]) ? dtStock.Rows[i]["WeekNo"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroupCode"]) ? dtStock.Rows[i]["ProductGroupCode"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroupDesc"]) ? dtStock.Rows[i]["ProductGroupDesc"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQty"]) ? dtStock.Rows[i]["SoldQty"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQtyLastYear"]) ? dtStock.Rows[i]["SoldQtyLastYear"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=D" + (j) + "-" + "E" + (j);
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQtyLastYearPer"]) ? dtStock.Rows[i]["SoldQtyLastYearPer"].ToString() + "%" : "0";
                BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldValue"]) ? dtStock.Rows[i]["SoldValue"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("H" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldValueLastYear"]) ? dtStock.Rows[i]["SoldValueLastYear"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("J" + j, misValue).Formula = "=H" + (j) + "-" + "I" + (j);
                BorderAround(myExcelWorksheet.get_Range("J" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldValueLastYearPer"]) ? dtStock.Rows[i]["SoldValueLastYearPer"].ToString() + "%" : "0";
                BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                if(dtStock.Rows[i]["ProductGroupCode"].ToString()=="")
                {
                    myExcelWorksheet.get_Range("A" +j, "K"+ j).Font.Bold = true;
                    myExcelWorksheet.get_Range("A" + j, "K" + j).Interior.Color = System.Drawing.Color.Yellow;
                }

            }


            //myExcelWorksheet.get_Range("A" + (j - 1), misValue).Formula = "Grand Total";
            //myExcelWorksheet.get_Range("A" + (j - 1), misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
            //myExcelWorksheet.get_Range("A" + (j - 1), misValue).EntireRow.Font.Bold = true;
            //BorderAround(myExcelWorksheet.get_Range("A" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            //myExcelWorksheet.get_Range("B" + (j - 1), misValue).Formula = "";
            //BorderAround(myExcelWorksheet.get_Range("B" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            //myExcelWorksheet.get_Range("F" + (j - 1), misValue).Formula = "=SUM(F3" + ":F" + (j - 2) + ")";
            //BorderAround(myExcelWorksheet.get_Range("F" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            //myExcelWorksheet.get_Range("J" + (j - 1), misValue).Formula = "=SUM(J3" + ":J" + (j - 2) + ")";
            //BorderAround(myExcelWorksheet.get_Range("J" + (j - 1), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            //myExcelWorksheet.get_Range("D" + j, misValue).Formula = "=SUM(D3" + ":D" + (j - 1) + ")";

            //myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=SUM(F3" + ":F" + (j - 1) + ")";


            //myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E3" + ":E" + (j - 1) + ")";
            //myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G3" + ":G" + (j - 1) + ")";

        }

        #endregion WriteToExcelPgCmpDivision


        #region GetProcessStatus
        /// <summary>
        /// Get Process Status
        /// </summary>
        private bool GetProcessStatus()
        {
            GetStockDetails objStock = new GetStockDetails();
            objStock.ProcessStatusId = StockStatusProcessId;
            DataTable dtStatus = objStock.GetProcessStatus();
            bool Flag = Convert.ToBoolean(dtStatus.Rows[0]["Flag"]);

            return Flag;
        }
        #endregion GetProcessStatus


        #region Insert Wssi Division Report
        /// <summary>
        /// Insert Wssi Division Report
        /// </summary>
        private void InsertWssiDivisionReport()
        {
            if (txtWeekNo.Text.Trim().Length > 0)
            {
                GetStockDetails objStock = new GetStockDetails();


                objStock.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                objStock.Year = ddlYear.SelectedItem.Text.ToString();

                objStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());

                objStock.InsertWssiDivisionReport();
            }
            else
            {
                lblMessage.Text = "Enter Valid Week No !";
            }
        }
        #endregion Insert Wssi Division Report

        #region GenerateWSSIDivisionReport
        /// <summary>
        /// To generate excel WSSI Division Report
        /// </summary>
        private void GenerateWSSIDivisionReport()
        {

            //try
            //{

            Excel1.Application myExcelApp;

            Excel1.Workbooks myExcelWorkbooks;

            Excel1.Workbook myExcelWorkbook=null;


            object misValue = System.Reflection.Missing.Value;

            myExcelApp = new Excel1.Application();

            myExcelApp.Visible = false;

            myExcelWorkbooks = myExcelApp.Workbooks;

            string fileName;

            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtWeekNo.Text.Trim().Length > 0)
            {
                // string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\WSSIDivisionReport.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                objStock.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                objStock.Year = ddlYear.SelectedItem.Text;
                objStock.DivisionCode = "C";

                dtStock = objStock.GetWSSIDivisionReport();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = "Summery";

                    //xlSheet.get_Range("A1", misValue).Formula = "Store No: " + location + " WSSI Report From " + txtFromDate.Text + " To " + txtToDate.Text + "  (Week No :" + dtStock.Rows[0]["WeekNo"].ToString() + ")";
                    //xlSheet.get_Range("A1", misValue).Font.Bold = true;

                    WriteToExcelWssiDivision(dtStock, xlSheet, "C");


                    objStock.DivisionCode = "F";
                    dtStock = objStock.GetWSSIDivisionReport();
                    WriteToExcelWssiDivision(dtStock, xlSheet, "F");

                    objStock.DivisionCode = "H";
                    dtStock = objStock.GetWSSIDivisionReport();
                    WriteToExcelWssiDivision(dtStock, xlSheet, "H");

                    objStock.DivisionCode = "L";
                    dtStock = objStock.GetWSSIDivisionReport();
                    WriteToExcelWssiDivision(dtStock, xlSheet, "L");

                    objStock.DivisionCode = "M";
                    dtStock = objStock.GetWSSIDivisionReport();
                    WriteToExcelWssiDivision(dtStock, xlSheet, "M");

                    
                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnWssiDivision.Visible = true;

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
                //do nothing
            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\WSSI_Division" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameWssiDivision"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();


            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateWSSIDivisionReport

        #region Write To Excel Wssi Division
        /// <summary>
        /// Write To Excel Wssi Division
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelWssiDivision(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string divisionCode)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 7;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                if (divisionCode == "C")
                {
                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["WeekNo"]) ? dtStock.Rows[i]["WeekNo"].ToString() : "0";

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";

                    myExcelWorksheet.get_Range("AM" + j, misValue).Formula = (null != dtStock.Rows[i]["AvgRetailPrice"]) ? dtStock.Rows[i]["AvgRetailPrice"].ToString() : "0";

                }

                if (divisionCode == "F")
                {
                    myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                }

                if (divisionCode == "H")
                {
                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("V" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                }

                if (divisionCode == "L")
                {
                    myExcelWorksheet.get_Range("Z" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("AA" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("AB" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("AC" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("AD" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                }

                if (divisionCode == "M")
                {
                    myExcelWorksheet.get_Range("AG" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldUnits"]) ? dtStock.Rows[i]["SoldUnits"].ToString() : "0";

                    myExcelWorksheet.get_Range("AH" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingStock"]) ? dtStock.Rows[i]["ClosingStock"].ToString() : "0";

                    myExcelWorksheet.get_Range("AI" + j, misValue).Formula = (null != dtStock.Rows[i]["DCIntake"]) ? dtStock.Rows[i]["DCIntake"].ToString() : "0";

                    myExcelWorksheet.get_Range("AJ" + j, misValue).Formula = (null != dtStock.Rows[i]["Cover"]) ? dtStock.Rows[i]["Cover"].ToString() : "0";

                    myExcelWorksheet.get_Range("AK" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                }


            }

            myExcelWorksheet.get_Range("B" + j, misValue).Formula = "Grand Total";
            myExcelWorksheet.get_Range("D" + j, misValue).Formula = "=SUM(D7" + ":D" + (j - 1) + ")";
            myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E7" + ":E" + (j - 1) + ")";
            myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
            myExcelWorksheet.get_Range("E" + j, misValue).EntireRow.Font.Bold = true;

            myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G7" + ":G" + (j - 1) + ")";
            myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=SUM(I7" + ":I" + (j - 1) + ")";
            myExcelWorksheet.get_Range("K" + j, misValue).Formula = "=SUM(K7" + ":K" + (j - 1) + ")";
            myExcelWorksheet.get_Range("L" + j, misValue).Formula = "=SUM(L7" + ":L" + (j - 1) + ")";
            
            myExcelWorksheet.get_Range("N" + j, misValue).Formula = "=SUM(N7" + ":N" + (j - 1) + ")";
            myExcelWorksheet.get_Range("P" + j, misValue).Formula = "=SUM(P7" + ":P" + (j - 1) + ")";
            myExcelWorksheet.get_Range("R" + j, misValue).Formula = "=SUM(R7" + ":R" + (j - 1) + ")";
            myExcelWorksheet.get_Range("S" + j, misValue).Formula = "=SUM(S7" + ":S" + (j - 1) + ")";
            
            myExcelWorksheet.get_Range("U" + j, misValue).Formula = "=SUM(U7" + ":U" + (j - 1) + ")";
            myExcelWorksheet.get_Range("W" + j, misValue).Formula = "=SUM(W7" + ":W" + (j - 1) + ")";
            myExcelWorksheet.get_Range("Y" + j, misValue).Formula = "=SUM(Y7" + ":Y" + (j - 1) + ")";
            myExcelWorksheet.get_Range("Z" + j, misValue).Formula = "=SUM(Z7" + ":Z" + (j - 1) + ")";

            myExcelWorksheet.get_Range("AB" + j, misValue).Formula = "=SUM(AB7" + ":AB" + (j - 1) + ")";
            myExcelWorksheet.get_Range("AD" + j, misValue).Formula = "=SUM(AD7" + ":AD" + (j - 1) + ")";
            myExcelWorksheet.get_Range("AF" + j, misValue).Formula = "=SUM(AF7" + ":AF" + (j - 1) + ")";
            myExcelWorksheet.get_Range("AG" + j, misValue).Formula = "=SUM(AG7" + ":AG" + (j - 1) + ")";
            
            myExcelWorksheet.get_Range("AI" + j, misValue).Formula = "=SUM(AI7" + ":AI" + (j - 1) + ")";
            myExcelWorksheet.get_Range("AK" + j, misValue).Formula = "=SUM(AK7" + ":AK" + (j - 1) + ")";
            myExcelWorksheet.get_Range("AM" + j, misValue).Formula = "=SUM(AM7" + ":AM" + (j - 1) + ")";

            return j;
        }
        #endregion Write To Excel Wssi Division



        #region InsertWssiProductGroupReport
        /// <summary>
        /// InsertWssiProductGroupReport
        /// </summary>
        private void InsertWssiProductGroupReport()
        {
            if (txtWeekNo.Text.Trim().Length > 0)
            {
                GetStockDetails objStock = new GetStockDetails();

                objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
               
                objStock.InsertWssiProductGroupReport();
            }
            else
            {
                lblMessage.Text = "Enter Valid Week No !";
            }
        }
        #endregion InsertWssiProductGroupReport

        #region GenerateWSSIProductReport
        /// <summary>
        /// To generate excel WSSI Product Group Report
        /// </summary>
        private void GenerateWSSIProductReport()
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\WSSI_pg.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtWeekNo.Text.Trim().Length > 0)
            {
                 string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\WSSI_pg.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                dtStock = objStock.GetWSSIProductGroupReport();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = "Summary";

                    xlSheet.get_Range("A1", misValue).Formula = "WSSI-Product Group Report Week No:"+WeekNo;
                    xlSheet.get_Range("A1", misValue).Font.Bold = true;

                    WriteToExcelWssiProductGroup(dtStock,xlSheet);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";
                    btnWssiPg.Visible = true;
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
                //do nothing
            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\WSSI_ProductGroup" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameWssiProductGroup"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();


            //}

            //catch (Exception e)
            // {

            // }

        }

        #endregion GenerateWSSIDivisionReport

        #region Write To Excel Wssi Product Group
        /// <summary>
        /// Write To Excel Wssi Product Group
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelWssiProductGroup(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 3;
            
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductGroup"]) ? dtStock.Rows[i]["ProductGroup"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["ClosingQty"]) ? dtStock.Rows[i]["ClosingQty"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["SoldQty"]) ? dtStock.Rows[i]["SoldQty"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["DcIntake"]) ? dtStock.Rows[i]["DcIntake"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["Adjustments"]) ? dtStock.Rows[i]["Adjustments"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            }

            myExcelWorksheet.get_Range("A" + j, misValue).Formula = "Grand Total";
            BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("B" + j, misValue).EntireRow.Interior.Color = System.Drawing.Color.Yellow;
            BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("B" + j, misValue).EntireRow.Font.Bold = true;

            myExcelWorksheet.get_Range("C" + j, misValue).Formula = "=SUM(C3" + ":C" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            
            myExcelWorksheet.get_Range("D" + j, misValue).Formula = "=SUM(D3" + ":D" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("E" + j, misValue).Formula = "=SUM(E3" + ":E" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=SUM(F3" + ":F" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            return j;
        }
        #endregion Write To Excel Wssi Product Group


        #endregion Methods
    }
}