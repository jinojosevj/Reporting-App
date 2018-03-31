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
    public partial class InventoryReport : System.Web.UI.Page
    {
        public DataTable dtInventory = null;
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
            ViewState["FileNameInventory"] = null;
            
            GenerateInventoryReport();
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
            string filename = ViewState["FileNameInventory"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownload_Click

        #endregion Events

        #region Methods

        #region GenerateInventoryReport
        /// <summary>
        /// To generate excel report for Inventory
        /// </summary>
        private void GenerateInventoryReport()
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




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\InventoryReport.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objInventory = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtFromDate.Text.Trim().Length > 0 && txtToDate.Text.Trim().Length > 0)
            {
                // string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\InventoryReport.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                objInventory.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objInventory.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                dtInventory = objInventory.GetInventoryReport();

                if (dtInventory.Rows.Count > 0)
                {
                    xlSheet.Name = "Summery";
                    xlSheet.get_Range("A1", misValue).Formula =  "Inventory Report :-"+ txtFromDate.Text.Trim()+ " To "+txtToDate.Text.Trim();
                    BorderAround(xlSheet.get_Range("A1", misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    xlSheet.get_Range("A1", misValue).EntireRow.Font.Bold = true;
                    xlSheet.get_Range("A1", misValue).EntireRow.Font.Size=14;
                    
                    WriteToExcelInventoryReport(dtInventory, xlSheet);

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
                //do nothing
            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\InvenoryReport_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameInventory"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();


            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateInventoryReport
        
        #region Write To Excel Inventory Report
        /// <summary>
        /// Write To Excel Inventory Report
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelInventoryReport(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 3;
            int k = 0;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Inventory"]) ? dtStock.Rows[i]["Inventory"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["0400"]) ? dtStock.Rows[i]["0400"].ToString(): "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["0400S%"]) ? dtStock.Rows[i]["0400S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["0401"]) ? dtStock.Rows[i]["0401"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["0401S%"]) ? dtStock.Rows[i]["0401S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["0402"]) ? dtStock.Rows[i]["0402"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["0402S%"]) ? dtStock.Rows[i]["0402S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["0403"]) ? dtStock.Rows[i]["0403"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("H" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["0403S%"]) ? dtStock.Rows[i]["0403S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["0404"]) ? dtStock.Rows[i]["0404"].ToString()  : "0";
                BorderAround(myExcelWorksheet.get_Range("J" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["0404S%"]) ? dtStock.Rows[i]["0404S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["0405"]) ? dtStock.Rows[i]["0405"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("L" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["0405S%"]) ? dtStock.Rows[i]["0405S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("M" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["0406"]) ? dtStock.Rows[i]["0406"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("N" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                
                myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["0406S%"]) ? dtStock.Rows[i]["0406S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("O" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["0407"]) ? dtStock.Rows[i]["0407"].ToString()  : "0";
                BorderAround(myExcelWorksheet.get_Range("P" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["0407S%"]) ? dtStock.Rows[i]["0407S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("Q" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["0408"]) ? dtStock.Rows[i]["0408"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("R" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["0408S%"]) ? dtStock.Rows[i]["0408S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("S" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["0409"]) ? dtStock.Rows[i]["0409"].ToString()  : "0";
                BorderAround(myExcelWorksheet.get_Range("T" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["0409S%"]) ? dtStock.Rows[i]["0409S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("U" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("V" + j, misValue).Formula = (null != dtStock.Rows[i]["0410"]) ? dtStock.Rows[i]["0410"].ToString(): "0";
                BorderAround(myExcelWorksheet.get_Range("V" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["0410S%"]) ? dtStock.Rows[i]["0410S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("W" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("X" + j, misValue).Formula = (null != dtStock.Rows[i]["0411"]) ? dtStock.Rows[i]["0411"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("X" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["0411S%"]) ? dtStock.Rows[i]["0411S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("Y" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("Z" + j, misValue).Formula = (null != dtStock.Rows[i]["0414"]) ? dtStock.Rows[i]["0414"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("Z" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AA" + j, misValue).Formula = (null != dtStock.Rows[i]["0414S%"]) ? dtStock.Rows[i]["0414S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AA" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AB" + j, misValue).Formula = (null != dtStock.Rows[i]["0415"]) ? dtStock.Rows[i]["0415"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AB" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AC" + j, misValue).Formula = (null != dtStock.Rows[i]["0415S%"]) ? dtStock.Rows[i]["0415S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AC" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AD" + j, misValue).Formula = (null != dtStock.Rows[i]["0416"]) ? dtStock.Rows[i]["0416"].ToString()  : "0";
                BorderAround(myExcelWorksheet.get_Range("AD" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AE" + j, misValue).Formula = (null != dtStock.Rows[i]["0416S%"]) ? dtStock.Rows[i]["0416S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AE" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AF" + j, misValue).Formula = (null != dtStock.Rows[i]["0417"]) ? dtStock.Rows[i]["0417"].ToString(): "0";
                BorderAround(myExcelWorksheet.get_Range("AF" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AG" + j, misValue).Formula = (null != dtStock.Rows[i]["0417S%"]) ? dtStock.Rows[i]["0417S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AG" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

              
                myExcelWorksheet.get_Range("AH" + j, misValue).Formula = (null != dtStock.Rows[i]["0412"]) ? dtStock.Rows[i]["0412"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AH" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AI" + j, misValue).Formula = (null != dtStock.Rows[i]["0412S%"]) ? dtStock.Rows[i]["0412S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AI" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                myExcelWorksheet.get_Range("AJ" + j, misValue).Formula = (null != dtStock.Rows[i]["0418"]) ? dtStock.Rows[i]["0418"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AJ" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AK" + j, misValue).Formula = (null != dtStock.Rows[i]["0418S%"]) ? dtStock.Rows[i]["0418S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AK" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AL" + j, misValue).Formula = (null != dtStock.Rows[i]["0419"]) ? dtStock.Rows[i]["0419"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AL" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AM" + j, misValue).Formula = (null != dtStock.Rows[i]["0419S%"]) ? dtStock.Rows[i]["0419S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AM" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AN" + j, misValue).Formula = (null != dtStock.Rows[i]["0421"]) ? dtStock.Rows[i]["0421"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AN" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AO" + j, misValue).Formula = (null != dtStock.Rows[i]["0421S%"]) ? dtStock.Rows[i]["0421S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AO" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["0422"]) ? dtStock.Rows[i]["0422"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AP" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["0422S%"]) ? dtStock.Rows[i]["0422S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AQ" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["0423"]) ? dtStock.Rows[i]["0423"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AR" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["0423S%"]) ? dtStock.Rows[i]["0423S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AS" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["0424"]) ? dtStock.Rows[i]["0424"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AT" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AU" + j, misValue).Formula = (null != dtStock.Rows[i]["0424S%"]) ? dtStock.Rows[i]["0424S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AU" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

               
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["0425"]) ? dtStock.Rows[i]["0425"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AV" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AW" + j, misValue).Formula = (null != dtStock.Rows[i]["0425S%"]) ? dtStock.Rows[i]["0425S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AW" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AX" + j, misValue).Formula = (null != dtStock.Rows[i]["0426"]) ? dtStock.Rows[i]["0426"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AX" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AY" + j, misValue).Formula = (null != dtStock.Rows[i]["0426S%"]) ? dtStock.Rows[i]["0426S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AY" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("AZ" + j, misValue).Formula = (null != dtStock.Rows[i]["0420"]) ? dtStock.Rows[i]["0420"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("AZ" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("BA" + j, misValue).Formula = (null != dtStock.Rows[i]["0420S%"]) ? dtStock.Rows[i]["0420S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("BA" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("BB" + j, misValue).Formula = (null != dtStock.Rows[i]["0427"]) ? dtStock.Rows[i]["0427"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("BB" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("BC" + j, misValue).Formula = (null != dtStock.Rows[i]["0427S%"]) ? dtStock.Rows[i]["0427S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("BC" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("BD" + j, misValue).Formula = (null != dtStock.Rows[i]["0428"]) ? dtStock.Rows[i]["0428"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("BD" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("BE" + j, misValue).Formula = (null != dtStock.Rows[i]["0428S%"]) ? dtStock.Rows[i]["0428S%"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("BE" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                k = i;
            }

            myExcelWorksheet.get_Range("A" + j, misValue).Formula = "Grand Total";
            BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("A" + j, "BC" + j).Interior.Color = System.Drawing.Color.Yellow;
            myExcelWorksheet.get_Range("B" + j, misValue).Formula = "=SUM(B3" + ":B" + (j - 1) + ")";
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

            myExcelWorksheet.get_Range("G" + j, misValue).Formula = "=SUM(G3" + ":G" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("H" + j, misValue).Formula = "=SUM(H3" + ":H" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("H" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
           
            myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=SUM(I3" + ":I" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("J" + j, misValue).Formula = "=SUM(J3" + ":J" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("J" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("K" + j, misValue).Formula = "=SUM(K3" + ":K" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("L" + j, misValue).Formula = "=SUM(L3" + ":L" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("L" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("M" + j, misValue).Formula = "=SUM(M3" + ":M" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("M" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("N" + j, misValue).Formula = "=SUM(N3" + ":N" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("N" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("O" + j, misValue).Formula = "=SUM(O3" + ":O" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("O" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("P" + j, misValue).Formula = "=SUM(P3" + ":P" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("P" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("Q" + j, misValue).Formula = "=SUM(Q3" + ":Q" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("Q" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("R" + j, misValue).Formula = "=SUM(R3" + ":R" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("R" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("S" + j, misValue).Formula = "=SUM(S3" + ":S" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("S" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("T" + j, misValue).Formula = "=SUM(T3" + ":T" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("T" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("U" + j, misValue).Formula = "=SUM(U3" + ":U" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("U" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("V" + j, misValue).Formula = "=SUM(V3" + ":V" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("V" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("W" + j, misValue).Formula = "=SUM(W3" + ":W" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("W" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("X" + j, misValue).Formula = "=SUM(X3" + ":X" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("X" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("Y" + j, misValue).Formula = "=SUM(Y3" + ":Y" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("Y" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("Z" + j, misValue).Formula = "=SUM(Z3" + ":Z" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("Z" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AA" + j, misValue).Formula = "=SUM(AA3" + ":AA" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AA" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AB" + j, misValue).Formula = "=SUM(AB3" + ":AB" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AB" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AC" + j, misValue).Formula = "=SUM(AC3" + ":AC" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AC" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AD" + j, misValue).Formula = "=SUM(AD3" + ":AD" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AD" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AE" + j, misValue).Formula = "=SUM(AE3" + ":AE" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AE" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AF" + j, misValue).Formula = "=SUM(AF3" + ":AF" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AF" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AG" + j, misValue).Formula = "=SUM(AG3" + ":AG" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AG" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AH" + j, misValue).Formula = "=SUM(AH3" + ":AH" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AH" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AI" + j, misValue).Formula = "=SUM(AI3" + ":AI" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AI" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AJ" + j, misValue).Formula = "=SUM(AJ3" + ":AJ" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AJ" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AK" + j, misValue).Formula = "=SUM(AK3" + ":AK" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AK" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AL" + j, misValue).Formula = "=SUM(AL3" + ":AL" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AL" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AM" + j, misValue).Formula = "=SUM(AM3" + ":AM" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AM" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AN" + j, misValue).Formula = "=SUM(AN3" + ":AN" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AN" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AO" + j, misValue).Formula = "=SUM(AO3" + ":AO" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AO" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AP" + j, misValue).Formula = "=SUM(AP3" + ":AP" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AP" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = "=SUM(AQ3" + ":AQ" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AQ" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("AR" + j, misValue).Formula = "=SUM(AR3" + ":AR" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AR" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AS" + j, misValue).Formula = "=SUM(AS3" + ":AS" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AS" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


            myExcelWorksheet.get_Range("AT" + j, misValue).Formula = "=SUM(AT3" + ":AT" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AT" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AU" + j, misValue).Formula = "=SUM(AU3" + ":AU" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AU" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AV" + j, misValue).Formula = "=SUM(AV3" + ":AV" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AV" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AW" + j, misValue).Formula = "=SUM(AW3" + ":AW" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AW" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            
            myExcelWorksheet.get_Range("AX" + j, misValue).Formula = "=SUM(AX3" + ":AX" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AX" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AY" + j, misValue).Formula = "=SUM(AY3" + ":AY" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AY" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("AZ" + j, misValue).Formula = "=SUM(AZ3" + ":AZ" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("AZ" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("BA" + j, misValue).Formula = "=SUM(BA3" + ":BA" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("BA" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("BB" + j, misValue).Formula = "=SUM(BB3" + ":BB" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("BB" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("BC" + j, misValue).Formula = "=SUM(BC3" + ":BC" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("BC" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("BD" + j, misValue).Formula = "=SUM(BD3" + ":BD" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("BD" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            myExcelWorksheet.get_Range("BE" + j, misValue).Formula = "=SUM(BE3" + ":BE" + (j - 1) + ")";
            BorderAround(myExcelWorksheet.get_Range("BE" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            return j;
        }
        #endregion Write To Excel Inventory Report

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