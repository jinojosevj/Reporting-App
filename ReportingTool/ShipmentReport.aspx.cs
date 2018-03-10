#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Test.BAL;
using Excel1 = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.IO;
using ReportingTool.BAL;
#endregion NameSpace

namespace ReportingTool
{
    public partial class ShipmentReport : System.Web.UI.Page
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

        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            GenerateReport();
        }
        #endregion btnGenerate_Click

        #region btnDownloadShip_Click
        /// <summary>
        /// btnDownloadShip_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadShip_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileName"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadShip_Click

        #endregion Events
        
        #region Methods
        
        #region GenerateReport
        /// <summary>
        /// To generate excel report 
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ShipmentReport.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            TatiBAL objStock = new TatiBAL();

            //objStock.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ShipmentReportOne.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;
                dtStock = objStock.GetProfitAndLossTati();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                    WriteToExcel(dtStock, xlSheet, location);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnDownloadShip.Visible = true;
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

                objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetSummary.Name = "Summary";
                objStock.Location = "Summary";
                dtStock = objStock.GetShipmentReport();
                WriteToExcel(dtStock, xlSheetSummary, "Summary");


                Excel1.Worksheet xlSheet4728 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheet4728.Name = "4728";
                objStock.Location = "4728";
                dtStock = objStock.GetShipmentReport();
                WriteToExcel(dtStock, xlSheet4728, "4728");


                Excel1.Worksheet xlSheet4729 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheet4729.Name = "4729";
                objStock.Location = "4729";
                dtStock = objStock.GetShipmentReport();
                WriteToExcel(dtStock, xlSheet4729, "4729");


                Excel1.Worksheet xlSheet4731 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet4731.Name = "4731";
                objStock.Location = "4731";
                dtStock = objStock.GetShipmentReport();
                WriteToExcel(dtStock, xlSheet4731, "4731");


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownloadShip.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\ShipmentReportTati_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

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
        /// WriteToExcel
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcel(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string location)
        {
            object misValue = System.Reflection.Missing.Value;

            // myExcelWorksheet.get_Range("A1", misValue).Formula = location;
           // myExcelWorksheet.get_Range("C1", misValue).Formula = location + " - Profit And Loss Report For  " + ddlMonth.SelectedItem.Text.ToString() + " - " + ddlYear.SelectedItem.Value.ToString();
            //BorderAround(myExcelWorksheet.get_Range("A2", misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("A1", misValue).Font.Bold = true;

            int j = 8;
            
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {



                if (location == "Summary")
                {

                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = i+1;
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = txtWeekNo.Text.Trim();
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalAmt"]) ? dtStock.Rows[i]["TotalAmt"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["ExpectedQty"]) ? dtStock.Rows[i]["ExpectedQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                   
                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = "=(E" + j + "+H" + j + "+M" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["ShortQty"]) ? dtStock.Rows[i]["ShortQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=(H" + j + "/E" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["ShortAmt"]) ? dtStock.Rows[i]["ShortAmt"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                   
                    myExcelWorksheet.get_Range("K" + j, misValue).Formula = "=(J" + j + "/D" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["ExcessQty"]) ? dtStock.Rows[i]["ExcessQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("N" + j, misValue).Formula = "=(M" + j + "/E" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("N" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["ExcessAmt"]) ? dtStock.Rows[i]["ExcessAmt"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    
                    myExcelWorksheet.get_Range("P" + j, misValue).Formula = "=(O" + j + "/D" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("P" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                                        
                    myExcelWorksheet.get_Range("R" + j, misValue).Formula = "=(H" + j + "+M" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("R" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = "=(R" + j + "/E" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("S" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("T" + j, misValue).Formula = "=(J" + j + "+O" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("T" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    
                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = "=(T" + j + "/D" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("U" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                }
                else
                {
                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = i + 1;
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Posting Date"]) ? Convert.ToDateTime(dtStock.Rows[i]["Posting Date"]).ToString("dd/MM/yyyy") : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + j, misValue).Formula = txtWeekNo.Text.Trim();
                    BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["Vendor Order No_"]) ? dtStock.Rows[i]["Vendor Order No_"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                   
                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["Vendor Shipment No_"]) ? dtStock.Rows[i]["Vendor Shipment No_"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["Vendor Invoice No_"]) ? dtStock.Rows[i]["Vendor Invoice No_"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalAmt"]) ? dtStock.Rows[i]["TotalAmt"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["ExpectedQty"]) ? dtStock.Rows[i]["ExpectedQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = "=(H" + j + "+K" + j + "+P" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["ShortQty"]) ? dtStock.Rows[i]["ShortQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("L" + j, misValue).Formula = "=(K" + j + "/H" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("L" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["ShortAmt"]) ? dtStock.Rows[i]["ShortAmt"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("N" + j, misValue).Formula = "=(M" + j + "/G" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("N" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["ExcessQty"]) ? dtStock.Rows[i]["ExcessQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("P" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                   
                    myExcelWorksheet.get_Range("Q" + j, misValue).Formula = "=(P" + j + "/H" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("Q" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["ExcessAmt"]) ? dtStock.Rows[i]["ExcessAmt"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("R" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = "=(R" + j + "/G" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("S" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = "=(K" + j + "+P" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("U" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    
                    myExcelWorksheet.get_Range("V" + j, misValue).Formula = "=(U" + j + "/H" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("V" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("W" + j, misValue).Formula = "=(M" + j + "+R" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("W" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("X" + j, misValue).Formula = "=(W" + j + "/G" + j + ")";
                    BorderAround(myExcelWorksheet.get_Range("X" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
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