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

using Microsoft.Office.Core;
using System.Diagnostics;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using Excel;
using System.Drawing;
//using Microsoft.Office.Interop.Excel;
using Test.DAL;
using ReportingTool.BAL;

#endregion NameSpace
namespace ReportingTool
{
    public partial class MISReports : System.Web.UI.Page
    {
        public const int UpdateTableProcessId = 1;
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

      
        #region ddlCountry_SelectedIndexChanged
        /// <summary>
        /// ddlCountry_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlCountry_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlCountry.SelectedItem.Text != "MME")
            {
                GetStockDetails ObjStock = new GetStockDetails();
                ObjStock.Country = ddlCountry.SelectedItem.Text;
                DataTable dt = ObjStock.GetStoreByCountry();
                ddlLocation.DataSource = dt;
                ddlLocation.DataMember = "LocationCode";
                ddlLocation.DataValueField = "LocationCode";
                ddlLocation.DataBind();
            }
        }

        #endregion ddlCountry_SelectedIndexChanged


        #region btnGenerate_Click
        /// <summary>
        /// btnGenerate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnGenerate_Click(object sender, EventArgs e)
        {
            GetStockDetails ObjStock = new GetStockDetails();

            ObjStock.IntType = Convert.ToInt32(ddlType.SelectedItem.Value);
            ObjStock.Country = ddlCountry.SelectedItem.Value;
            ObjStock.Location = ddlLocation.SelectedItem.Value;
            ObjStock.FromDate = txtFromDate.Text.Length>0? Convert.ToDateTime(txtFromDate.Text) : default(DateTime);
            ObjStock.ToDate= txtToDate.Text.Length > 0 ? Convert.ToDateTime(txtToDate.Text) : default(DateTime);

            ObjStock.LineCode7 = txtLinecode7.Text.ToString().Trim();
            ObjStock.DivisionCode = ddlDivision.SelectedItem.Value.Trim();


            DataTable dt = null;
            if (ddlType.SelectedItem.Value == "1" || ddlType.SelectedItem.Value == "4" || ddlType.SelectedItem.Value == "5")
            {
                if (ddlDivision.Text.Trim() == "Select" )
                {
                    lblMessage.Text = "Please Put Proper Filters(Division)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }
            else if (ddlType.SelectedItem.Value == "2" || ddlType.SelectedItem.Value == "8" || ddlType.SelectedItem.Value == "9" || ddlType.SelectedItem.Value == "16")
            {
                if (txtFromDate.Text.Trim().Length == 0 && txtToDate.Text.Trim().Length == 0 && ddlLocation.Text=="Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(FromDate,ToDate and StoreNo)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }

            else if (ddlType.SelectedItem.Value == "10" )
            {
                if (txtFromDate.Text.Trim().Length == 0 && txtToDate.Text.Trim().Length == 0 && ddlCountry.Text == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(FromDate,ToDate and Country)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }

            else if (ddlType.SelectedItem.Value == "3" || ddlType.SelectedItem.Value == "19")
            {
                if (txtLinecode7.Text.Trim().Length == 0 && ddlLocation.Text == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(LineCode7 and StoreNo)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }
            else if (ddlType.SelectedItem.Value == "6" || ddlType.SelectedItem.Value == "7" )
            {
                if (ddlLocation.Text.Trim() == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(Location)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }

            else if (ddlType.SelectedItem.Value == "11"|| ddlType.SelectedItem.Value == "12")
            {
                if (ddlCountry.Text.Trim() == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(Country)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }

            else if (ddlType.SelectedItem.Value == "13" )
            {
                if (ddlCountry.Text.Trim() == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(Country)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    ImportUnitPrice();
                }
            }
            else if (ddlType.SelectedItem.Value == "14")
            {
                if (GetProcessStatus())
                {
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "Tables Are Locked By Another User,Try Again Later";
                }
                else
                {
                    UpdateTables();
                }

            }

            else if (ddlType.SelectedItem.Value == "15")
            {
                if (ddlCountry.Text.Trim() == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(Country)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    ImportSalesPrice();
                }
            }
            else if (ddlType.SelectedItem.Value == "17")
            {
                if (ddlCountry.Text.Trim() == "Select")
                {
                    lblMessage.Text = "Please Put Proper Filters(Country)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    GenerateMarkdownReport();
                }
            }
            else if (ddlType.SelectedItem.Value == "18")
            {
                if (txtToDate.Text.Trim().Length == 0)
                {
                    lblMessage.Text = "Please Put Proper Filters(To Date)";
                    lblMessage.ForeColor = Color.Red;
                    btnDownload.Visible = false;
                }
                else
                {
                    dt = ObjStock.GetMISReports();
                }
            }


            if (dt != null && dt.Rows.Count > 0)
            {

                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\" + ddlType.SelectedItem.Text +"_"+ddlCountry.SelectedItem.Value+ "_" + rnd.Next() + ".csv";
                ViewState["FileName"] = filePath;
                StreamWriter sw = new StreamWriter(@filePath, false);

                ExportToCsv(dt, sw);
                sw.Close();

                btnDownload.Visible = true;
                lblMessage.Text = "Report Generated";
                lblMessage.ForeColor = Color.Green;
            }
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
            string fileName = ViewState["FileName"].ToString();
            FileDownload(fileName);
        }
        #endregion btnDownload_Click

        #endregion Events

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

            if ((file.Extension == ".CSV") || (file.Extension == ".csv") || (file.Extension == ".xlsx"))
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

        #region ImportUnitPrice
        /// <summary>
        /// ImportUnitPrice
        /// </summary>
        private void ImportUnitPrice()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "UnitPrice" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            GetStockDetails ObjImport = new GetStockDetails();

            Boolean Result = false;
            int Count = 0;
            for (int i=0;i<dtImport.Rows.Count;i++)
            {
                ObjImport.UnitPrice = Convert.ToDecimal(dtImport.Rows[i]["UnitPrice"]);
                ObjImport.Country   = ddlCountry.SelectedItem.Value;
                ObjImport.ItemNo = dtImport.Rows[i]["ItemNo"].ToString();
                Result=ObjImport.UpdateUnitPrice();
                if(Result)
                {
                    Count++;
                }

            }

            if (Result)
            {
                lblMessage.Text = Count.ToString()+" Rows Updated!";
                lblMessage.ForeColor = Color.Green;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion ImportUnitPrice


        #region ImportSalesPrice
        /// <summary>
        /// ImportSalesPrice
        /// </summary>
        private void ImportSalesPrice()
        {
            String path = Server.MapPath("~/FileImport/");
            Random rnd = new Random();
            String fileName = "SalesPrice" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".csv";
            fileuploadExcel.PostedFile.SaveAs(path
                + fileName);

            //FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

            DataTable dtImport = CsvReader.ReadCSVFile(path + fileName, true);

            GetStockDetails ObjImport = new GetStockDetails();

            Boolean Result = false;
            int Count = 0;
            for (int i = 0; i < dtImport.Rows.Count; i++)
            {
                ObjImport.UnitPrice = Convert.ToDecimal(dtImport.Rows[i]["UnitPrice"]);
                ObjImport.UnitCost = Convert.ToDecimal(dtImport.Rows[i]["UnitCostLCY"]);
                ObjImport.LineAmount = Convert.ToDecimal(dtImport.Rows[i]["LineAmount"]);
                ObjImport.Country = ddlCountry.SelectedItem.Value;

                ObjImport.ItemNo = dtImport.Rows[i]["ItemNo"].ToString();
                ObjImport.DocNo = dtImport.Rows[i]["DocumentNo"].ToString();
                Result = ObjImport.UpdateSalesLine();
                if (Result)
                {
                    Count++;
                }

            }

            if (Result)
            {
                lblMessage.Text = Count.ToString() + " Rows Updated!";
                lblMessage.ForeColor = Color.Green;
            }
            else
            {
                lblMessage.Text = "Faild!";
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion ImportSalesPrice


        #region UpdateTables
        /// <summary>
        /// Update Tables
        /// </summary>
        private void UpdateTables()
        {
            GetStockDetails ObjStock = new GetStockDetails();

            ObjStock.ItemOperationType =  false;
            ObjStock.ILEOperationType = 1;
            ObjStock.ValueOperationType = 2;
            ObjStock.FootFallOperationType = true;
            ObjStock.TransactionOperationType = true;

            bool Result = ObjStock.UpdateTables();


            if (Result == true)
            {
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Successfuly Completed.";
            }
            else
            {
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Tables Updation Failed.";
            }
        }
        #endregion UpdateTables

        #region GetProcessStatus
        /// <summary>
        /// Get Process Status
        /// </summary>
        private bool GetProcessStatus()
        {
            GetStockDetails objStock = new GetStockDetails();
            objStock.ProcessStatusId = UpdateTableProcessId;
            DataTable dtStatus = objStock.GetProcessStatus();
            bool Flag = Convert.ToBoolean(dtStatus.Rows[0]["Flag"]);

            return Flag;
        }
        #endregion GetProcessStatus


        #region Generate Markdown Report
        /// <summary>
        /// Generate Markdown Report
        /// </summary>
        private void GenerateMarkdownReport()
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

            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\Markdown.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            GetStockDetails ObjMarkdown = new GetStockDetails();

             fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\Markdown.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

               DateTime dtStart = new DateTime();
               DateTime dtEnd = new DateTime();

               dtStart = DateTime.Now;
              

                ObjMarkdown.InsertMarkDown();

                ObjMarkdown.ReportType = "1"; 
                ObjMarkdown.Location = "MME";
                DataTable dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetMMES = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetMMES.Name = "MME-Season";
                WriteToExcelMarkdown(dtMarkdown, xlSheetMMES, "1");


                ObjMarkdown.ReportType = "2";
                ObjMarkdown.Location = "MMEF";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetMMEF = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetMMEF.Name = "MME-Family";
                WriteToExcelMarkdown(dtMarkdown, xlSheetMMEF, "2");


                ObjMarkdown.ReportType = "3";
                ObjMarkdown.Location = "UAE";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetUAE = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetUAE.Name = "UAE";
                WriteToExcelMarkdown(dtMarkdown, xlSheetUAE, "3");


                ObjMarkdown.ReportType = "3";
                ObjMarkdown.Location = "JOR";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetJOR = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheetJOR.Name = "JOR";
                WriteToExcelMarkdown(dtMarkdown, xlSheetJOR, "3");


                ObjMarkdown.ReportType = "3";
                ObjMarkdown.Location = "OMAN";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetOMAN = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheetOMAN.Name = "OMAN";
                WriteToExcelMarkdown(dtMarkdown, xlSheetOMAN, "3");

                ObjMarkdown.ReportType = "3";
                ObjMarkdown.Location = "BAH";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetBAH = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheetBAH.Name = "BAH";
                WriteToExcelMarkdown(dtMarkdown, xlSheetBAH, "3");


                ObjMarkdown.ReportType = "3";
                ObjMarkdown.Location = "QAT";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetQAT = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                xlSheetQAT.Name = "QAT";
                WriteToExcelMarkdown(dtMarkdown, xlSheetQAT, "3");

                ObjMarkdown.ReportType = "3";
                ObjMarkdown.Location = "KSA";
                dtMarkdown = ObjMarkdown.GetMarkdown();
                Excel1.Worksheet xlSheetKSA = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                xlSheetKSA.Name = "KSA";
                WriteToExcelMarkdown(dtMarkdown, xlSheetKSA, "3");

                dtEnd = DateTime.Now;
                
                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\Markdown_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                ViewState["FileName"] = filePath;
                myExcelWorkbook.SaveAs(@filePath);

                lblMessage.Text = "Report Generated  Total Time:-  "+Math.Round((dtEnd-dtStart).TotalMinutes).ToString()+"M :"+ Math.Round((((dtEnd-dtStart).TotalSeconds)%60)).ToString()+"S";
                lblMessage.ForeColor = Color.Green;
                myExcelWorkbook.Close();
                myExcelWorkbooks.Close();
                btnDownload.Visible = true;
            

        }
        #endregion Generate Markdown Report

        #region Write To Excel Markdown 
        /// <summary>
        /// Write To Excel Markdown 
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelMarkdown(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string ReportType)
        {
            object misValue = System.Reflection.Missing.Value;

            //myExcelWorksheet.get_Range("G" + 12, misValue).Formula = StoreCode + "  Divisional Sales Report For Week :- " + txtWeekNo.Text.ToString();// +" From  " + txtFromDate.Text.ToString().Trim() + "  To  " + txtToDate.Text.ToString().Trim();
          
            for (int i = 0; i < dtStock.Rows.Count; i++)
            {


                if (ReportType == "1")
                {

                    if (dtStock.Rows[i]["Location"].ToString() == "zzz")
                    {
                        myExcelWorksheet.get_Range("A" + (i + 2), misValue).Formula = "Total";

                        myExcelWorksheet.get_Range("A" + (i + 2), "M" + (i + 2)).Interior.Color = System.Drawing.Color.Yellow; 
                    }
                    else
                    {
                       myExcelWorksheet.get_Range("A" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Location"]) ? dtStock.Rows[i]["Location"].ToString() : "-";
                    }

                    BorderAround(myExcelWorksheet.get_Range("A" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("B" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["SeasonCode"]) ? dtStock.Rows[i]["SeasonCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("C" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["TotalQty"]) ? dtStock.Rows[i]["TotalQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("D" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_FP"]) ? dtStock.Rows[i]["Qty_FP"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_FP%"]) ? dtStock.Rows[i]["Qty_FP%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("F" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_25"]) ? dtStock.Rows[i]["Qty_25"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("G" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_25%"]) ? dtStock.Rows[i]["Qty_25%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("H" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_50"]) ? dtStock.Rows[i]["Qty_50"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_50%"]) ? dtStock.Rows[i]["Qty_50%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("J" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_75"]) ? dtStock.Rows[i]["Qty_75"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("K" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_75%"]) ? dtStock.Rows[i]["Qty_75%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("L" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_7599"]) ? dtStock.Rows[i]["Qty_7599"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("L" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_7599%"]) ? dtStock.Rows[i]["Qty_7599%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }


                if (ReportType == "2")
                {

                    if (dtStock.Rows[i]["Location"].ToString() == "zzz")
                    {
                        myExcelWorksheet.get_Range("A" + (i + 2), misValue).Formula = "Total";

                        myExcelWorksheet.get_Range("A" + (i + 2), "N" + (i + 2)).Interior.Color = System.Drawing.Color.Yellow;
                    }
                    else
                    {
                        myExcelWorksheet.get_Range("A" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Location"]) ? dtStock.Rows[i]["Location"].ToString() : "-";
                    }

                   // myExcelWorksheet.get_Range("A" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Location"]) ? dtStock.Rows[i]["Location"].ToString() : "-";
                    BorderAround(myExcelWorksheet.get_Range("A" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("B" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["ItemFamilyCode"]) ? dtStock.Rows[i]["ItemFamilyCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["FamilyDescription"]) ? dtStock.Rows[i]["FamilyDescription"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("D" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["TotalQty"]) ? dtStock.Rows[i]["TotalQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_FP"]) ? dtStock.Rows[i]["Qty_FP"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("F" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_FP%"]) ? dtStock.Rows[i]["Qty_FP%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_25"]) ? dtStock.Rows[i]["Qty_25"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("H" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_25%"]) ? dtStock.Rows[i]["Qty_25%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_50"]) ? dtStock.Rows[i]["Qty_50"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("J" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_50%"]) ? dtStock.Rows[i]["Qty_50%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_75"]) ? dtStock.Rows[i]["Qty_75"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("L" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_75%"]) ? dtStock.Rows[i]["Qty_75%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("L" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_7599"]) ? dtStock.Rows[i]["Qty_7599"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("N" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_7599%"]) ? dtStock.Rows[i]["Qty_7599%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("N" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }

                if (ReportType == "3")
                {
                    myExcelWorksheet.get_Range("A" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Location"]) ? dtStock.Rows[i]["Location"].ToString() : "-";
                    BorderAround(myExcelWorksheet.get_Range("A" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("B" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["SeasonCode"]) ? dtStock.Rows[i]["SeasonCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["ItemFamilyCode"]) ? dtStock.Rows[i]["ItemFamilyCode"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("D" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["FamilyDescription"]) ? dtStock.Rows[i]["FamilyDescription"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("D" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["TotalQty"]) ? dtStock.Rows[i]["TotalQty"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("F" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_FP"]) ? dtStock.Rows[i]["Qty_FP"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("F" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("G" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_FP%"]) ? dtStock.Rows[i]["Qty_FP%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("H" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_25"]) ? dtStock.Rows[i]["Qty_25"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("H" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_25%"]) ? dtStock.Rows[i]["Qty_25%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("J" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_50"]) ? dtStock.Rows[i]["Qty_50"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("J" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_50%"]) ? dtStock.Rows[i]["Qty_50%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("L" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_75"]) ? dtStock.Rows[i]["Qty_75"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("L" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_75%"]) ? dtStock.Rows[i]["Qty_75%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    myExcelWorksheet.get_Range("N" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_7599"]) ? dtStock.Rows[i]["Qty_7599"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("N" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("O" + (i + 2), misValue).Formula = (null != dtStock.Rows[i]["Qty_7599%"]) ? dtStock.Rows[i]["Qty_7599%"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + (i + 2), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }
                //Excel1.Range Line = (Excel1.Range)myExcelWorksheet.Rows[i + 1];
                //Line.Insert();
            }

         
        }
        #endregion Write To Excel Markdown

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