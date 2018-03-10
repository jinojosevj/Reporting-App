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

#endregion NameSpace

namespace ReportingTool
{
    public partial class RetailKpiTati : System.Web.UI.Page
    {
        public DataTable dtRetailKPI = null;
        #region Events

        #region Page_Load
        /// <summary>
        /// Page_Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                setDefault();
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

            try
            {
                if (rdlRetailKpi.SelectedValue == "1")
                {
                    InsertWeeklySales();
                    InsertDailySales();
                    InsertRetailKpi();

                    //Not Using InsertRetailKpiYear();
                    InsertDsrReport();
                }

                GenerateRetailKPI();
                //Not Using GenerateRetailKPIMonthYear();
                GenerateDsrReport();
            }
            catch(Exception ex)
            {
                lblMessage.Text = ex.ToString();
            }
        }
        #endregion btnGenerate_Click

        #region btnImport_Click
        /// <summary>
        /// btnImport_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnImport_Click(object sender, EventArgs e)
        {
            Boolean fileOK = false;
            Boolean fileFormat = false;
            String Msg = ""; ;
            String path = Server.MapPath("~/FileImport/");
            bool Result = false;
            if (IsPostBack)
            {

                if (fileuploadExcel.HasFile)
                {
                    String fileExtension =
                        System.IO.Path.GetExtension(fileuploadExcel.FileName).ToLower();
                    String[] allowedExtensions = { ".xls", ".xlsx" };
                    for (int i = 0; i < allowedExtensions.Length; i++)
                    {
                        if (fileExtension == allowedExtensions[i])
                        {
                            fileOK = true;
                        }
                    }
                }

                if (fileOK)
                {
                    try
                    {
                        Random rnd = new Random();
                        String fileName = "Sales_Plan_Tati" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                        fileuploadExcel.PostedFile.SaveAs(path
                            + fileName);

                        FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

                        //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)

                        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        //...
                        //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                        DataSet result = excelReader.AsDataSet();
                        //...
                        //4. DataSet - Create column names from first row
                        excelReader.IsFirstRowAsColumnNames = true;
                        result = excelReader.AsDataSet();

                        //5. Data Reader methods
                        //while (excelReader.Read())
                        //{
                        //    excelReader.GetInt32(0);
                        //}

                        DataTable DtSource = result.Tables[0];

                        GetStockDetails objKPI = new GetStockDetails();

                        // for deleting old sales plan
                        for (int i = 0; i < DtSource.Rows.Count; i++)
                        {
                            objKPI.WeekNo = Convert.ToInt32(DtSource.Rows[i]["WeekNo"]);
                            objKPI.PostingDate = Convert.ToDateTime(DtSource.Rows[i]["PostingDate"]);
                            objKPI.Location = DtSource.Rows[i]["StoreCode"].ToString();
                            objKPI.DeleteSalesPlanTati();
                        }


                        // import new sales paln

                        objKPI.DtSource = DtSource;
                        Result = objKPI.ImportSalesPlanTati();
                        //6. Free resources (IExcelDataReader is IDisposable)
                        excelReader.Close();



                        if (Result)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Green;
                            lblMessage.Text = "Successfully Import Sales Plan Data!";
                        }
                        else if (Msg.Length > 0)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = Msg;
                        }
                        else
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = "Failed To Import Sales Plan Data!";

                        }

                    }
                    catch (Exception ex)
                    {
                        lblMessage.ForeColor = System.Drawing.Color.Red;
                        lblMessage.Text = "File could not be uploaded.";
                    }
                }
                else
                {
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "Cannot accept files of this type.";
                }
            }

        }
        #endregion btnImport_Click

        #region btnImportLinear_Click
        /// <summary>
        /// btnImportLinear_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnImportLinear_Click(object sender, EventArgs e)
        {
            Boolean fileOK = false;
            Boolean fileFormat = false;
            String Msg = ""; ;
            String path = Server.MapPath("~/FileImport/");
            bool Result = false;
            if (IsPostBack)
            {

                if (fudLinearCount.HasFile)
                {
                    String fileExtension =
                        System.IO.Path.GetExtension(fudLinearCount.FileName).ToLower();
                    String[] allowedExtensions = { ".xls", ".xlsx" };
                    for (int i = 0; i < allowedExtensions.Length; i++)
                    {
                        if (fileExtension == allowedExtensions[i])
                        {
                            fileOK = true;
                        }
                    }
                }

                if (fileOK)
                {
                    try
                    {
                        Random rnd = new Random();
                        String fileName = "Linear_Count_Tati" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                        fudLinearCount.PostedFile.SaveAs(path
                            + fileName);

                        FileStream stream = File.Open(path + fileName, FileMode.Open, FileAccess.Read);

                        //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)

                        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        //...
                        //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                        DataSet result = excelReader.AsDataSet();
                        //...
                        //4. DataSet - Create column names from first row
                        excelReader.IsFirstRowAsColumnNames = true;
                        result = excelReader.AsDataSet();

                        //5. Data Reader methods
                        //while (excelReader.Read())
                        //{
                        //    excelReader.GetInt32(0);
                        //}

                        DataTable DtSource = result.Tables[0];

                        GetStockDetails objKPI = new GetStockDetails();

                        //eliminate empty rows

                        for (int i = DtSource.Rows.Count - 1; i >= 0; i += -1)
                        {
                            DataRow row = DtSource.Rows[i];
                            if (row[0] == null)
                            {
                                DtSource.Rows.Remove(row);
                            }
                            else if (string.IsNullOrEmpty(row[0].ToString()))
                            {
                                DtSource.Rows.Remove(row);
                            }
                        }

                        // for deleting old Linear count
                        for (int i = 0; i < DtSource.Rows.Count; i++)
                        {
                            objKPI.WeekNo = Convert.ToInt32(DtSource.Rows[i]["WeekNo"]);
                            objKPI.Year = DtSource.Rows[i]["Year"].ToString();
                            objKPI.Location = DtSource.Rows[i]["LocationCode"].ToString();
                            objKPI.CategoryCode = DtSource.Rows[i]["CategoryCode"].ToString();

                            objKPI.DeleteLinearCountTati();
                        }

                        // import new Linear Count
                        DtSource.Columns.Add("CreatedDate");
                        for (int i = 0; i < DtSource.Rows.Count; i++)
                        {
                            DtSource.Rows[i]["CreatedDate"] = DateTime.Now;
                        }
                        objKPI.DtSource = DtSource;
                        Result = objKPI.ImportLinearCountTati();
                        //6. Free resources (IExcelDataReader is IDisposable)
                        excelReader.Close();

                        if (Result)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Green;
                            lblMessage.Text = "Successfully Import Linear Count Data!";
                        }
                        else if (Msg.Length > 0)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = Msg;
                        }
                        else
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = "Failed To Import Linear Count Data!";
                        }

                    }
                    catch (Exception ex)
                    {
                        lblMessage.ForeColor = System.Drawing.Color.Red;
                        lblMessage.Text = "File could not be uploaded.";
                    }
                }
                else
                {
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "Cannot accept files of this type.";
                }
            }
        }
        #endregion btnImportLinear_Click

        #region btnDownload_Click
        /// <summary>
        /// btnDownload_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownload_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameRetailKPI"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownload_Click

        #region btnDownloadDsr_Click
        /// <summary>
        /// btnDownloadDsr_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadDsr_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameDsr"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadDsr_Click

        #endregion Events

        #region Methods

        #region InsertWeeklySales
        /// <summary>
        /// InsertWeeklySales
        /// </summary>
        private void InsertWeeklySales()
        {
            GetStockDetails objRetailKPI = new GetStockDetails();

            objRetailKPI.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
            objRetailKPI.Year = ddlYear.SelectedItem.Text;
            objRetailKPI.InsertWeeklySalesTati();
        }
        #endregion InsertWeeklySales

        #region Insert Daily Sales
        /// <summary>
        /// Insert Daily Sales
        /// </summary>
        private void InsertDailySales()
        {
            GetStockDetails objRetailKPI = new GetStockDetails();

            objRetailKPI.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                      
            objRetailKPI.Country = "JORDAN";
            objRetailKPI.InsertDailySalesTati();

            objRetailKPI.Country = "MYQATAR";
            objRetailKPI.InsertDailySalesTati();

        }
        #endregion Insert Daily Sales

        #region InsertRetailKpi
        /// <summary>
        /// InsertRetailKpi
        /// </summary>
        private void InsertRetailKpi()
        {
            GetStockDetails objRetailKPI = new GetStockDetails();

            objRetailKPI.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.ReportDate = DateTime.ParseExact(txtReportDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

            objRetailKPI.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            objRetailKPI.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
            objRetailKPI.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            objRetailKPI.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

            objRetailKPI.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());

            objRetailKPI.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
            objRetailKPI.Year = ddlYear.SelectedItem.Text;

            objRetailKPI.LYear = (Convert.ToInt32(ddlYear.SelectedItem.Value) - 1) + "-" + ddlYear.SelectedItem.Value;
            objRetailKPI.L2Year = (Convert.ToInt32(ddlYear.SelectedItem.Value) - 2) + "-" + (Convert.ToInt32(ddlYear.SelectedItem.Value) - 1);

            objRetailKPI.InsertRetailKPITati();
            objRetailKPI.InsertRetailKPIByDivisionTati();
        }
        #endregion InsertRetailKpi

        #region InsertRetailKpiYear
        /// <summary>
        /// InsertRetailKpiYear
        /// </summary>
        private void InsertRetailKpiYear()
        {
            GetStockDetails objRetailKPI = new GetStockDetails();

            objRetailKPI.FromDate = DateTime.ParseExact(txtMonthStart.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.ToDate = DateTime.Now.AddDays(-1);


            objRetailKPI.FromDateLY = DateTime.ParseExact(txtMonthStart.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).AddYears(-1);
            objRetailKPI.ToDateLY = DateTime.Now.AddDays(-1).AddYears(-1);
            objRetailKPI.FromDate2LY = DateTime.ParseExact(txtMonthStart.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).AddYears(-2);
            objRetailKPI.ToDate2LY = DateTime.Now.AddDays(-1).AddYears(-2);

            objRetailKPI.FromDateYear = DateTime.ParseExact(txtYearStart.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.ToDateYear = DateTime.Now.AddDays(-1);
            objRetailKPI.FromDateYearLY = DateTime.ParseExact(txtYearStart.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).AddYears(-1);
            objRetailKPI.ToDateYearLY = DateTime.Now.AddDays(-1).AddYears(-1);

            objRetailKPI.FromDateYear2LY = DateTime.ParseExact(txtYearStart.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).AddYears(-2);
            objRetailKPI.ToDateYear2LY = DateTime.Now.AddDays(-1).AddYears(-2);

            objRetailKPI.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            objRetailKPI.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
            objRetailKPI.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            objRetailKPI.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

            objRetailKPI.InsertRetailKpiMonth();
            objRetailKPI.InsertRetailKPIYearByDivision();
        }
        #endregion InsertRetailKpiYear

        #region GenerateRetailKPI
        /// <summary>
        /// To generate excel report for Retail KPI
        /// </summary>
        private void GenerateRetailKPI()
        {

            try
            {

                Excel1.Application myExcelApp;

                Excel1.Workbooks myExcelWorkbooks;

                Excel1.Workbook myExcelWorkbook;


                object misValue = System.Reflection.Missing.Value;

                myExcelApp = new Excel1.Application();

                myExcelApp.Visible = false;

                myExcelWorkbooks = myExcelApp.Workbooks;

                string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\RetailKPITati1.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                GetStockDetails objRetailKPI = new GetStockDetails();

                if (txtFromDate.Text.Trim().Length > 0 && txtToDate.Text.Trim().Length > 0)
                {
                    fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\RetailKPITati1.xlsx";
                    myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                    objRetailKPI.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    objRetailKPI.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                    objRetailKPI.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                    objRetailKPI.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                    objRetailKPI.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                    objRetailKPI.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                    objRetailKPI.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());

                    objRetailKPI.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
                    objRetailKPI.Year = ddlYear.Text.Trim();

                    //------------------Jordan-------------------------------
                    int j;
                    objRetailKPI.Country = "JORDAN";
                    dtRetailKPI = objRetailKPI.GetRetailKPITati();
                    Excel1.Worksheet xlSheetJordan = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetJordan.Name = "JORDAN";
                    xlSheetJordan.get_Range("E3", misValue).Formula = "Retail KPI Dashboard - Jordan - Week " + txtWeekNo.Text.Trim();
                    xlSheetJordan.get_Range("E6", misValue).Formula = DateTime.ParseExact(txtReportDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).DayOfWeek.ToString();

                    j = 8;
                    if (dtRetailKPI.Rows.Count > 0)
                        j = WriteToExcelRetailKPI(dtRetailKPI, xlSheetJordan, false, j);

                    dtRetailKPI = objRetailKPI.GetRetailKPILFLTati();
                    j = 11;
                    j = WriteToExcelRetailKPILFL(dtRetailKPI, xlSheetJordan, j);

                    dtRetailKPI = objRetailKPI.GetRetailKPIByDivisionTati();
                    j = WriteToExcelRetailKPIDivision(dtRetailKPI, xlSheetJordan, j + 4, false);

                    dtRetailKPI = objRetailKPI.GetRetailKPIByDivisionLFLTati();
                    j = 19;
                    j = WriteToExcelRetailKPIDivisionLFL(dtRetailKPI, xlSheetJordan, j);




                    //------------------MY Qatar------------------------------ -

                    objRetailKPI.Country = "MYQATAR";
                    dtRetailKPI = objRetailKPI.GetRetailKPITati();
                    Excel1.Worksheet xlSheetMYQatar = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheetMYQatar.Name = "MYQATAR";
                    xlSheetMYQatar.get_Range("E3", misValue).Formula = "Mellow Yellow Retail KPI Dashboard - Qatar - Week " + txtWeekNo.Text.Trim();
                    xlSheetMYQatar.get_Range("E6", misValue).Formula = DateTime.ParseExact(txtReportDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).DayOfWeek.ToString();

                    j = 8;
                    if (dtRetailKPI.Rows.Count > 0)
                    {
                        j = WriteToExcelRetailKPI(dtRetailKPI, xlSheetMYQatar, false, j);
                        dtRetailKPI = objRetailKPI.GetRetailKPILFLTati();
                        j = 11;
                        j = WriteToExcelRetailKPILFL(dtRetailKPI, xlSheetMYQatar, j);

                        dtRetailKPI = objRetailKPI.GetRetailKPIByDivisionTati();
                        j = WriteToExcelRetailKPIDivision(dtRetailKPI, xlSheetMYQatar, j + 5, false);

                        dtRetailKPI = objRetailKPI.GetRetailKPIByDivisionLFLTati();
                        j = 19;
                        j = WriteToExcelRetailKPIDivisionLFL(dtRetailKPI, xlSheetMYQatar, j);

                    }

                    Random rnd = new Random();
                    string filePath = Server.MapPath(".") + "\\Reports\\KPIBTCF" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                    ViewState["FileNameRetailKPI"] = filePath;
                    myExcelWorkbook.SaveAs(@filePath);

                    myExcelWorkbook.Close();
                    myExcelWorkbooks.Close();
                    btnDownload.Visible = true;
                }
            }
            catch(Exception ex)
            {
                lblMessage.Text = ex.ToString();
            }

        }
        #endregion GenerateRetailKPI

        #region Write To Excel Retail KPI
        /// <summary>
        /// Write To Excel Retail KPI
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPI(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, bool reportType, int j)
        {
            object misValue = System.Reflection.Missing.Value;


            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["StoreCode"]) ? dtStock.Rows[i]["StoreCode"].ToString() : "0";

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesTD"]) ? dtStock.Rows[i]["SalesTD"].ToString() : "0";
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";


                myExcelWorksheet.get_Range("BM" + j, misValue).Formula = (null != dtStock.Rows[i]["SQ.Mtr"]) ? dtStock.Rows[i]["SQ.Mtr"].ToString() : "0";

                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLastWeek"]) ? dtStock.Rows[i]["SalesLastWeek"].ToString() : "0";
                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLY"]) ? dtStock.Rows[i]["SalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPlan"]) ? dtStock.Rows[i]["SalesPlan"].ToString() : "0";
                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales2Year"]) ? dtStock.Rows[i]["Sales2Year"].ToString() : "0";

                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["Cost"]) ? dtStock.Rows[i]["Cost"].ToString() : "0";
                myExcelWorksheet.get_Range("AU" + j, misValue).Formula = (null != dtStock.Rows[i]["CostLY"]) ? dtStock.Rows[i]["CostLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowThisWeek"]) ? dtStock.Rows[i]["FlowThisWeek"].ToString() : "0";
                myExcelWorksheet.get_Range("AW" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowVsLastWeek"]) ? dtStock.Rows[i]["FlowVsLastWeek"].ToString() : "0";

                myExcelWorksheet.get_Range("AX" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowVsLastYear"]) ? dtStock.Rows[i]["FlowVsLastYear"].ToString() : "0";
                myExcelWorksheet.get_Range("AZ" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemsThisWeek"]) ? dtStock.Rows[i]["ItemsThisWeek"].ToString() : "0";
                myExcelWorksheet.get_Range("BA" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemsVsLY"]) ? dtStock.Rows[i]["ItemsVsLY"].ToString() : "0";
                myExcelWorksheet.get_Range("BE" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLWD"]) ? dtStock.Rows[i]["SalesLWD"].ToString() : "0";

                myExcelWorksheet.get_Range("BF" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLYD"]) ? dtStock.Rows[i]["SalesLYD"].ToString() : "0";
                myExcelWorksheet.get_Range("BG" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPlanTD"]) ? dtStock.Rows[i]["SalesPlanTD"].ToString() : "0";


                if (reportType == false)
                {
                    myExcelWorksheet.get_Range("BH" + j, misValue).Formula = (null != dtStock.Rows[i]["VisitorsTW"]) ? dtStock.Rows[i]["VisitorsTW"].ToString() : "0";
                    myExcelWorksheet.get_Range("BI" + j, misValue).Formula = (null != dtStock.Rows[i]["VisitorsLW"]) ? dtStock.Rows[i]["VisitorsLW"].ToString() : "0";
                    myExcelWorksheet.get_Range("BJ" + j, misValue).Formula = (null != dtStock.Rows[i]["VisitorsLY"]) ? dtStock.Rows[i]["VisitorsLY"].ToString() : "0";
                    myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["RankThisWeek"]) ? dtStock.Rows[i]["RankThisWeek"].ToString() : "0";

                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["RankLastWeek"]) ? dtStock.Rows[i]["RankLastWeek"].ToString() : "0";
                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["LFL"]) ? dtStock.Rows[i]["LFL"].ToString() : "0";
                }
            }
            return j;
        }
        #endregion Write To Excel RetailKPI

        #region Write To Excel Retail KPIDivision
        /// <summary>
        /// Write To Excel Retail KPI Division
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPIDivision(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j, bool reportType)
        {
            object misValue = System.Reflection.Missing.Value;
            //int j = 0;

            //switch (Country)
            //{
            //    case "UAE": j = 25;
            //        break;
            //    case "QATAR": j = 28;
            //        break;
            //    case "BAHRAIN": j = 28;
            //        break;
            //    case "OMAN": j = 29;
            //        break;
            //    case "JORDAN": j = 34;
            //        break;
            //}

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["StoreCode"]) ? dtStock.Rows[i]["StoreCode"].ToString() : "0";

                if (reportType == false)
                {
                    myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["LFL"]) ? dtStock.Rows[i]["LFL"].ToString() : "0";
                }
                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["KidsSales"]) ? dtStock.Rows[i]["KidsSales"].ToString() : "0";//Ready
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["MensSales"]) ? dtStock.Rows[i]["MensSales"].ToString() : "0";//Lingery

                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["LadiesSales"]) ? dtStock.Rows[i]["LadiesSales"].ToString() : "0";//Acc
                myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["HomeSales"]) ? dtStock.Rows[i]["HomeSales"].ToString() : "0";//Home
                myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["FootSales"]) ? dtStock.Rows[i]["FootSales"].ToString() : "0";//hygine
                myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["EssentialSales"]) ? dtStock.Rows[i]["EssentialSales"].ToString() : "0";//Services

                myExcelWorksheet.get_Range("AC" + j, misValue).Formula = (null != dtStock.Rows[i]["OtherSales"]) ? dtStock.Rows[i]["OtherSales"].ToString() : "0";//Furniture


                //Last Year


                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["KidsSalesLY"]) ? dtStock.Rows[i]["KidsSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["MensSalesLY"]) ? dtStock.Rows[i]["MensSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["LadiesSalesLY"]) ? dtStock.Rows[i]["LadiesSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["HomeSalesLY"]) ? dtStock.Rows[i]["HomeSalesLY"].ToString() : "0";

                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["FootSalesLY"]) ? dtStock.Rows[i]["FootSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AU" + j, misValue).Formula = (null != dtStock.Rows[i]["EssentialSalesLY"]) ? dtStock.Rows[i]["EssentialSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["OtherSalesLY"]) ? dtStock.Rows[i]["OtherSalesLY"].ToString() : "0";

            }
            return j;
        }
        #endregion Write To Excel Retail KPI Division

        #region Write To Excel Retail KPI LFL
        /// <summary>
        /// Write To Excel Retail KPI LFL
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPILFL(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j)
        {
            object misValue = System.Reflection.Missing.Value;
            //int j = 0;

            //switch (Country)
            //{
            //    case "UAE": j = 24;
            //        break;
            //    case "QATAR": j = 11;
            //        break;
            //    case "BAHRAIN": j = 11;
            //        break;
            //    case "OMAN": j = 12;
            //        break;
            //    case "JORDAN": j = 17;
            //        break;
            //}

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {


                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesTD"]) ? dtStock.Rows[i]["SalesTD"].ToString() : "0";


                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLastWeek"]) ? dtStock.Rows[i]["SalesLastWeek"].ToString() : "0";
                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLY"]) ? dtStock.Rows[i]["SalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPlan"]) ? dtStock.Rows[i]["SalesPlan"].ToString() : "0";

                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales2Year"]) ? dtStock.Rows[i]["Sales2Year"].ToString() : "0";
                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["Cost"]) ? dtStock.Rows[i]["Cost"].ToString() : "0";
                myExcelWorksheet.get_Range("AU" + j, misValue).Formula = (null != dtStock.Rows[i]["CostLY"]) ? dtStock.Rows[i]["CostLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowThisWeek"]) ? dtStock.Rows[i]["FlowThisWeek"].ToString() : "0";

                myExcelWorksheet.get_Range("AW" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowVsLastWeek"]) ? dtStock.Rows[i]["FlowVsLastWeek"].ToString() : "0";
                myExcelWorksheet.get_Range("AX" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowVsLastYear"]) ? dtStock.Rows[i]["FlowVsLastYear"].ToString() : "0";
                myExcelWorksheet.get_Range("AZ" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemsThisWeek"]) ? dtStock.Rows[i]["ItemsThisWeek"].ToString() : "0";
                myExcelWorksheet.get_Range("BA" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemsVsLY"]) ? dtStock.Rows[i]["ItemsVsLY"].ToString() : "0";

                myExcelWorksheet.get_Range("BE" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLWD"]) ? dtStock.Rows[i]["SalesLWD"].ToString() : "0";
                myExcelWorksheet.get_Range("BF" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLYD"]) ? dtStock.Rows[i]["SalesLYD"].ToString() : "0";
                myExcelWorksheet.get_Range("BG" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPlanTD"]) ? dtStock.Rows[i]["SalesPlanTD"].ToString() : "0";

                myExcelWorksheet.get_Range("BM" + j, misValue).Formula = (null != dtStock.Rows[i]["SQ.Mtr"]) ? dtStock.Rows[i]["SQ.Mtr"].ToString() : "0";


            }
            return j;
        }
        #endregion Write To Excel Retail KPI LFL

        #region WriteToExcelRetailKPIDivisionLFL
        /// <summary>
        /// WriteToExcelRetailKPIDivisionLFL
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPIDivisionLFL(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j)
        {
            object misValue = System.Reflection.Missing.Value;
            //int j = 0;

            //switch(Country)
            //{
            //    case "UAE": j = 54;
            //                break;
            //    case "QATAR": j = 32;
            //                break;
            //    case "BAHRAIN": j = 32;
            //                break;
            //    case "OMAN": j = 34;
            //                break;
            //    case "JORDAN": j = 44;
            //                break;
            //}

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["KidsSales"]) ? dtStock.Rows[i]["KidsSales"].ToString() : "0";
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["HomeSales"]) ? dtStock.Rows[i]["HomeSales"].ToString() : "0";
                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["LadiesSales"]) ? dtStock.Rows[i]["LadiesSales"].ToString() : "0";
                myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["MensSales"]) ? dtStock.Rows[i]["MensSales"].ToString() : "0";

                myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["FootSales"]) ? dtStock.Rows[i]["FootSales"].ToString() : "0";
                myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["EssentialSales"]) ? dtStock.Rows[i]["EssentialSales"].ToString() : "0";
                myExcelWorksheet.get_Range("AC" + j, misValue).Formula = (null != dtStock.Rows[i]["OtherSales"]) ? dtStock.Rows[i]["OtherSales"].ToString() : "0";

                //Last Year

                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["KidsSalesLY"]) ? dtStock.Rows[i]["KidsSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["HomeSalesLY"]) ? dtStock.Rows[i]["HomeSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["LadiesSalesLY"]) ? dtStock.Rows[i]["LadiesSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["MensSalesLY"]) ? dtStock.Rows[i]["MensSalesLY"].ToString() : "0";

                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["FootSalesLY"]) ? dtStock.Rows[i]["FootSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AU" + j, misValue).Formula = (null != dtStock.Rows[i]["EssentialSalesLY"]) ? dtStock.Rows[i]["EssentialSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["OtherSalesLY"]) ? dtStock.Rows[i]["OtherSalesLY"].ToString() : "0";
            }
            return j;
        }
        #endregion WriteToExcelRetailKPIDivisionLFL



        #region GenerateRetailKPIMonthYear
        /// <summary>
        /// To generate excel report for Retail KPI Month Year
        /// </summary>
        private void GenerateRetailKPIMonthYear()
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

            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\RetailKPIMonth.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            GetStockDetails objRetailKPI = new GetStockDetails();

            if (txtFromDate.Text.Trim().Length > 0 && txtToDate.Text.Trim().Length > 0)
            {
                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\RetailKPIMonth.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                //------------------UAE-------------------------------

                objRetailKPI.Country = "UAE";
                dtRetailKPI = objRetailKPI.GetRetailKpiMonth();
                Excel1.Worksheet xlSheetUAE = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetUAE.Name = "UAE";

                xlSheetUAE.get_Range("D2", misValue).Formula = "Retail KPI Dashboard - UAE - Week " + txtWeekNo.Text.Trim();

                int j = 0;
                if (dtRetailKPI.Rows.Count > 0)
                    j = WriteToExcelRetailKPIMonth(dtRetailKPI, xlSheetUAE);
                else
                    j = 3;
                dtRetailKPI = objRetailKPI.GetRetailKpiYear();
                j = WriteToExcelRetailKPIYear(dtRetailKPI, xlSheetUAE);

                dtRetailKPI = objRetailKPI.GetRetailKpiYearDivision();
                j = WriteToExcelRetailKPIYearDivision(dtRetailKPI, xlSheetUAE, "UAE");

                //dtRetailKPI = objRetailKPI.GetRetailKPIByDivisionLFL();
                //j = WriteToExcelRetailKPIDivisionLFL(dtRetailKPI, xlSheetUAE, "UAE");


                //------------------Jordan-------------------------------

                objRetailKPI.Country = "JORDAN";
                dtRetailKPI = objRetailKPI.GetRetailKpiMonth();
                Excel1.Worksheet xlSheetJordan = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetJordan.Name = "Jordan";

                xlSheetJordan.get_Range("D2", misValue).Formula = "Retail KPI Dashboard - Jordan - Week " + txtWeekNo.Text.Trim();


                if (dtRetailKPI.Rows.Count > 0)
                    j = WriteToExcelRetailKPIMonth(dtRetailKPI, xlSheetJordan);
                else
                    j = 3;
                dtRetailKPI = objRetailKPI.GetRetailKpiYear();
                j = WriteToExcelRetailKPIYear(dtRetailKPI, xlSheetJordan);

                dtRetailKPI = objRetailKPI.GetRetailKpiYearDivision();
                j = WriteToExcelRetailKPIYearDivision(dtRetailKPI, xlSheetJordan, "JORDAN");


                //------------------Oman-------------------------------

                objRetailKPI.Country = "OMAN";
                dtRetailKPI = objRetailKPI.GetRetailKpiMonth();
                Excel1.Worksheet xlSheetOman = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetOman.Name = "Oman";

                xlSheetOman.get_Range("D2", misValue).Formula = "Retail KPI Dashboard - Oman - Week " + txtWeekNo.Text.Trim();


                if (dtRetailKPI.Rows.Count > 0)
                    j = WriteToExcelRetailKPIMonth(dtRetailKPI, xlSheetOman);
                else
                    j = 3;
                dtRetailKPI = objRetailKPI.GetRetailKpiYear();
                j = WriteToExcelRetailKPIYear(dtRetailKPI, xlSheetOman);

                dtRetailKPI = objRetailKPI.GetRetailKpiYearDivision();
                j = WriteToExcelRetailKPIYearDivision(dtRetailKPI, xlSheetOman, "OMAN");



                //------------------Bahrain-------------------------------

                objRetailKPI.Country = "Bahrain";
                dtRetailKPI = objRetailKPI.GetRetailKpiMonth();
                Excel1.Worksheet xlSheetBahrain = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheetBahrain.Name = "Bahrain";

                xlSheetBahrain.get_Range("D2", misValue).Formula = "Retail KPI Dashboard - Bahrain - Week " + txtWeekNo.Text.Trim();


                if (dtRetailKPI.Rows.Count > 0)
                    j = WriteToExcelRetailKPIMonth(dtRetailKPI, xlSheetBahrain);
                else
                    j = 3;
                dtRetailKPI = objRetailKPI.GetRetailKpiYear();
                j = WriteToExcelRetailKPIYear(dtRetailKPI, xlSheetBahrain);

                dtRetailKPI = objRetailKPI.GetRetailKpiYearDivision();
                j = WriteToExcelRetailKPIYearDivision(dtRetailKPI, xlSheetBahrain, "BAHRAIN");


                //------------------Qatar-------------------------------

                objRetailKPI.Country = "Qatar";
                dtRetailKPI = objRetailKPI.GetRetailKpiMonth();
                Excel1.Worksheet xlSheetQatar = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheetQatar.Name = "Qatar";

                xlSheetQatar.get_Range("D2", misValue).Formula = "Retail KPI Dashboard - Qatar - Week " + txtWeekNo.Text.Trim();


                if (dtRetailKPI.Rows.Count > 0)
                    j = WriteToExcelRetailKPIMonth(dtRetailKPI, xlSheetQatar);
                else
                    j = 3;
                dtRetailKPI = objRetailKPI.GetRetailKpiYear();
                j = WriteToExcelRetailKPIYear(dtRetailKPI, xlSheetQatar);

                dtRetailKPI = objRetailKPI.GetRetailKpiYearDivision();
                j = WriteToExcelRetailKPIYearDivision(dtRetailKPI, xlSheetQatar, "QATAR");


                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\RetailKPIMonthYear_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                ViewState["FileNameRetailKPIYear"] = filePath;
                myExcelWorkbook.SaveAs(@filePath);

                myExcelWorkbook.Close();
                myExcelWorkbooks.Close();
               // btnDownloadYear.Visible = true;
            }

        }
        #endregion GenerateRetailKPIMonthYear

        #region Write To Excel Retail KPI Month
        /// <summary>
        /// Write To Excel Retail KPI Month
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPIMonth(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 8;

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["StoreCode"]) ? dtStock.Rows[i]["StoreCode"].ToString() : "0";
                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["LFL"]) ? dtStock.Rows[i]["LFL"].ToString() : "0";
                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLY"]) ? dtStock.Rows[i]["SalesLY"].ToString() : "0";

                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPlan"]) ? dtStock.Rows[i]["SalesPlan"].ToString() : "0";
                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesL2Y"]) ? dtStock.Rows[i]["SalesL2Y"].ToString() : "0";
                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowTM"]) ? dtStock.Rows[i]["FlowTM"].ToString() : "0";
                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowLY"]) ? dtStock.Rows[i]["FlowLY"].ToString() : "0";

            }
            return j;
        }
        #endregion Write To Excel Retail KPI Month

        #region Write To Excel Retail KPI Year
        /// <summary>
        /// Write To Excel Retail KPI Year
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPIYear(DataTable dtStock, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 8;

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesLY"]) ? dtStock.Rows[i]["SalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AW" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPlan"]) ? dtStock.Rows[i]["SalesPlan"].ToString() : "0";
                myExcelWorksheet.get_Range("AX" + j, misValue).Formula = (null != dtStock.Rows[i]["Salesvs2yr"]) ? dtStock.Rows[i]["Salesvs2yr"].ToString() : "0";

                myExcelWorksheet.get_Range("AY" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowTY"]) ? dtStock.Rows[i]["FlowTY"].ToString() : "0";
                myExcelWorksheet.get_Range("AZ" + j, misValue).Formula = (null != dtStock.Rows[i]["FlowLY"]) ? dtStock.Rows[i]["FlowLY"].ToString() : "0";
                myExcelWorksheet.get_Range("BA" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemsTY"]) ? dtStock.Rows[i]["ItemsTY"].ToString() : "0";
                myExcelWorksheet.get_Range("BB" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemsLY"]) ? dtStock.Rows[i]["ItemsLY"].ToString() : "0";

            }
            return j;
        }
        #endregion Write To Excel Retail KPI Year

        #region Write To Excel Retail KPI Year Division
        /// <summary>
        /// Write To Excel Retail KPI Year Division
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelRetailKPIYearDivision(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string Country)
        {
            object misValue = System.Reflection.Missing.Value;
            int j = 0;

            switch (Country)
            {
                case "UAE": j = 31;
                    break;
                case "QATAR": j = 17;
                    break;
                case "BAHRAIN": j = 17;
                    break;
                case "OMAN": j = 18;
                    break;
                case "JORDAN": j = 22;
                    break;
            }

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["StoreCode"]) ? dtStock.Rows[i]["StoreCode"].ToString() : "0";
                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["LFL"]) ? dtStock.Rows[i]["LFL"].ToString() : "0";
                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["LadiesSales"]) ? dtStock.Rows[i]["LadiesSales"].ToString() : "0";
                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["MensSales"]) ? dtStock.Rows[i]["MensSales"].ToString() : "0";

                myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["KidsSales"]) ? dtStock.Rows[i]["KidsSales"].ToString() : "0";
                myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["HomeSales"]) ? dtStock.Rows[i]["HomeSales"].ToString() : "0";
                myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["FootSales"]) ? dtStock.Rows[i]["FootSales"].ToString() : "0";
                myExcelWorksheet.get_Range("X" + j, misValue).Formula = (null != dtStock.Rows[i]["EssentialSales"]) ? dtStock.Rows[i]["EssentialSales"].ToString() : "0";

                myExcelWorksheet.get_Range("AB" + j, misValue).Formula = (null != dtStock.Rows[i]["OtherSales"]) ? dtStock.Rows[i]["OtherSales"].ToString() : "0";

                //Last Year
                myExcelWorksheet.get_Range("AP" + j, misValue).Formula = (null != dtStock.Rows[i]["LadiesSalesLY"]) ? dtStock.Rows[i]["LadiesSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AQ" + j, misValue).Formula = (null != dtStock.Rows[i]["MensSalesLY"]) ? dtStock.Rows[i]["MensSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AR" + j, misValue).Formula = (null != dtStock.Rows[i]["KidsSalesLY"]) ? dtStock.Rows[i]["KidsSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AS" + j, misValue).Formula = (null != dtStock.Rows[i]["HomeSalesLY"]) ? dtStock.Rows[i]["HomeSalesLY"].ToString() : "0";

                myExcelWorksheet.get_Range("AT" + j, misValue).Formula = (null != dtStock.Rows[i]["FootSalesLY"]) ? dtStock.Rows[i]["FootSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AU" + j, misValue).Formula = (null != dtStock.Rows[i]["EssentialSalesLY"]) ? dtStock.Rows[i]["EssentialSalesLY"].ToString() : "0";
                myExcelWorksheet.get_Range("AV" + j, misValue).Formula = (null != dtStock.Rows[i]["OtherSalesLY"]) ? dtStock.Rows[i]["OtherSalesLY"].ToString() : "0";

            }
            return j;
        }
        #endregion Write To Excel Retail KPI Year Division

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


        #region setDefault
        /// <summary>
        /// setDefault
        /// </summary>
        private void setDefault()
        {
            DateTime now = DateTime.Now;
            var monthStartDate = new DateTime(now.Year, now.Month, 1);
            var monthEndDate = monthStartDate.AddMonths(1).AddDays(-1);

            var thisWeekStart = now.AddDays(-(int)now.DayOfWeek);
            var thisWeekEnd = thisWeekStart.AddDays(6);
            var lastWeekStart = thisWeekStart.AddDays(-7);
            var lastWeekEnd = thisWeekStart.AddDays(-1);

            var reportDate = DateTime.Now.AddDays(-1);

            DateTime yearStartDate = new DateTime(DateTime.Now.Year, 1, 1);

            txtYearStart.Text = yearStartDate.ToString("dd/MM/yyyy");

            txtMonthStart.Text = monthStartDate.ToString("dd/MM/yyyy");
            txtMonthEnd.Text = monthEndDate.ToString("dd/MM/yyyy");
            txtReportDate.Text = reportDate.ToString("dd/MM/yyyy");

            if (DateTime.Now.DayOfWeek.ToString() == "Sunday" )
            {
                txtFromDate.Text = lastWeekStart.ToString("dd/MM/yyyy");
                txtToDate.Text = reportDate.ToString("dd/MM/yyyy");
            }
            else
            {
                txtFromDate.Text = thisWeekStart.ToString("dd/MM/yyyy");
                txtToDate.Text = reportDate.ToString("dd/MM/yyyy");
            }

            GetStockDetails objWeeks = new GetStockDetails();

            if (DateTime.Now.DayOfWeek.ToString() == "Sunday" )
            {
                objWeeks.FromDate = lastWeekStart;
                objWeeks.ToDate = lastWeekEnd;
            }
            else
            {
                objWeeks.FromDate = thisWeekStart;
                objWeeks.ToDate = thisWeekEnd;
            }

            DataTable dt = objWeeks.GetWeekDetailsTati();

            if (dt.Rows.Count > 0)
            {
                if (!IsPostBack)
                {
                    ddlYear.SelectedItem.Text = dt.Rows[0]["YEAR"].ToString();
                    ddlYear.SelectedItem.Value = dt.Rows[0]["YEAR"].ToString().Substring(0, 4);
                    txtWeekNo.Text = dt.Rows[0]["WeekNo"].ToString();
                }
            }

        }
        #endregion setDefault

        #region Insert Dsr Report
        /// <summary>
        /// Insert Dsr Report
        /// </summary>
        private void InsertDsrReport()
        {
            GetStockDetails objRetailKPI = new GetStockDetails();

            objRetailKPI.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objRetailKPI.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            objRetailKPI.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());

            objRetailKPI.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            objRetailKPI.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
            objRetailKPI.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());


            objRetailKPI.Location = "4728";
            objRetailKPI.InsertDsrReportTati();
            objRetailKPI.Location = "4729";
            objRetailKPI.InsertDsrReportTati();

            objRetailKPI.Location = "4731";
            objRetailKPI.InsertDsrReportTati();


            objRetailKPI.InsertDsrDivisionTati();
        }
        #endregion Insert Dsr Report

        #region Generate DSR Report
        /// <summary>
        /// Generate DSR Report
        /// </summary>
        private void GenerateDsrReport()
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

            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\DsrReportTati.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            GetStockDetails objRetailKPI = new GetStockDetails();

            if (txtFromDate.Text.Trim().Length > 0 && txtToDate.Text.Trim().Length > 0)
            {
                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\DsrReportTati.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


                objRetailKPI.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objRetailKPI.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objRetailKPI.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objRetailKPI.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());

                objRetailKPI.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objRetailKPI.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objRetailKPI.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());
                objRetailKPI.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());

                objRetailKPI.Year = ddlYear.Text.Trim();

                objRetailKPI.Location = "4728";
                dtRetailKPI = objRetailKPI.GetDsrReportTati();
                Excel1.Worksheet xlSheet4728 = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet4728.Name = "4728";
                int j = 34;
                WriteToExcelDsr(dtRetailKPI, xlSheet4728, j, "4728");
                dtRetailKPI = objRetailKPI.GetDsrDivisionTati();
                j = 20;
                WriteToExcelDsrDivision(dtRetailKPI, xlSheet4728, j, "4728");

                objRetailKPI.Location = "4729";
                dtRetailKPI = objRetailKPI.GetDsrReportTati();
                Excel1.Worksheet xlSheet4729 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheet4729.Name = "4729";
                j = 34;
                WriteToExcelDsr(dtRetailKPI, xlSheet4729, j, "4729");
                dtRetailKPI = objRetailKPI.GetDsrDivisionTati();
                j = 20;
                WriteToExcelDsrDivision(dtRetailKPI, xlSheet4729, j, "4729");


                objRetailKPI.Location = "4731";
                dtRetailKPI = objRetailKPI.GetDsrReportTati();
                Excel1.Worksheet xlSheet4731 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheet4731.Name = "4731";
                j = 34;
                WriteToExcelDsr(dtRetailKPI, xlSheet4731, j, "4731");
                dtRetailKPI = objRetailKPI.GetDsrDivisionTati();
                j = 20;
                WriteToExcelDsrDivision(dtRetailKPI, xlSheet4731, j, "4731");


                
                Random rnd = new Random();
                string filePath = Server.MapPath(".") + "\\Reports\\DSR_Tati_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                ViewState["FileNameDsr"] = filePath;
                myExcelWorkbook.SaveAs(@filePath);

                myExcelWorkbook.Close();
                myExcelWorkbooks.Close();
                btnDownloadDsr.Visible = true;
            }

        }
        #endregion GenerateRetailKPI

        #region Write To Excel Dsr
        /// <summary>
        /// Write To Excel Dsr 
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelDsr(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j, string StoreCode)
        {
            object misValue = System.Reflection.Missing.Value;


            myExcelWorksheet.get_Range("G" + 12, misValue).Formula = StoreCode + "  Divisional Sales Report For Week :- " + txtWeekNo.Text.ToString();// +" From  " + txtFromDate.Text.ToString().Trim() + "  To  " + txtToDate.Text.ToString().Trim();
            myExcelWorksheet.get_Range("I" + 33, misValue).Formula = txtFromDate.Text.Trim();

            myExcelWorksheet.get_Range("I" + 32, misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).DayOfWeek.ToString();

            string order = "A";
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                string dtOrder = (null != dtStock.Rows[i]["Order"]) ? dtStock.Rows[i]["Order"].ToString() : "-";
                if (order != dtOrder)
                {
                    order = dtOrder;
                    Excel1.Range Line1 = (Excel1.Range)myExcelWorksheet.Rows[j];
                    Line1.Insert();
                    
                    myExcelWorksheet.get_Range("G" + j, "AA" + j).Interior.Color = System.Drawing.Color.Black;
                    j++;
                }

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemCategoryCode"]) ? dtStock.Rows[i]["ItemCategoryCode"].ToString() : "-";
                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemCategoryDesc"]) ? dtStock.Rows[i]["ItemCategoryDesc"].ToString() : "0";
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Day1Sales"]) ? dtStock.Rows[i]["Day1Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["Day1Mix"]) ? dtStock.Rows[i]["Day1Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["Day2Sales"]) ? dtStock.Rows[i]["Day2Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["Day2Mix"]) ? dtStock.Rows[i]["Day2Mix"].ToString() : "0";
                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["Day3Sales"]) ? dtStock.Rows[i]["Day3Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["Day3Mix"]) ? dtStock.Rows[i]["Day3Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["Day4Sales"]) ? dtStock.Rows[i]["Day4Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["Day4Mix"]) ? dtStock.Rows[i]["Day4Mix"].ToString() : "0";
                myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["Day5Sales"]) ? dtStock.Rows[i]["Day5Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["Day5Mix"]) ? dtStock.Rows[i]["Day5Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["Day6Sales"]) ? dtStock.Rows[i]["Day6Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["Day6Mix"]) ? dtStock.Rows[i]["Day6Mix"].ToString() : "0";
                myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["Day7Sales"]) ? dtStock.Rows[i]["Day7Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("V" + j, misValue).Formula = (null != dtStock.Rows[i]["Day7Mix"]) ? dtStock.Rows[i]["Day7Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["Total"]) ? dtStock.Rows[i]["Total"].ToString() : "0";
                myExcelWorksheet.get_Range("X" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalMix"]) ? dtStock.Rows[i]["TotalMix"].ToString() : "0";
                myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["LinearCount"]) ? dtStock.Rows[i]["LinearCount"].ToString() : "0";
                myExcelWorksheet.get_Range("Z" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPerLinear"]) ? dtStock.Rows[i]["SalesPerLinear"].ToString() : "0";

                myExcelWorksheet.get_Range("AA" + j, misValue).Formula = (null != dtStock.Rows[i]["LinearMix"]) ? dtStock.Rows[i]["LinearMix"].ToString() : "0";

                Excel1.Range Line = (Excel1.Range)myExcelWorksheet.Rows[j + 1];
                Line.Insert();
            }

            Excel1.Range Line2 = (Excel1.Range)myExcelWorksheet.Rows[j];
            Line2.Delete();

            Excel1.Range Line3 = (Excel1.Range)myExcelWorksheet.Rows[j];
            Line3.Delete();

            Excel1.Range Line4 = (Excel1.Range)myExcelWorksheet.Rows[j];
            Line4.Delete();

            return j;
        }
        #endregion Write To Excel Dsr


        #region Write To Excel Dsr New
        /// <summary>
        /// Write To Excel Dsr New
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelDsrNew(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j, string StoreCode)
        {
            object misValue = System.Reflection.Missing.Value;


            myExcelWorksheet.get_Range("G" + 12, misValue).Formula = StoreCode + "  Divisional Sales Report For Week :- " + txtWeekNo.Text.ToString();// +" From  " + txtFromDate.Text.ToString().Trim() + "  To  " + txtToDate.Text.ToString().Trim();
            myExcelWorksheet.get_Range("I" + 33, misValue).Formula = txtFromDate.Text.Trim();

            myExcelWorksheet.get_Range("I" + 32, misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).DayOfWeek.ToString();

            string order = "A";
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                string dtOrder = (null != dtStock.Rows[i]["Order"]) ? dtStock.Rows[i]["Order"].ToString() : "-";
                if (order != dtOrder)
                {
                    order = dtOrder;
                    Excel1.Range Line1 = (Excel1.Range)myExcelWorksheet.Rows[j];
                    Line1.Insert();

                    myExcelWorksheet.get_Range("G" + j, "AA" + j).Interior.Color = System.Drawing.Color.Black;
                    j++;
                }

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["DivisionCode"]) ? dtStock.Rows[i]["DivisionCode"].ToString() : "-";
                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["DivisionDesc"]) ? dtStock.Rows[i]["DivisionDesc"].ToString() : "0";
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemCategoryCode"]) ? dtStock.Rows[i]["ItemCategoryCode"].ToString() : "-";
                myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemCategoryDesc"]) ? dtStock.Rows[i]["ItemCategoryDesc"].ToString() : "0";

                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductCode"]) ? dtStock.Rows[i]["ProductCode"].ToString() : "-";
                myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["ProductDesc"]) ? dtStock.Rows[i]["ProductDesc"].ToString() : "0";
                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["FamilyCode"]) ? dtStock.Rows[i]["FamilyCode"].ToString() : "-";
                myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["FamilyDesc"]) ? dtStock.Rows[i]["FamilyDesc"].ToString() : "0";

                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["Day1Sales"]) ? dtStock.Rows[i]["Day1Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["Day1Mix"]) ? dtStock.Rows[i]["Day1Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["Day2Sales"]) ? dtStock.Rows[i]["Day2Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["Day2Mix"]) ? dtStock.Rows[i]["Day2Mix"].ToString() : "0";
                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["Day3Sales"]) ? dtStock.Rows[i]["Day3Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["Day3Mix"]) ? dtStock.Rows[i]["Day3Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["Day4Sales"]) ? dtStock.Rows[i]["Day4Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["Day4Mix"]) ? dtStock.Rows[i]["Day4Mix"].ToString() : "0";
                myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["Day5Sales"]) ? dtStock.Rows[i]["Day5Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["Day5Mix"]) ? dtStock.Rows[i]["Day5Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["Day6Sales"]) ? dtStock.Rows[i]["Day6Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["Day6Mix"]) ? dtStock.Rows[i]["Day6Mix"].ToString() : "0";
                myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["Day7Sales"]) ? dtStock.Rows[i]["Day7Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("V" + j, misValue).Formula = (null != dtStock.Rows[i]["Day7Mix"]) ? dtStock.Rows[i]["Day7Mix"].ToString() : "0";

                myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["Total"]) ? dtStock.Rows[i]["Total"].ToString() : "0";
                myExcelWorksheet.get_Range("X" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalMix"]) ? dtStock.Rows[i]["TotalMix"].ToString() : "0";
                myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["LinearCount"]) ? dtStock.Rows[i]["LinearCount"].ToString() : "0";
                myExcelWorksheet.get_Range("Z" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPerLinear"]) ? dtStock.Rows[i]["SalesPerLinear"].ToString() : "0";

                myExcelWorksheet.get_Range("AA" + j, misValue).Formula = (null != dtStock.Rows[i]["LinearMix"]) ? dtStock.Rows[i]["LinearMix"].ToString() : "0";

                Excel1.Range Line = (Excel1.Range)myExcelWorksheet.Rows[j + 1];
                Line.Insert();
            }

            Excel1.Range Line2 = (Excel1.Range)myExcelWorksheet.Rows[j];
            Line2.Delete();

            Excel1.Range Line3 = (Excel1.Range)myExcelWorksheet.Rows[j];
            Line3.Delete();

            Excel1.Range Line4 = (Excel1.Range)myExcelWorksheet.Rows[j];
            Line4.Delete();

            return j;
        }
        #endregion WriteToExcelDsrNew




        #region Write To Excel Dsr Division
        /// <summary>
        /// Write To Excel Dsr Division
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelDsrDivision(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j, string StoreCode)
        {
            object misValue = System.Reflection.Missing.Value;


            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["DivisionCode"]) ? dtStock.Rows[i]["DivisionCode"].ToString() : "-";
                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtStock.Rows[i]["DivisionDesc"]) ? dtStock.Rows[i]["DivisionDesc"].ToString() : "0";
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";
                myExcelWorksheet.get_Range("J" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesMix"]) ? dtStock.Rows[i]["SalesMix"].ToString() : "0";

                myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["QtySold"]) ? dtStock.Rows[i]["QtySold"].ToString() : "0";
                myExcelWorksheet.get_Range("L" + j, misValue).Formula = (null != dtStock.Rows[i]["QtyMix"]) ? dtStock.Rows[i]["QtyMix"].ToString() : "0";
                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["ASP"]) ? dtStock.Rows[i]["ASP"].ToString() : "0";
                myExcelWorksheet.get_Range("N" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalClsQty"]) ? dtStock.Rows[i]["TotalClsQty"].ToString() : "0";


                myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalClsQtyMix"]) ? dtStock.Rows[i]["TotalClsQtyMix"].ToString() : "0";
                myExcelWorksheet.get_Range("P" + j, misValue).Formula = (null != dtStock.Rows[i]["TotalStockValue"]) ? dtStock.Rows[i]["TotalStockValue"].ToString() : "0";
                myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["StockValueMix"]) ? dtStock.Rows[i]["StockValueMix"].ToString() : "0";
                myExcelWorksheet.get_Range("R" + j, misValue).Formula = (null != dtStock.Rows[i]["LinearCount"]) ? dtStock.Rows[i]["LinearCount"].ToString() : "0";


                myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["SalesPerLinear"]) ? dtStock.Rows[i]["SalesPerLinear"].ToString() : "0";
                myExcelWorksheet.get_Range("T" + j, misValue).Formula = (null != dtStock.Rows[i]["LinearMix"]) ? dtStock.Rows[i]["LinearMix"].ToString() : "0";

            }
            return j;
        }
        #endregion Write To Excel Dsr Division


        #endregion Methods
    }
}