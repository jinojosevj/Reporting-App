
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
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using Excel;

#endregion NameSpace


namespace Test
{
    public partial class Visitors : System.Web.UI.Page
    {

        public DataTable dtVisitors = null;

        public const int VisitorReportProcessId = 3; 

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
                      String fileName = "Visitor_Data" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
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

                        for(int i=0;i<DtSource.Rows.Count;i++)
                        {
                           
                            if (   DtSource.Rows[i]["StoreNo"].ToString() == "0409" 
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0410"
                                ||DtSource.Rows[i]["StoreNo"].ToString() == "0414"
                                ||DtSource.Rows[i]["StoreNo"].ToString() == "0415"

                                || DtSource.Rows[i]["StoreNo"].ToString() == "0416" 
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0417"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0419"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0418"

                                || DtSource.Rows[i]["StoreNo"].ToString() == "0421"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0412"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0422"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0423"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0424"

                                || DtSource.Rows[i]["StoreNo"].ToString() == "0425"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0426"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0427"
                                || DtSource.Rows[i]["StoreNo"].ToString() == "0428"
                                )
                            {
                                fileFormat = true; 
                            }
                            else
                            {
                                fileFormat = false;
                                Msg = "Store Number Is Not Correct";
                            }


                            if (   DtSource.Rows[i]["Entrance"].ToString() == "0409-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0410-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0414-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0415-01"

                                || DtSource.Rows[i]["Entrance"].ToString() == "0415-02"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0415-03"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0416-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0417-01"

                                || DtSource.Rows[i]["Entrance"].ToString() == "0419-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0419-02"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0419-03"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0418-01"

                                || DtSource.Rows[i]["Entrance"].ToString() == "0421-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0412-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0412-02"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0412-03"

                                || DtSource.Rows[i]["Entrance"].ToString() == "0422-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0423-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0424-01"

                                || DtSource.Rows[i]["Entrance"].ToString() == "0425-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0425-02"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0425-03"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0426-01"

                                || DtSource.Rows[i]["Entrance"].ToString() == "0427-01"
                                || DtSource.Rows[i]["Entrance"].ToString() == "0428-01"
                                )
                            {
                                fileFormat = true;
                            }
                            else
                            {
                                fileFormat = false;
                                Msg = "Entrance Name Is Not Correct";
                            }

                        }

                        if (fileFormat)
                        {
                            GetStockDetails objVisitor = new GetStockDetails();
                            objVisitor.DtSource = DtSource;
                            Result=objVisitor.BulkInsert();
                            //6. Free resources (IExcelDataReader is IDisposable)
                            excelReader.Close();
                            
                        }
                       
                        if(Result)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Green;
                            lblMessage.Text = "Successfully Import The Data!";
                        }
                        else if(Msg.Length>0)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = Msg;
                        }
                        else
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = "Failed To Import The Data!";

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

        #region btnReport_Click
        /// <summary>
        /// btnReport_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnReport_Click(object sender, EventArgs e)
        {
            ViewState["FileName"] = null;
            ViewState["FileNameWeekly"] = null;

            if (txtDate.Text.Trim().Length > 0)
            {
                if (GetProcessStatus())
                {
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "Tables Are Locked By Another User,Try Again Later";
                }
                else
                {

                    GetStockDetails objStock = new GetStockDetails();
                    objStock.ProcessStatusFlag = true;
                    objStock.ProcessStatusId = VisitorReportProcessId;
                    objStock.UpdateProcessStatus();


                   InsertVisitorsReport();
                   // lblMessage.Text = "Done";
                   GenerateReport();
                   GenerateVisitorsWeeklyReport();

                   //string weekDay= DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dddd");
                   //if (weekDay == "Saturday")
                   //   GenerateVisitorsVsSales();

                    objStock.ProcessStatusFlag = false;
                    objStock.ProcessStatusId = VisitorReportProcessId;
                    objStock.UpdateProcessStatus();
                }
            }
            else
            {
                  lblMessage.ForeColor = System.Drawing.Color.Red;
                  lblMessage.Text = "Please Enter Posting Date";
            }

            Page.ClientScript.RegisterStartupScript(this.GetType(), "CallMyFunction", "$('#btnReport').Show();", true);
            //DateTime test_date = DateTime.Now;
            //double week_of_year = Math.Ceiling(Convert.ToDouble(test_date.DayOfYear) / 7);

            //double week_of_year = GetWeekNumber(test_date);


        }
        #endregion btnReport_Click


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


        #region btnDownloadWeekly_Click
        /// <summary>
        /// btnDownloadWeekly_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadWeekly_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameWeekly"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadWeekly_Click


        #region btnVisitorsVsSales_Click
        /// <summary>
        /// btnVisitorsVsSales_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnVisitorsVsSales_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameVisitorsVsSales"].ToString();
            FileDownload(filename);
        }
        #endregion btnVisitorsVsSales_Click

        #endregion Events


        #region Methods

        #region GenerateReport
        /// <summary>
        /// To generate excel report for Visitors counting
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

            
            String fileName = HttpContext.Current.Server.MapPath(".")+"\\Template\\HourlyReport.xlsx";

           

             myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.ActiveSheet;

            //myExcelWorkbooks.Close();

           // myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);

            Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            
            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                dtVisitors = objStock.GetVisitorsReport(location);

                if (dtVisitors.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                    WriteToExcel(dtVisitors, xlSheet, location);

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
                    btnDownload.Visible = false;
                }

            }

            else
            {


                //Excel.Sheets xlSheets = myExcelWorkbook.Sheets as Excel.Sheets;

                Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheet0409.Name = "0409";
                dtVisitors = objStock.GetVisitorsReport("0409");
                WriteToExcel(dtVisitors, xlSheet0409, "0409");

                Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheet0410.Name = "0410";
                dtVisitors = objStock.GetVisitorsReport("0410");
                WriteToExcel(dtVisitors, xlSheet0410, "0410");

                Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheet0414.Name = "0414";
                dtVisitors = objStock.GetVisitorsReport("0414");
                WriteToExcel(dtVisitors, xlSheet0414, "0414");

                Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet0415.Name = "0415";
                dtVisitors = objStock.GetVisitorsReport("0415");
                WriteToExcel(dtVisitors, xlSheet0415, "0415");

                Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheet0416.Name = "0416";
                dtVisitors = objStock.GetVisitorsReport("0416");
                WriteToExcel(dtVisitors, xlSheet0416, "0416");

                Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheet0417.Name = "0417";
                dtVisitors = objStock.GetVisitorsReport("0417");
                WriteToExcel(dtVisitors, xlSheet0417, "0417");

                Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                xlSheet0419.Name = "0419";
                dtVisitors = objStock.GetVisitorsReport("0419");
                WriteToExcel(dtVisitors, xlSheet0419, "0419");

                Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                xlSheet0418.Name = "0418";
                dtVisitors = objStock.GetVisitorsReport("0418");
                WriteToExcel(dtVisitors, xlSheet0418, "0418");

                Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                xlSheet0421.Name = "0421";
                dtVisitors = objStock.GetVisitorsReport("0421");
                WriteToExcel(dtVisitors, xlSheet0421, "0421");

                Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                xlSheet0412.Name = "0412";
                dtVisitors = objStock.GetVisitorsReport("0412");
                WriteToExcel(dtVisitors, xlSheet0412, "0412");

                Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                xlSheet0422.Name = "0422";
                dtVisitors = objStock.GetVisitorsReport("0422");
                WriteToExcel(dtVisitors, xlSheet0422, "0422");

                Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                xlSheet0423.Name = "0423";
                dtVisitors = objStock.GetVisitorsReport("0423");
                if (dtVisitors.Rows.Count > 0)
                {
                    WriteToExcel(dtVisitors, xlSheet0423, "0423");
                }

                Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                xlSheet0424.Name = "0424";
                dtVisitors = objStock.GetVisitorsReport("0424");
                if (dtVisitors.Rows.Count > 0)
                {
                    WriteToExcel(dtVisitors, xlSheet0424, "0424");
                }

                Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                xlSheet0425.Name = "0425";
                dtVisitors = objStock.GetVisitorsReport("0425");
                if (dtVisitors.Rows.Count > 0)
                {
                    WriteToExcel(dtVisitors, xlSheet0425, "0425");
                }

                Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                xlSheet0426.Name = "0426";
                dtVisitors = objStock.GetVisitorsReport("0426");
                if (dtVisitors.Rows.Count > 0)
                {
                    WriteToExcel(dtVisitors, xlSheet0426, "0426");
                }

                Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                xlSheet0427.Name = "0427";
                dtVisitors = objStock.GetVisitorsReport("0427");
                if (dtVisitors.Rows.Count > 0)
                {
                    WriteToExcel(dtVisitors, xlSheet0427, "0427");
                }

                Excel1.Worksheet xlSheet0428 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                xlSheet0428.Name = "0428";
                dtVisitors = objStock.GetVisitorsReport("0428");
                if (dtVisitors.Rows.Count > 0)
                {
                    WriteToExcel(dtVisitors, xlSheet0428, "0428");
                }

                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownload.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\Visitors_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

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
        private void WriteToExcel(DataTable dtVisitors, Excel1.Worksheet myExcelWorksheet, string location)
        {
            object misValue = System.Reflection.Missing.Value;

            myExcelWorksheet.get_Range("J4", misValue).Formula = txtDate.Text.Trim().ToString() +"- Hourly Report "+location;

            for (int i = 0; i < dtVisitors.Rows.Count; i++)
            {
                myExcelWorksheet.get_Range("M" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["Visitor"] && dtVisitors.Rows[i]["Visitor"].ToString().Length>0) ? dtVisitors.Rows[i]["Visitor"].ToString() : "0";
                myExcelWorksheet.get_Range("N" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["VistorvsLW"] && dtVisitors.Rows[i]["VistorvsLW"].ToString().Length>0) ? dtVisitors.Rows[i]["VistorvsLW"].ToString() : "0";
                myExcelWorksheet.get_Range("O" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["VistorvsLY"] && dtVisitors.Rows[i]["VistorvsLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["VistorvsLY"].ToString() : "0";
                
                myExcelWorksheet.get_Range("Q" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["Buyer"] && dtVisitors.Rows[i]["Buyer"].ToString().Length > 0) ? dtVisitors.Rows[i]["Buyer"].ToString() : "0";
                myExcelWorksheet.get_Range("R" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["BuyervsLW"] && dtVisitors.Rows[i]["BuyervsLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["BuyervsLW"].ToString() : "0";
                myExcelWorksheet.get_Range("S" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["BuyervsLY"] && dtVisitors.Rows[i]["BuyervsLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["BuyervsLY"].ToString() : "0";


                myExcelWorksheet.get_Range("U" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["Conversion"] && dtVisitors.Rows[i]["Conversion"].ToString().Length > 0) ? dtVisitors.Rows[i]["Conversion"].ToString() : "0";
                myExcelWorksheet.get_Range("V" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["ConversionvsLW"] && dtVisitors.Rows[i]["ConversionvsLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["ConversionvsLW"].ToString() : "0";
                myExcelWorksheet.get_Range("W" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["CoversionvsLY"] && dtVisitors.Rows[i]["CoversionvsLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["CoversionvsLY"].ToString() : "0";

                myExcelWorksheet.get_Range("X" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["IPC"] && dtVisitors.Rows[i]["IPC"].ToString().Length > 0) ? dtVisitors.Rows[i]["IPC"].ToString() : "0";
                myExcelWorksheet.get_Range("Y" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["IPCvsLW"] && dtVisitors.Rows[i]["IPCvsLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["IPCvsLW"].ToString() : "0";
                myExcelWorksheet.get_Range("Z" + (i + 6), misValue).Formula = (null != dtVisitors.Rows[i]["IPCvsLY"] && dtVisitors.Rows[i]["IPCvsLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["IPCvsLY"].ToString() : "0";
               
            }    
        }

        #endregion WriteToExcel

        #region InsertVisitorsReport
        /// <summary>
        /// Insert Visitors Report
        /// </summary>
        private void InsertVisitorsReport()
        {
          
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            bool Result = ObjStock.InsertVisitorsReport();

        }

        #endregion InsertVisitorsReport

        
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

            else
            {
                // Do nothing
            }
        }

        #endregion FileDownload


        #region GetWeekNumber
        /// <summary>
        /// Get Week Number
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        private int GetWeekNumber(DateTime date)
        {
            //Constants
            const int JAN = 1;
            const int DEC = 12;
            const int LASTDAYOFDEC =31 ;
            const int FIRSTDAYOFJAN = 1;
            const int THURSDAY = 4;
            bool thursdayFlag = false;

            //Get the day number since the beginning of the year
            int dayOfYear = date.DayOfYear;

            //Get the first and last weekday of the year
            int startWeekDayOfYear = (int)(new DateTime(date.Year, JAN, FIRSTDAYOFJAN)).DayOfWeek;
            int endWeekDayOfYear = (int)(new DateTime(date.Year, DEC, LASTDAYOFDEC)).DayOfWeek;

            //Compensate for using monday as the first day of the week
            if (startWeekDayOfYear == 0)
                startWeekDayOfYear = 7;
            if (endWeekDayOfYear == 0)
                endWeekDayOfYear = 7;

            //Calculate the number of days in the first week
            int daysInFirstWeek = 8 - (startWeekDayOfYear);

            //Year starting and ending on a thursday will have 53 weeks
            if (startWeekDayOfYear == THURSDAY || endWeekDayOfYear == THURSDAY)
                thursdayFlag = true;

            //We begin by calculating the number of FULL weeks between
            //the year start and our date. The number is rounded up so
            //the smallest possible value is 0.
            int fullWeeks = (int)Math.Ceiling((dayOfYear - (daysInFirstWeek)) / 7.0);
            int resultWeekNumber = fullWeeks;

            //If the first week of the year has at least four days, the
            //actual week number for our date can be incremented by one.
            if (daysInFirstWeek >= THURSDAY)
                resultWeekNumber = resultWeekNumber + 1;

            //If the week number is larger than 52 (and the year doesn't
            //start or end on a thursday), the correct week number is 1.
            if (resultWeekNumber > 52 && !thursdayFlag)
                resultWeekNumber = 1;

            //If the week number is still 0, it means that we are trying
            //to evaluate the week number for a week that belongs to the
            //previous year (since it has 3 days or less in this year).
            //We therefore execute this function recursively, using the
            //last day of the previous year.
            if (resultWeekNumber == 0)
                resultWeekNumber = GetWeekNumber(new DateTime(date.Year - 1, DEC, LASTDAYOFDEC));
            return resultWeekNumber;
        }

        #endregion GetWeekNumber


        #region GenerateVisitorsWeeklyReport
        /// <summary>
        ///  Generate Visitors Weekly Report
        /// </summary>
        private void GenerateVisitorsWeeklyReport()
        {
             //try
            //{
                DataTable dtVisitor = null;
                
                Excel1.Application myExcelApp;

                Excel1.Workbooks myExcelWorkbooks;

                Excel1.Workbook myExcelWorkbook;


                object misValue = System.Reflection.Missing.Value;

                myExcelApp = new Excel1.Application();

                myExcelApp.Visible = false;

                myExcelWorkbooks = myExcelApp.Workbooks;

                string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\WeeklyReport3.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            
            
                //myExcelWorkbooks = myExcelApp.Workbooks;

                //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
               
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                GetStockDetails objVisitors = new GetStockDetails();

                DateTime PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                string weekDay = PostingDate.DayOfWeek.ToString();
               // DateTime LastSat = DateTime.Now;
                DateTime NextSat = DateTime.Now;


                switch (weekDay)
                {
                    case "Sunday":NextSat = PostingDate.AddDays(6);
                        break;
                    case "Monday":NextSat = PostingDate.AddDays(5);
                        break;
                    case "Tuesday":NextSat = PostingDate.AddDays(4);
                        break;
                    case "Wednesday": NextSat = PostingDate.AddDays(3);  
                        break;
                    case "Thursday": NextSat = PostingDate.AddDays(2);    
                        break;
                    case "Friday":NextSat = PostingDate.AddDays(1);
                        break;
                    case "Saturday": NextSat = PostingDate.AddDays(0);
                        break;

                    default: NextSat = PostingDate;
                        break;

                }
            
                DateTime Week0 =NextSat;
                DateTime Week1 =NextSat.AddDays(-7); 
                DateTime Week2 =NextSat.AddDays(-14);
                DateTime Week3 =NextSat.AddDays(-21);

               //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
                
                if (txtLocation.Text.Trim().Length > 0)
                {
                    string location = txtLocation.Text.Trim();

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport(location, Week0);

                   if (dtVisitor.Rows.Count > 0)
                    {
                        xlSheet.Name = location;
                        WriteToExcelWeekly(dtVisitor, xlSheet, location, 6,Week0);

                        dtVisitor = objVisitors.GetVisitorsWeeklyReport(location, Week1);
                        WriteToExcelWeekly(dtVisitor, xlSheet, location,17,Week1);

                        dtVisitor = objVisitors.GetVisitorsWeeklyReport(location, Week2);
                        WriteToExcelWeekly(dtVisitor, xlSheet, location, 28,Week2);

                        dtVisitor = objVisitors.GetVisitorsWeeklyReport(location, Week3);
                        WriteToExcelWeekly(dtVisitor, xlSheet, location, 39,Week3);
                        
                        lblMessage.Visible = true;
                        lblMessage.ForeColor = System.Drawing.Color.Green;
                        lblMessage.Text = "Report Generation Complete";

                        btnDownloadWeekly.Visible = true;
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
                    
                    
                    //Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;
                    Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheet0409.Name = "0409";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0409", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0409, "0409", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0409", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0409, "0409", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0409", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0409, "0409", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0409", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0409, "0409", 39, Week3);


                    Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheet0410.Name = "0410";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0410", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0410, "0410", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0410", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0410, "0410", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0410", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0410, "0410", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0410", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0410, "0410", 39, Week3);

                    Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheet0414.Name = "0414";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0414", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0414, "0414", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0414", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0414, "0414", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0414", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0414, "0414", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0414", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0414, "0414", 39, Week3);


                    Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheet0415.Name = "0415";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0415", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0415, "0415", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0415", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0415, "0415", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0415", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0415, "0415", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0415", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0415, "0415", 39, Week3);

                    Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheet0416.Name = "0416";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0416", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0416, "0416", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0416", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0416, "0416", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0416", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0416, "0416", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0416", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0416, "0416", 39, Week3);

                    Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheet0417.Name = "0417";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0417", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0417, "0417", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0417", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0417, "0417", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0417", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0417, "0417", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0417", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0417, "0417", 39, Week3);


                    Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheet0419.Name = "0419";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0419", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0419, "0419", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0419", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0419, "0419", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0419", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0419, "0419", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0419", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0419, "0419", 39, Week3);


                    Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                    xlSheet0418.Name = "0418";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0418", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0418, "0418", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0418", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0418, "0418", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0418", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0418, "0418", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0418", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0418, "0418", 39, Week3);

                    
                    
                    Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                    xlSheet0421.Name = "0421";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0421", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0421, "0421", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0421", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0421, "0421", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0421", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0421, "0421", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0421", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0421, "0421", 39, Week3);



                    Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                    xlSheet0412.Name = "0412";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0412", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0412, "0412", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0412", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0412, "0412", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0412", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0412, "0412", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0412", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0412, "0412", 39, Week3);

                    Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                    xlSheet0422.Name = "0422";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0422", Week0);
                    WriteToExcelWeekly(dtVisitor, xlSheet0422, "0422", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0422", Week1);
                    WriteToExcelWeekly(dtVisitor, xlSheet0422, "0422", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0422", Week2);
                    WriteToExcelWeekly(dtVisitor, xlSheet0422, "0422", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0422", Week3);
                    WriteToExcelWeekly(dtVisitor, xlSheet0422, "0422", 39, Week3);


                    Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                    xlSheet0423.Name = "0423";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0423", Week0);
                    if(dtVisitor.Rows.Count>0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0423, "0423", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0423", Week1);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0423, "0423", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0423", Week2);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0423, "0423", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0423", Week3);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0423, "0423", 39, Week3);

                    //0424

                    Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                    xlSheet0424.Name = "0424";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0424", Week0);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0424, "0424", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0424", Week1);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0424, "0424", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0424", Week2);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0424, "0424", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0424", Week3);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0424, "0424", 39, Week3);



                    //0425

                    Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                    xlSheet0425.Name = "0425";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0425", Week0);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0425, "0425", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0425", Week1);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0425, "0425", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0425", Week2);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0425, "0425", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0425", Week3);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0425, "0425", 39, Week3);



                    //0426

                    Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                    xlSheet0426.Name = "0426";
                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0426", Week0);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0426, "0426", 6, Week0);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0426", Week1);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0426, "0426", 17, Week1);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0426", Week2);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0426, "0426", 28, Week2);

                    dtVisitor = objVisitors.GetVisitorsWeeklyReport("0426", Week3);
                    if (dtVisitor.Rows.Count > 0)
                        WriteToExcelWeekly(dtVisitor, xlSheet0426, "0426", 39, Week3);



                //0427

                Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                xlSheet0427.Name = "0427";
                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0427", Week0);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0427, "0427", 6, Week0);

                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0427", Week1);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0427, "0427", 17, Week1);

                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0427", Week2);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0427, "0427", 28, Week2);

                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0427", Week3);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0427, "0427", 39, Week3);

                //0428

                Excel1.Worksheet xlSheet0428 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                xlSheet0428.Name = "0428";
                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0428", Week0);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0428, "0428", 6, Week0);

                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0428", Week1);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0428, "0428", 17, Week1);

                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0428", Week2);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0428, "0428", 28, Week2);

                dtVisitor = objVisitors.GetVisitorsWeeklyReport("0428", Week3);
                if (dtVisitor.Rows.Count > 0)
                    WriteToExcelWeekly(dtVisitor, xlSheet0428, "0428", 39, Week3);



                lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";
                    btnDownloadWeekly.Visible = true;

                }

                Random rnd = new Random();
                string filePath = Server.MapPath(".") +"\\Reports\\Visitors_Weekly_"+DateTime.Now.Day+"-"+DateTime.Now.Month+"-"+DateTime.Now.Year+"_"+ rnd.Next()+".xlsx";

                ViewState["FileNameWeekly"] = filePath;
                myExcelWorkbook.SaveAs(@filePath);
                
                myExcelWorkbook.Close();
                myExcelWorkbooks.Close();
            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateVisitorsWeeklyReport

               
        #region WriteToExcelWeekly
        /// <summary>
        /// Write To Excel Weekly
        /// </summary>
        /// <param name="dtVisitors"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelWeekly(DataTable dtVisitors, Excel1.Worksheet myExcelWorksheet, string location, int j, DateTime WeekDate)
        {
            object misValue = System.Reflection.Missing.Value;
           
            String FromDate = WeekDate.AddDays(-6).ToString("dd/MM/yyyy");
            String ToDate = WeekDate.ToString("dd/MM/yyyy"); 

            myExcelWorksheet.get_Range("M" + (j - 2), misValue).Formula = "Week:__ - Date - " + FromDate + " To " + ToDate;

            for (int i = 0; i < dtVisitors.Rows.Count; i++, j++)
            {

                int Visitors = (null != dtVisitors.Rows[i]["Visitors"] && dtVisitors.Rows[i]["Visitors"].ToString().Length > 0) ? Convert.ToInt32(dtVisitors.Rows[i]["Visitors"]) : 0;
                int Buyers = (null != dtVisitors.Rows[i]["Buyers"] && dtVisitors.Rows[i]["Buyers"].ToString().Length > 0) ? Convert.ToInt32(dtVisitors.Rows[i]["Buyers"]) : 0;

                if (Buyers > Visitors)
                {
                    myExcelWorksheet.get_Range("J" + j, misValue).EntireRow.Interior.Color = System.Drawing.Color.Red;
                }

                myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtVisitors.Rows[i]["PostingDate"] && dtVisitors.Rows[i]["PostingDate"].ToString().Length > 0) ? DateTime.ParseExact(dtVisitors.Rows[i]["PostingDate"].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("ddd, dd/MM/yyyy") + "" : "0";
                
                myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtVisitors.Rows[i]["Visitors"] && dtVisitors.Rows[i]["Visitors"].ToString().Length > 0) ? dtVisitors.Rows[i]["Visitors"].ToString() : "0";
                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtVisitors.Rows[i]["VisitorsLW"] && dtVisitors.Rows[i]["VisitorsLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["VisitorsLW"].ToString() : "0";
                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtVisitors.Rows[i]["VisitorsLY"] && dtVisitors.Rows[i]["VisitorsLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["VisitorsLY"].ToString() : "0";
                
                myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtVisitors.Rows[i]["Buyers"] && dtVisitors.Rows[i]["Buyers"].ToString().Length > 0) ? dtVisitors.Rows[i]["Buyers"].ToString() : "0";
                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtVisitors.Rows[i]["BuyersLW"] && dtVisitors.Rows[i]["BuyersLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["BuyersLW"].ToString() : "0";
                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtVisitors.Rows[i]["BuyersLY"] && dtVisitors.Rows[i]["BuyersLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["BuyersLY"].ToString() : "0";
                
                //myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtVisitors.Rows[i]["Conversion"] && dtVisitors.Rows[i]["Conversion"].ToString().Length > 0) ? dtVisitors.Rows[i]["Conversion"].ToString() : "0";
                //myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtVisitors.Rows[i]["ConversionLW"] && dtVisitors.Rows[i]["ConversionLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["ConversionLW"].ToString() : "0";

                myExcelWorksheet.get_Range("AA" + j, misValue).Formula = (null != dtVisitors.Rows[i]["Ipc"] && dtVisitors.Rows[i]["Ipc"].ToString().Length > 0) ? dtVisitors.Rows[i]["Ipc"].ToString() : "0";
                myExcelWorksheet.get_Range("H" + j, misValue).Formula = (null != dtVisitors.Rows[i]["IpcLW"] && dtVisitors.Rows[i]["IpcLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["IpcLW"].ToString() : "0";
                myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtVisitors.Rows[i]["IpcLY"] && dtVisitors.Rows[i]["IpcLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["IpcLY"].ToString() : "0";
            }
        }

        #endregion WriteToExcelWeekly

        #region GetProcessStatus
        /// <summary>
        /// Get Process Status
        /// </summary>
        private bool GetProcessStatus()
        {
            GetStockDetails objStock = new GetStockDetails();
            objStock.ProcessStatusId = VisitorReportProcessId;
            DataTable dtStatus = objStock.GetProcessStatus();
            bool Flag = Convert.ToBoolean(dtStatus.Rows[0]["Flag"]);

            return Flag;
        }
        #endregion GetProcessStatus


        #region GenerateVisitorsVsSales
        /// <summary>
        /// Generate Visitors Vs Sales
        /// </summary>
        private void GenerateVisitorsVsSales()
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


            String fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\VisitorsVsSalesReport.xlsx";



            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.ActiveSheet;

            //myExcelWorkbooks.Close();

            // myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);

            

            GetStockDetails objStock = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();

                Excel1.Worksheet xlSheetUAE = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetUAE.Name = "UAE";
                objStock.PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.Location = "UAE";
                dtVisitors = objStock.GetVisitorsVsSales();
                WriteToExcelVisitorsVsSales(dtVisitors, xlSheetUAE);

                Excel1.Worksheet xlSheetBAH = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetBAH.Name = "BAHRAIN";
                objStock.PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.Location = "BAHRAIN";
                dtVisitors = objStock.GetVisitorsVsSales();
                WriteToExcelVisitorsVsSales(dtVisitors, xlSheetBAH);

                Excel1.Worksheet xlSheetJOR = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetJOR.Name = "JORDAN";
                objStock.PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.Location = "JORDAN";
                dtVisitors = objStock.GetVisitorsVsSales();
                WriteToExcelVisitorsVsSales(dtVisitors, xlSheetJOR);

                Excel1.Worksheet xlSheetOMAN = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheetOMAN.Name = "OMAN";
                objStock.PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.Location = "OMAN";
                dtVisitors = objStock.GetVisitorsVsSales();
                WriteToExcelVisitorsVsSales(dtVisitors, xlSheetOMAN);

                Excel1.Worksheet xlSheetQatar = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheetQatar.Name = "QATAR";
                objStock.PostingDate = DateTime.ParseExact(txtDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objStock.Location = "QATAR";
                dtVisitors = objStock.GetVisitorsVsSales();
                WriteToExcelVisitorsVsSales(dtVisitors, xlSheetQatar);
            
                
            
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnVisitorsVsSales.Visible = true;

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\VisitorsVsSales_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameVisitorsVsSales"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();
            //}

            //catch (Exception e)
            //{

            //}

        }

        #endregion GenerateVisitorsVsSales

        #region WriteToExcelVisitorsVsSales
        /// <summary>
        /// Write To Excel Visitors Vs Sales
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelVisitorsVsSales(DataTable dtVisitors, Excel1.Worksheet myExcelWorksheet)
        {
            object misValue = System.Reflection.Missing.Value;

            myExcelWorksheet.get_Range("A1", misValue).Formula = myExcelWorksheet.get_Range("A1", misValue).Formula + " - " + txtDate.Text.Trim();

            for (int i = 0; i < dtVisitors.Rows.Count; i++)
            {
                myExcelWorksheet.get_Range("A" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["StoreNo"] && dtVisitors.Rows[i]["StoreNo"].ToString().Length > 0) ? dtVisitors.Rows[i]["StoreNo"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["Visitors"] && dtVisitors.Rows[i]["Visitors"].ToString().Length > 0) ? dtVisitors.Rows[i]["Visitors"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["Buyers"] && dtVisitors.Rows[i]["Buyers"].ToString().Length > 0) ? dtVisitors.Rows[i]["Buyers"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                
                myExcelWorksheet.get_Range("D" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["Conversion"] && dtVisitors.Rows[i]["Conversion"].ToString().Length > 0) ? dtVisitors.Rows[i]["Conversion"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["Sales"] && dtVisitors.Rows[i]["Sales"].ToString().Length > 0) ? dtVisitors.Rows[i]["Sales"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                
                myExcelWorksheet.get_Range("F" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["VisitorsLW"] && dtVisitors.Rows[i]["VisitorsLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["VisitorsLW"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("G" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["BuyersLW"] && dtVisitors.Rows[i]["BuyersLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["BuyersLW"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("G" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("H" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["ConversionLW"] && dtVisitors.Rows[i]["ConversionLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["ConversionLW"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("H" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("I" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["SalesLW"] && dtVisitors.Rows[i]["SalesLW"].ToString().Length > 0) ? dtVisitors.Rows[i]["SalesLW"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("I" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("J" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["VisitorsLY"] && dtVisitors.Rows[i]["VisitorsLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["VisitorsLY"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("J" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("K" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["BuyersLY"] && dtVisitors.Rows[i]["BuyersLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["BuyersLY"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("K" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("L" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["ConversionLY"] && dtVisitors.Rows[i]["ConversionLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["ConversionLY"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("L" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("M" + (i + 4), misValue).Formula = (null != dtVisitors.Rows[i]["SalesLY"] && dtVisitors.Rows[i]["SalesLY"].ToString().Length > 0) ? dtVisitors.Rows[i]["SalesLY"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("M" + (i + 4), misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            }
        }

        #endregion WriteToExcelVisitorsVsSales

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