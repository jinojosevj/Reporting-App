
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
    public partial class CashFlowTATI : System.Web.UI.Page
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
            TatiBAL objTati = new TatiBAL();

            objTati.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
            objTati.Year = ddlYear.SelectedItem.Text;
            objTati.MonthNo = Convert.ToInt32(ddlMonth.SelectedItem.Value);
            bool Result = false;

            if (chkCashFlow.Checked)
                Result =objTati.InsertCashFlow();
            if (chkProfitLoss.Checked)
                Result = objTati.InsertProfitAndLossReport();
            if(chkCashFlowMY.Checked)
                Result = objTati.InsertCashFlowMY();
            if (chkProfitLossMY.Checked)
                Result = objTati.InsertProfitAndLossReportMY();

            if (Result)
            {
                if(chkCashFlow.Checked)
                    GenerateReport();
                
                if(chkProfitLoss.Checked)
                    GeneratePLReport();

                if (chkCashFlowMY.Checked)
                    GenerateReportMY();

                if (chkProfitLossMY.Checked)
                    GeneratePLReportMY();
            }
            else
            {
                lblMessage.Text = "Report Failed !";
                lblMessage.ForeColor = System.Drawing.Color.Red;
            }
        }
        #endregion btnGenerate_Click

        #region btnDownloadCashFlow_Click
        /// <summary>
        /// btnDownloadCashFlow_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadCashFlow_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileName"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadCashFlow_Click


        #region btnDownloadCashFlowMY_Click
        /// <summary>
        /// btnDownloadCashFlowMY_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownloadCashFlowMY_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameMY"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownloadCashFlowMY_Click


        #region btnProfitLoss_Click
        /// <summary>
        /// btnProfitLoss_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnProfitLoss_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNamePL"].ToString();
            FileDownload(filename);
        }
        #endregion btnProfitLoss_Click


        #region btnProfitLossMY_Click
        /// <summary>
        /// btnProfitLossMY_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnProfitLossMY_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNamePLMY"].ToString();
            FileDownload(filename);

        }
        #endregion btnProfitLossMY_Click

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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\CashFlowReport.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            TatiBAL objStock = new TatiBAL();

            //objStock.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\CashFlowReportOne.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;

                dtStock = objStock.GetCashFlowTati();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                    WriteToExcelHeader(dtStock, xlSheet, location);
                    WriteToExcel(dtStock, xlSheet, location);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnDownloadCashFlow.Visible = true;
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


                Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetSummary.Name = "Summary";
                objStock.Location = "Summary";
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text);
                dtStock = objStock.GetCashFlowTati();
                WriteToExcelHeader(dtStock, xlSheetSummary, "Summary");
                WriteToExcel(dtStock, xlSheetSummary, "Summary");


                Excel1.Worksheet xlSheetJordan = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetJordan.Name = "Jordan";
                objStock.Location = "Jordan";
                dtStock = objStock.GetCashFlowTati();
                WriteToExcelHeader(dtStock, xlSheetJordan, "Jordan");
                WriteToExcel(dtStock, xlSheetJordan, "Jordan");



                Excel1.Worksheet xlSheetQHO = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetQHO.Name = "TQHO";
                objStock.Location = "TQHO";
                dtStock = objStock.GetCashFlowTati();
                WriteToExcelHeader(dtStock, xlSheetQHO, "TQHO");
                WriteToExcel(dtStock, xlSheetQHO, "TQHO");

                

                Excel1.Worksheet xlSheet4728 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet4728.Name = "4728";
                objStock.Location = "4728";
                dtStock = objStock.GetCashFlowTati();
                WriteToExcelHeader(dtStock, xlSheet4728, "4728");
                WriteToExcel(dtStock, xlSheet4728, "4728");


                Excel1.Worksheet xlSheet4729 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheet4729.Name = "4729";
                objStock.Location = "4729";
                dtStock = objStock.GetCashFlowTati();
                WriteToExcelHeader(dtStock, xlSheet4729, "4729");
                WriteToExcel(dtStock, xlSheet4729, "4729");

                Excel1.Worksheet xlSheet4731 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheet4731.Name = "4731";
                objStock.Location = "4731";
                dtStock = objStock.GetCashFlowTati();
                WriteToExcelHeader(dtStock, xlSheet4731, "4731");
                WriteToExcel(dtStock, xlSheet4731, "4731");


               
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownloadCashFlow.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\CashFlowTati_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

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


        #region GenerateReportMY
        /// <summary>
        /// To generate excel report for stock values
        /// </summary>
        private void GenerateReportMY()
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\CashFlowMY.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            TatiBAL objStock = new TatiBAL();

            //objStock.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\CashFlowMY.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;

                dtStock = objStock.GetCashFlowMY();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                    WriteToExcelHeader(dtStock, xlSheet, location);
                    WriteToExcel(dtStock, xlSheet, location);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";

                    btnDownloadCashFlow.Visible = true;
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


                Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetSummary.Name = "Summary";
                objStock.Location = "Summary";
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text);
                dtStock = objStock.GetCashFlowMY();
                WriteToExcelHeader(dtStock, xlSheetSummary, "Summary");
                WriteToExcel(dtStock, xlSheetSummary, "MY-Summary");


                Excel1.Worksheet xlSheetHO = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetHO.Name = "HO";
                objStock.Location = "HO";
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text);
                dtStock = objStock.GetCashFlowMY();
                WriteToExcelHeader(dtStock, xlSheetHO, "HO");
                WriteToExcel(dtStock, xlSheetHO, "HO");


                Excel1.Worksheet xlSheetF004 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetF004.Name = "F004";
                objStock.Location = "F004";
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text);
                dtStock = objStock.GetCashFlowMY();
                WriteToExcelHeader(dtStock, xlSheetF004, "F004");
                WriteToExcel(dtStock, xlSheetF004, "F004");

                Excel1.Worksheet xlSheetF007 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheetF007.Name = "F007";
                objStock.Location = "F007";
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text);
                dtStock = objStock.GetCashFlowMY();
                WriteToExcelHeader(dtStock, xlSheetF007, "F007");
                WriteToExcel(dtStock, xlSheetF007, "F007");


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnDownloadCashFlowMY.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\CashFlowMY_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameMY"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);


            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();



            //}

            //catch (Exception e)
            //{

            //}


        }

        #endregion GenerateReportMY


        #region WriteToExcel
        /// <summary>
        /// Write To Excel
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcel(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, String storeName)
        {
            object misValue = System.Reflection.Missing.Value;
            
            myExcelWorksheet.get_Range("C1", misValue).Formula = storeName+ " Cash Flow Report For Week No: "+ txtWeekNo.Text+" - "+ddlYear.SelectedItem.Value;
            //myExcelWorksheet.get_Range("A2", misValue).Formula = "Highest " + ddlType.SelectedItem.Text + " As Of " + txtAsOfDate.Text.ToString();
            //BorderAround(myExcelWorksheet.get_Range("A2", misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("C1", misValue).Font.Bold = true;

            

              TatiBAL objTati=new TatiBAL();
              objTati.Year = ddlYear.SelectedItem.Text;
              objTati.Location = storeName;
              objTati.JorRate = Convert.ToDecimal(txtJordanRate.Text);

              DataTable dtBank = null;

            if (storeName == "F007" || storeName == "F004" || storeName == "HO" || storeName == "MY-Summary")
            {
                dtBank = objTati.GetCashFlowBankOpeningMY();
            }
            else
            {
                dtBank = objTati.GetCashFlowBankOpening();
            }

            if (dtBank.Rows.Count > 0)
            {
                myExcelWorksheet.get_Range("C4", misValue).Formula = dtBank.Rows[0]["Amount"].ToString();
            }
            
              DataTable dtMonth = null;
              int WeekStart = 0;
              String WeekStartColumn = "A";
              String WeekEndColumn = "D";
              int WeekEnd = 0;
              char ColumnName = 'C';
              char ColumnName1 = 'A';
              int ColumnIndex = 1;
              string StrColumnName = "C";

              int SumColumnStart = 0;
              int SumColumnEnd = 0;
              
              String LastColumnName="C";
             

              for (int k = 1; k <= 12; k++)
              {

                  switch (k)
                  {
                      case 1: objTati.Month = "Jan";
                              break;
                      case 2: objTati.Month = "Feb";
                              break;
                      case 3: objTati.Month = "Mar";
                              break;
                      case 4: objTati.Month = "Apr";
                              break;

                      case 5: objTati.Month = "May";
                              break;
                      case 6: objTati.Month = "Jun";
                              break;
                      case 7: objTati.Month = "Jul";
                              break;
                      case 8: objTati.Month = "Aug";
                              break;

                      case 9: objTati.Month = "Sep";
                              break;
                      case 10: objTati.Month = "Oct";
                              break;
                      case 11: objTati.Month = "Nov";
                              break;
                      case 12: objTati.Month = "Dec";
                              break;
                  
                  }
                  
                  
                  dtMonth = objTati.GetMonth();

                  if (dtMonth.Rows.Count > 0)
                  {
                      WeekStart = Convert.ToInt32(dtMonth.Rows[0]["WeekStart"]);
                      WeekStartColumn = StrColumnName;
                      WeekEnd = Convert.ToInt32(dtMonth.Rows[0]["WeekEnd"]);
                  }
                  
                  while (WeekStart <= WeekEnd)
                  {
                      if(StrColumnName!="C")
                      {
                          myExcelWorksheet.get_Range(StrColumnName.ToString() + 4, misValue).Formula = "="+LastColumnName+"128";
                      }
                      
                      int j = 3;
                      for (int i = 0; i < dtStock.Rows.Count; i++, j++)
                      {

                          if (j == 3)
                          {
                              myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Formula = "Week" + WeekStart;
                              i =-1;
                              j++;
                          }
                          else
                          {
                              int ValueType = (DBNull.Value != dtStock.Rows[i]["CashFlowType"]) ? Convert.ToInt32(dtStock.Rows[i]["CashFlowType"].ToString()) : 0;
                             
                              if (ValueType == 1)
                              {
                                  SumColumnStart = j;
                              
                              }


                              if (ValueType == 2 )
                              {
                                  SumColumnEnd = j - 1;

                                  myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Formula = "=SUM(" + StrColumnName.ToString() + SumColumnStart + ":" + StrColumnName.ToString() + SumColumnEnd + ")";
                                  BorderAround(myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                                 
                                 // myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Interior.Color = System.Drawing.Color.Yellow;
                                  j++;
                              }
                              else if (ValueType == 3)
                              {
                                  j++;
                              
                              }

                              else
                              {
                                  myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Formula = (null != dtStock.Rows[i]["Week" + WeekStart]) ? dtStock.Rows[i]["Week" + WeekStart].ToString() : "0";
                                  BorderAround(myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                              }
                          }
                          //myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["ClsQty"]) ? dtStock.Rows[i]["ClsQty"].ToString() : "0";
                          //BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                      }


                      LastColumnName = StrColumnName;


                      if(ColumnName=='Z')
                      {
                          ColumnName = 'A';

                          if(ColumnName!='A')
                            ColumnName++;

                          if (ColumnIndex != 1)
                            ColumnName1++;

                          ColumnIndex++;

                          StrColumnName = ColumnName1.ToString() + ColumnName.ToString();


                      }
                      else
                      {
                          ColumnName++;
                         if (ColumnIndex == 1)
                        {
                            StrColumnName = ColumnName.ToString();
                        }
                        else
                        {
                            StrColumnName = ColumnName1.ToString() + ColumnName.ToString();
                        }


                      }


                      if ((WeekEnd - WeekStart) == 1)
                      {
                          WeekEndColumn = StrColumnName;
                      }


                      if (WeekStart == WeekEnd)
                      {

                          myExcelWorksheet.get_Range(StrColumnName.ToString() + 3, misValue).Formula = objTati.Month + " Total";
                          myExcelWorksheet.get_Range(StrColumnName.ToString() + 3, misValue).Interior.Color = System.Drawing.Color.LightBlue;

                          //myExcelWorksheet.get_Range(StrColumnName.ToString() + 4, misValue).Formula = "";
                          //myExcelWorksheet.get_Range(StrColumnName.ToString() + 4, misValue).Interior.Color = System.Drawing.Color.LightBlue;

                          myExcelWorksheet.get_Range(StrColumnName.ToString() + 128, misValue).Formula = "";

                          j = 5;
                          for (int i = 0; i < dtStock.Rows.Count; i++, j++)
                          {

                              int ValueType1 = (DBNull.Value != dtStock.Rows[i]["CashFlowType"]) ? Convert.ToInt32(dtStock.Rows[i]["CashFlowType"].ToString()) : 0;
                              if (ValueType1 == 2 || ValueType1 == 3)
                              {
                                  myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Formula = "=SUM(" + WeekStartColumn + j + ":" + WeekEndColumn.ToString() + j + ")";
                                  BorderAround(myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                                  myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Interior.Color = System.Drawing.Color.LightBlue;

                                  j++;
                              }
                              else
                              {
                                  myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Formula = "=SUM(" + WeekStartColumn + j + ":" + WeekEndColumn.ToString() + j + ")";
                                  BorderAround(myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                                  myExcelWorksheet.get_Range(StrColumnName.ToString() + j, misValue).Interior.Color = System.Drawing.Color.LightBlue;
                              }
                          }

                          ColumnName++;
                          if (ColumnIndex == 1)
                          {
                              StrColumnName = ColumnName.ToString();
                          }
                          else
                          {
                              StrColumnName = ColumnName1.ToString() + ColumnName.ToString();
                          }

                      }

                    

                      WeekStart++;
                  }

              }

        }

        #endregion WriteToExcel


        #region WriteToExcelHeader
        /// <summary>
        /// WriteToExcelHeader
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelHeader(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, String storeName)
        {
            object misValue = System.Reflection.Missing.Value;

            int j = 5;

                    for (int i = 0; i < dtStock.Rows.Count; i++, j++)
                    {

                        int ValueType = (DBNull.Value != dtStock.Rows[i]["CashFlowType"]) ? Convert.ToInt32(dtStock.Rows[i]["CashFlowType"].ToString()) : 0;

                        if (ValueType == 2 || ValueType == 3)
                        {

                            myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["CashFlowCode"]) ? dtStock.Rows[i]["CashFlowCode"].ToString() : "0";
                            BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                            //myExcelWorksheet.get_Range("A" + j, misValue).Interior.Color = System.Drawing.Color.Yellow;
                            myExcelWorksheet.get_Range("A" + j, misValue).Font.Bold = true;

                            myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["CashFlowDescription"]) ? dtStock.Rows[i]["CashFlowDescription"].ToString() : "0";
                            BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                           // myExcelWorksheet.get_Range("B" + j, misValue).Interior.Color = System.Drawing.Color.Yellow;
                            myExcelWorksheet.get_Range("B" + j, misValue).Font.Bold = true;
                            j++;
                        }
                        else
                        {

                            myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["CashFlowCode"]) ? dtStock.Rows[i]["CashFlowCode"].ToString() : "0";
                            BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                            myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["CashFlowDescription"]) ? dtStock.Rows[i]["CashFlowDescription"].ToString() : "0";
                            BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                        }
                    }

        }

        #endregion WriteToExcelHeader

                
        #region GeneratePLReport
        /// <summary>
        /// To generate excel report for PL
        /// </summary>
        private void GeneratePLReport()
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReport.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            TatiBAL objStock = new TatiBAL();

            //objStock.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportOne.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                dtStock = objStock.GetProfitAndLossTati();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                   // WriteToExcelHeader(dtStock, xlSheet, location);
                    WriteToExcelPL(dtStock, xlSheet, location);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";
                    btnProfitLoss.Visible = true;
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
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetSummary.Name = "Summary";
                objStock.Location = "Summary";
                dtStock = objStock.GetProfitAndLossTati();
                WriteToExcelPL(dtStock, xlSheetSummary, "Summary");

                Excel1.Worksheet xlSheetJorSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetJorSummary.Name = "Jordan Summary";
                objStock.Location = "Jordan";
                dtStock = objStock.GetProfitAndLossTati();
                WriteToExcelPL(dtStock, xlSheetJorSummary, "Jordan");

                Excel1.Worksheet xlSheetQho = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetQho.Name = "TQHO";
                objStock.Location = "TQHO";
                dtStock = objStock.GetProfitAndLossTati();
                WriteToExcelPL(dtStock, xlSheetQho, "TQHO");



                Excel1.Worksheet xlSheet4728 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheet4728.Name = "4728";
                objStock.Location = "4728";
                dtStock = objStock.GetProfitAndLossTati();
                WriteToExcelPL(dtStock, xlSheet4728, "4728");


                Excel1.Worksheet xlSheet4729 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                xlSheet4729.Name = "4729";
                objStock.Location = "4729";
                dtStock = objStock.GetProfitAndLossTati();
                WriteToExcelPL(dtStock, xlSheet4729, "4729");


                Excel1.Worksheet xlSheet4731 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                xlSheet4731.Name = "4731";
                objStock.Location = "4731";
                dtStock = objStock.GetProfitAndLossTati();
                WriteToExcelPL(dtStock, xlSheet4731, "4731");


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnProfitLoss.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\ProfitAndLossTati_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNamePL"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);


            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();



            //}

            //catch (Exception e)
            //{

            //}

        }

        #endregion GeneratePLReport

        #region GeneratePLReportMY
        /// <summary>
        /// To generate excel report for PL
        /// </summary>
        private void GeneratePLReportMY()
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportMY.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            TatiBAL objStock = new TatiBAL();

            //objStock.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportMY.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                dtStock = objStock.GetProfitAndLossMY();

                if (dtStock.Rows.Count > 0)
                {
                    xlSheet.Name = location;
                    // WriteToExcelHeader(dtStock, xlSheet, location);
                    WriteToExcelPL(dtStock, xlSheet, location);

                    lblMessage.Visible = true;
                    lblMessage.ForeColor = System.Drawing.Color.Green;
                    lblMessage.Text = "Report Generation Complete";
                    btnProfitLossMY.Visible = true;
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
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                xlSheetSummary.Name = "Summary";
                objStock.Location = "Summary";
                dtStock = objStock.GetProfitAndLossMY();
                WriteToExcelPL(dtStock, xlSheetSummary, "Summary");


                Excel1.Worksheet xlSheetHO = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                xlSheetHO.Name = "HO";
                objStock.Location = "HO";
                dtStock = objStock.GetProfitAndLossMY();
                WriteToExcelPL(dtStock, xlSheetHO, "HO");


                Excel1.Worksheet xlSheetF004 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                xlSheetF004.Name = "F004";
                objStock.Location = "F004";
                dtStock = objStock.GetProfitAndLossMY();
                WriteToExcelPL(dtStock, xlSheetF004, "F004");


                Excel1.Worksheet xlSheetF007 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                xlSheetF007.Name = "F007";
                objStock.Location = "F007";
                dtStock = objStock.GetProfitAndLossMY();
                WriteToExcelPL(dtStock, xlSheetF007, "F007");


                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnProfitLossMY.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\ProfitAndLossMY_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNamePLMY"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);


            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();



            //}

            //catch (Exception e)
            //{

            //}

        }

        #endregion GeneratePLReportMY


        #region WriteToExcelPL
        /// <summary>
        /// WriteToExcelPL
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private void WriteToExcelPL(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, string location)
        {
            object misValue = System.Reflection.Missing.Value;
            string Rate = "";

            if (location == "Summary" || location == "QHO" || location == "F004")
            {
                Rate = "QAR";

            }
            else
            {
                Rate = "JOD";
            }

           // myExcelWorksheet.get_Range("A1", misValue).Formula = location;
            myExcelWorksheet.get_Range("C1", misValue).Formula = location+" - Profit And Loss Report For  " + ddlMonth.SelectedItem.Text.ToString() + " - " + ddlYear.SelectedItem.Value.ToString()+" ("+Rate+")";
            //BorderAround(myExcelWorksheet.get_Range("A2", misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
            myExcelWorksheet.get_Range("A1", misValue).Font.Bold = true;

            int j = 3;
             int Value = 0;
            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                 Value = (null != dtStock.Rows[i]["ValueType"]) ? Convert.ToInt32(dtStock.Rows[i]["ValueType"].ToString()) : 0;

                if (Value == 1)
                {
                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Code"]) ? dtStock.Rows[i]["Code"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                    j++;
                    //i++;
                }
                else if (Value == 2)
                {
                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Code"]) ? dtStock.Rows[i]["Code"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                    
                }
                else if (Value == 4)
                {
                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Code"]) ? dtStock.Rows[i]["Code"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["Month1"]) ? dtStock.Rows[i]["Month1"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["Month2"]) ? dtStock.Rows[i]["Month2"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["Month3"]) ? dtStock.Rows[i]["Month3"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Month4"]) ? dtStock.Rows[i]["Month4"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["Month5"]) ? dtStock.Rows[i]["Month5"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["Month6"]) ? dtStock.Rows[i]["Month6"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["Month7"]) ? dtStock.Rows[i]["Month7"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["Month8"]) ? dtStock.Rows[i]["Month8"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["Month9"]) ? dtStock.Rows[i]["Month9"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["Month10"]) ? dtStock.Rows[i]["Month10"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("U" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["Month11"]) ? dtStock.Rows[i]["Month11"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["Month12"]) ? dtStock.Rows[i]["Month12"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Y" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    j++;
                }
                else
                {

                    myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Code"]) ? dtStock.Rows[i]["Code"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["Month1"]) ? dtStock.Rows[i]["Month1"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["Month2"]) ? dtStock.Rows[i]["Month2"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["Month3"]) ? dtStock.Rows[i]["Month3"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("I" + j, misValue).Formula = (null != dtStock.Rows[i]["Month4"]) ? dtStock.Rows[i]["Month4"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("I" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("K" + j, misValue).Formula = (null != dtStock.Rows[i]["Month5"]) ? dtStock.Rows[i]["Month5"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("K" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("M" + j, misValue).Formula = (null != dtStock.Rows[i]["Month6"]) ? dtStock.Rows[i]["Month6"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("M" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("O" + j, misValue).Formula = (null != dtStock.Rows[i]["Month7"]) ? dtStock.Rows[i]["Month7"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("O" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Q" + j, misValue).Formula = (null != dtStock.Rows[i]["Month8"]) ? dtStock.Rows[i]["Month8"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Q" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("S" + j, misValue).Formula = (null != dtStock.Rows[i]["Month9"]) ? dtStock.Rows[i]["Month9"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("S" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("U" + j, misValue).Formula = (null != dtStock.Rows[i]["Month10"]) ? dtStock.Rows[i]["Month10"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("U" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                    myExcelWorksheet.get_Range("W" + j, misValue).Formula = (null != dtStock.Rows[i]["Month11"]) ? dtStock.Rows[i]["Month11"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("W" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                    myExcelWorksheet.get_Range("Y" + j, misValue).Formula = (null != dtStock.Rows[i]["Month12"]) ? dtStock.Rows[i]["Month12"].ToString() : "0";
                    BorderAround(myExcelWorksheet.get_Range("Y" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                }
                
            }

         
        }

        #endregion WriteToExcelPL
       

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