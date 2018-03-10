
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
    public partial class CashFlowDynamic : System.Web.UI.Page
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
            if(!IsPostBack)
               BindBrand();
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

           
            if(ddlBrand.SelectedItem.Value == "Matalan")
            {
                if (RdlRefresh.SelectedItem.Value == "0")
                {
                    MatalanTableUpdation();
                    GetStockDetails objMatalan = new GetStockDetails();
                    GeneratePLReport(objMatalan);
                }
                else if(RdlRefresh.SelectedItem.Value == "1")
                {
                    MatalanTableUpdation();
                }
                else if (RdlRefresh.SelectedItem.Value == "2")
                {
                    GetStockDetails objMatalan = new GetStockDetails();
                    GeneratePLReport(objMatalan);
                }

            }

            else if (ddlBrand.SelectedItem.Value == "Summary")
            {
                if (RdlRefresh.SelectedItem.Value == "0")
                {
                    //TatiTableUpdation();
                    TatiBAL objTati = new TatiBAL();
                    
                    objTati.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                    objTati.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                    objTati.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                    objTati.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());

                    objTati.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());
                    objTati.InsertProfitLossReportConsolidated();
                    GeneratePLReportConsolidated(objTati);
                }
                else if (RdlRefresh.SelectedItem.Value == "1")
                {
                    TatiBAL objTati = new TatiBAL();

                    objTati.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                    objTati.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                    objTati.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                    objTati.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());

                    objTati.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());
                    objTati.InsertProfitLossReportConsolidated();
                }
                else if (RdlRefresh.SelectedItem.Value == "2")
                {
                    TatiBAL objTati = new TatiBAL();
                    GeneratePLReportConsolidated(objTati);
                }

            }
            else
            {
                    if (RdlRefresh.SelectedItem.Value == "0")
                    {
                        TatiTableUpdation();
                        TatiBAL objTati = new TatiBAL();
                        GeneratePLReport(objTati);
                    }
                    else if (RdlRefresh.SelectedItem.Value == "1")
                    {
                        TatiTableUpdation();
                    }
                    else if (RdlRefresh.SelectedItem.Value == "2")
                    {
                        TatiBAL objTati = new TatiBAL();
                        GeneratePLReport(objTati);
                    }

            }



        }
        #endregion btnGenerate_Click


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

        #endregion Events

        #region Methods


        #region BindBrand
        /// <summary>
        /// BindBrand
        /// </summary>
        private void BindBrand()
        {
            TatiBAL objTati = new TatiBAL();
            DataTable dt = objTati.GetBrandDetails();
            ddlBrand.DataSource = dt;
            ddlBrand.DataMember = "BrandName";
            ddlBrand.DataValueField = "BrandName";
            ddlBrand.DataBind();
        }
        #endregion BindBrand


        #region GeneratePLReport
        /// <summary>
        /// To generate excel report for PL
        /// </summary>
        private void GeneratePLReport(TatiBAL objTati)
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportDynamic.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            //TatiBAL objTati = new TatiBAL();

            //objTati.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

          
            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportDynamic.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objTati.Location = location;
                objTati.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                dtStock = objTati.GetProfitAndLossTati();

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

                // Excel.Application xlApp = Marshal.GetActiveObject("Excel.Application") as Excel.Application;

                //Excel1.Workbook xlWb = myExcelApp.ActiveWorkbook as Excel1.Workbook;


                objTati.Type = true;
                objTati.Brand = ddlBrand.SelectedItem.Text;
                DataTable dtStore = objTati.GetStoreDetails();

                Excel1.Worksheet xlSht = myExcelWorkbook.Sheets[1];

                for (int i = 0; i < dtStore.Rows.Count; i++)
                {
                    xlSht.Copy(Type.Missing, myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count]); // copy
                    myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count].Name = "NEW SHEET";        // rename

                    xlSht = myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count];
                    //Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;
                    objTati.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                    //Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    string Location = dtStore.Rows[i]["LocationCode"].ToString();

                    xlSht.Name = Location;
                    objTati.Location = Location;
                    dtStock = objTati.GetProfitAndLossTati();
                    WriteToExcelPL(dtStock, xlSht, Location);
                  
                }
                    xlSht = myExcelWorkbook.Sheets[1];
                    xlSht.Visible = 0;

                


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


        #region GeneratePLReport
        /// <summary>
        /// To generate excel report for PL
        /// </summary>
        private void GeneratePLReport(GetStockDetails objMatalan)
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportDynamic.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            //TatiBAL objMatalan = new TatiBAL();

            //objMatalan.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;


            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportDynamic.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objMatalan.Location = location;

                objMatalan.JorRate  = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objMatalan.BahRate  = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                objMatalan.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objMatalan.UaeRate  = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objMatalan.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());

                dtStock = objMatalan.GetProfitAndLoss();

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

                // Excel.Application xlApp = Marshal.GetActiveObject("Excel.Application") as Excel.Application;

                //Excel1.Workbook xlWb = myExcelApp.ActiveWorkbook as Excel1.Workbook;


                objMatalan.IntType = 1;
                DataTable dtStore = objMatalan.GetStoreDetails();

                Excel1.Worksheet xlSht = myExcelWorkbook.Sheets[1];

                for (int i = 0; i < dtStore.Rows.Count; i++)
                {
                    xlSht.Copy(Type.Missing, myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count]); // copy
                    myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count].Name = "NEW SHEET";        // rename

                    xlSht = myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count];
                    //Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;
                    objMatalan.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                    objMatalan.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                    objMatalan.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                    objMatalan.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                    objMatalan.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());

                    //Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    string Location = dtStore.Rows[i]["LocationCode"].ToString();

                    xlSht.Name = Location;
                    objMatalan.Location = Location;
                    dtStock = objMatalan.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSht, Location);

                }
                xlSht = myExcelWorkbook.Sheets[1];
                xlSht.Visible = 0;
                              

                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnProfitLoss.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\ProfitAndLossMatalan_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

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


        #region GeneratePLReportConsolidated
        /// <summary>
        /// To generate excel report for PL
        /// </summary>
        private void GeneratePLReportConsolidated(TatiBAL objTati)
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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportDynamic.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            //TatiBAL objTati = new TatiBAL();

            //objTati.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;


            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportDynamic.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objTati.Location = location;
                objTati.Brand = ddlBrand.SelectedItem.Text;
                dtStock = objTati.GetProfitLossReportConsolidated();

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

                // Excel.Application xlApp = Marshal.GetActiveObject("Excel.Application") as Excel.Application;

                //Excel1.Workbook xlWb = myExcelApp.ActiveWorkbook as Excel1.Workbook;


                objTati.Type = true;
                DataTable dtStore = objTati.GetBrandDetails();

                Excel1.Worksheet xlSht = myExcelWorkbook.Sheets[1];

                for (int i = 0; i < dtStore.Rows.Count; i++)
                {
                    xlSht.Copy(Type.Missing, myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count]); // copy
                    myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count].Name = "NEW SHEET";        // rename

                    xlSht = myExcelWorkbook.Sheets[myExcelWorkbook.Sheets.Count];
                    //Excel1.Sheets xlSheets = myExcelWorkbook.Sheets as Excel1.Sheets;
                    objTati.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                    //Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    string Brand = dtStore.Rows[i]["BrandName"].ToString();

                    xlSht.Name = Brand;
                    objTati.Brand = Brand;
                    dtStock = objTati.GetProfitLossReportConsolidated();
                    WriteToExcelPL(dtStock, xlSht, Brand);

                }
                xlSht = myExcelWorkbook.Sheets[1];
                xlSht.Visible = 0;
                
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Report Generation Complete";
                btnProfitLoss.Visible = true;

            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\ProfitAndLossBTCF_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNamePL"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);


            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();

            //}

            //catch (Exception e)
            //{

            //}

        }

        #endregion GeneratePLReportConsolidated



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

            if (location == "Summary" || location == "TQHO" || location == "BTCF-HO" || location == "Qatar")
            {
                Rate = "QAR";

            }
            else if(location=="Oman")
            {
                Rate = "OMR";
            }
            else if (location == "Jordan")
            {
                Rate = "JOD";
            }
            else if (location == "UAE")
            {
                Rate = "UAE";
            }
            else
            {
                Rate = "";
            }

            // myExcelWorksheet.get_Range("A1", misValue).Formula = location;
            myExcelWorksheet.get_Range("C1", misValue).Formula = location + " - Profit And Loss Report For  " + ddlMonth.SelectedItem.Text.ToString() + " - " + ddlYear.SelectedItem.Value.ToString() + " (" + Rate + ")";
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


        #region TatiTableUpdation
        /// <summary>
        /// Tati Table Updation
        /// </summary>
        private void TatiTableUpdation()
        {
            TatiBAL objTati = new TatiBAL();
            objTati.Type = false;
            objTati.Brand = ddlBrand.SelectedItem.Text;
            DataTable dtStore = objTati.GetStoreDetails();
            objTati.DeleteProfitLossReport();

            for (int i = 0; i < dtStore.Rows.Count; i++)
            {
                objTati.Location = dtStore.Rows[i]["LocationCode"].ToString();
                objTati.Brand= dtStore.Rows[i]["Brand"].ToString();
                objTati.InsertGLAccountDetails();
            }

            objTati.Year = ddlYear.SelectedItem.Text;
            objTati.FromMonth = 1;
            objTati.ToMonth = Convert.ToInt32(ddlMonth.SelectedItem.Value);
            DataTable dtMonth = objTati.GetMonthDetails();

            for (int i = 0; i < dtMonth.Rows.Count; i++)
            {
                for (int j = 0; j < dtStore.Rows.Count; j++)
                {
                    objTati.FromDate = Convert.ToDateTime(dtMonth.Rows[i]["MStart"].ToString());
                    objTati.ToDate = Convert.ToDateTime(dtMonth.Rows[i]["MEnd"].ToString());
                    objTati.Month = dtMonth.Rows[i]["Month"].ToString();
                    objTati.Location = dtStore.Rows[j]["LocationCode"].ToString();
                    objTati.UpdateProfitLossActualReport();
                }

            }

            //update budgets

            objTati.ImportProfitLoseBudgets();

            objTati.Year = ddlYear.SelectedItem.Text;
            objTati.FromMonth = Convert.ToInt32(ddlMonth.SelectedItem.Value) + 1;
            objTati.ToMonth = 12;
            dtMonth = objTati.GetMonthDetails();

            for (int i = 0; i < dtMonth.Rows.Count; i++)
            {
                for (int j = 0; j < dtStore.Rows.Count; j++)
                {
                    objTati.FromDate = Convert.ToDateTime(dtMonth.Rows[i]["MStart"].ToString());
                    objTati.ToDate = Convert.ToDateTime(dtMonth.Rows[i]["MEnd"].ToString());
                    objTati.Month = dtMonth.Rows[i]["Month"].ToString();
                    objTati.Location = dtStore.Rows[j]["LocationCode"].ToString();

                    objTati.Year = ddlYear.SelectedItem.Value;
                    objTati.UpdateProfitLossBudgetReport();
                }
            }

        }

        #endregion TatiTableUpdation


        #region MatalanTableUpdation
        /// <summary>
        /// MatalanTableUpdation
        /// </summary>
        /// <returns></returns>
        private void MatalanTableUpdation()
        {
            GetStockDetails objMatalan = new GetStockDetails();
            objMatalan.Type = false;
            DataTable dtStore = objMatalan.GetStoreDetails();
            objMatalan.DeleteProfitLossReport();

            for (int i = 0; i < dtStore.Rows.Count; i++)
            {
                objMatalan.Location = dtStore.Rows[i]["LocationCode"].ToString();
                objMatalan.Country = dtStore.Rows[i]["Country"].ToString();
                objMatalan.InsertGLAccountDetails();
            }

            objMatalan.Year = ddlYear.SelectedItem.Text;
            objMatalan.FromMonth = 1;
            objMatalan.ToMonth = Convert.ToInt32(ddlMonth.SelectedItem.Value);
            DataTable dtMonth = objMatalan.GetMonthDetails();

            for (int i = 0; i < dtMonth.Rows.Count; i++)
            {
                for (int j = 0; j < dtStore.Rows.Count; j++)
                {
                    objMatalan.FromDate = Convert.ToDateTime(dtMonth.Rows[i]["MStart"].ToString());
                    objMatalan.ToDate = Convert.ToDateTime(dtMonth.Rows[i]["MEnd"].ToString());
                    objMatalan.Month = dtMonth.Rows[i]["Month"].ToString();
                    objMatalan.Location = dtStore.Rows[j]["LocationCode"].ToString();
                    objMatalan.Country = dtStore.Rows[j]["Country"].ToString();
                    objMatalan.UpdateProfitLossActualReport();
                }

            }

            //  update budgets

            objMatalan.ImportProfitLoseBudgets();

            objMatalan.Year = ddlYear.SelectedItem.Text;
            objMatalan.FromMonth = Convert.ToInt32(ddlMonth.SelectedItem.Value) + 1;
            objMatalan.ToMonth = 12;
            dtMonth = objMatalan.GetMonthDetails();

            for (int i = 0; i < dtMonth.Rows.Count; i++)
            {
                for (int j = 0; j < dtStore.Rows.Count; j++)
                {
                    objMatalan.FromDate = Convert.ToDateTime(dtMonth.Rows[i]["MStart"].ToString());
                    objMatalan.ToDate = Convert.ToDateTime(dtMonth.Rows[i]["MEnd"].ToString());
                    objMatalan.Month = dtMonth.Rows[i]["Month"].ToString();
                    objMatalan.Location = dtStore.Rows[j]["LocationCode"].ToString();

                    objMatalan.Year = ddlYear.SelectedItem.Value;
                    objMatalan.UpdateProfitLossBudgetReport();
                }
            }

            
        }
        #endregion MatalanTableUpdation

        #endregion Methods

        
    }
}