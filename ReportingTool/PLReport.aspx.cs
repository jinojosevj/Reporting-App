
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
    public partial class PLReport : System.Web.UI.Page
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
            GetStockDetails objMatalan = new GetStockDetails();
            // objTati.WeekNo = Convert.ToInt32(txtWeekNo.Text.Trim());
            objMatalan.Year = ddlYear.SelectedItem.Text;
            objMatalan.MonthNo = Convert.ToInt32(ddlMonth.SelectedItem.Value);

            bool Result = true;

            Result = objMatalan.InsertProfitAndLossReport();

            if (Result)
            {
                if (chkProfitLoss.Checked)
                    GeneratePLReport();
            }
            else
            {
                lblMessage.Text = "Report Failed !";
                lblMessage.ForeColor = System.Drawing.Color.Red;
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
            string fileName = ViewState["FileNamePL"].ToString();
            FileDownload(fileName);
        }
        #endregion btnProfitLoss_Click

        #endregion Events


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


            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportMatalan.xlsx";

            switch(ddlCountry.SelectedItem.Text)
            {
                case "All": fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportMatalan.xlsx";
                            break;
                case "Jordan":
                            fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportJordan.xlsx";
                            break;
                case "UAE":
                            fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportUAE.xlsx";
                            break;
                case "Oman":
                            fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportOman.xlsx";
                            break;
                case "Qatar":
                            fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportQatar.xlsx";
                            break;
                case "Bahrain":
                    fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportBahrain.xlsx";
                    break;
            }



            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            GetStockDetails objStock = new GetStockDetails();

            //objStock.ReportType = ddlType.SelectedItem.Value;

            DataTable dtStock = null;
            //DataTable dtMonth = null;

            if (txtLocation.Text.Trim().Length > 0)
            {
                string location = txtLocation.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\ProfitLossReportOneMatalan.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];

                objStock.Location = location;
                objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objStock.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());
                objStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());

                dtStock = objStock.GetProfitAndLoss();

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
                objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());
                objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objStock.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());
                objStock.BahRate= Convert.ToDecimal(txtBahrainRate.Text.Trim());


                if (ddlCountry.SelectedItem.Text =="All")
                {

                    Excel1.Worksheet xlSheetSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetSummary.Name = "Consolidated Summary";
                    objStock.Location = "Summary";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetSummary, "Summary");


                    Excel1.Worksheet xlSheetJordanSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheetJordanSummary.Name = "Jordan Summary";
                    objStock.Location = "Jordan";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetJordanSummary, "Jordan");


                    Excel1.Worksheet xlSheetUaeSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheetUaeSummary.Name = "UAE Summary";
                    objStock.Location = "UAE";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetUaeSummary, "UAE");

                    Excel1.Worksheet xlSheetOmanSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheetOmanSummary.Name = "Oman Summary";
                    objStock.Location = "Oman";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetOmanSummary, "Oman");

                    Excel1.Worksheet xlSheetQatarSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheetQatarSummary.Name = "Qatar Summary";
                    objStock.Location = "Qatar";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetQatarSummary, "Qatar");

                    Excel1.Worksheet xlSheetBahrainSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheetBahrainSummary.Name = "Bahrain Summary";
                    objStock.Location = "Bahrain";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetBahrainSummary, "Bahrain");

                    Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheet0408.Name = "DC";
                    objStock.Location = "DC";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0408, "DC");


                    Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                    xlSheet0400.Name = "0400";
                    objStock.Location = "0400";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0400, "0400");

                    Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                    xlSheet0401.Name = "0401";
                    objStock.Location = "0401";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0401, "0401");

                    Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                    xlSheet0402.Name = "0402";
                    objStock.Location = "0402";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0402, "0401");

                    Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                    xlSheet0403.Name = "0403";
                    objStock.Location = "0403";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0403, "0403");


                    Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                    xlSheet0404.Name = "0404";
                    objStock.Location = "0404";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0404, "0404");

                    Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                    xlSheet0405.Name = "0405";
                    objStock.Location = "0405";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0405, "0405");

                    Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                    xlSheet0406.Name = "0406";
                    objStock.Location = "0406";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0406, "0406");

                    Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[15];
                    xlSheet0407.Name = "0407";
                    objStock.Location = "0407";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0407, "0407");

                    Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[16];
                    xlSheet0409.Name = "0409";
                    objStock.Location = "0409";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0409, "0409");

                    Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[17];
                    xlSheet0410.Name = "0410";
                    objStock.Location = "0410";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0410, "0410");

                    Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[18];
                    xlSheet0411.Name = "0411";
                    objStock.Location = "0411";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0411, "0411");

                    Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[19];
                    xlSheet0412.Name = "0412";
                    objStock.Location = "0412";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0412, "0412");

                    Excel1.Worksheet xlSheetQHO = (Excel1.Worksheet)myExcelWorkbook.Sheets[20];
                    xlSheetQHO.Name = "QATARHO";
                    objStock.Location = "QATARHO";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetQHO, "QATARHO");

                    Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[21];
                    xlSheet0414.Name = "0414";
                    objStock.Location = "0414";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0414, "0414");

                    Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[22];
                    xlSheet0415.Name = "0415";
                    objStock.Location = "0415";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0415, "0415");

                    Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[23];
                    xlSheet0416.Name = "0416";
                    objStock.Location = "0416";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0416, "0416");


                    Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[24];
                    xlSheet0417.Name = "0417";
                    objStock.Location = "0417";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0417, "0417");

                    Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[25];
                    xlSheet0418.Name = "0418";
                    objStock.Location = "0418";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0418, "0418");

                    Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[26];
                    xlSheet0419.Name = "0419";
                    objStock.Location = "0419";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0419, "0419");


                    Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[27];
                    xlSheet0421.Name = "0421";
                    objStock.Location = "0421";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0421, "0421");

                    Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[28];
                    xlSheet0422.Name = "0422";
                    objStock.Location = "0422";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0422, "0422");

                    Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[29];
                    xlSheet0423.Name = "0423";
                    objStock.Location = "0423";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0423, "0423");


                    Excel1.Worksheet xlSheet0424 = (Excel1.Worksheet)myExcelWorkbook.Sheets[30];
                    xlSheet0424.Name = "0424";
                    objStock.Location = "0424";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0424, "0424");

                    Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[31];
                    xlSheet0425.Name = "0425";
                    objStock.Location = "0425";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0425, "0425");

                    Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[32];
                    xlSheet0426.Name = "0426";
                    objStock.Location = "0426";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0426, "0426");

                    Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[33];
                    xlSheet0427.Name = "0427";
                    objStock.Location = "0427";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0427, "0427");
                }
                else if(ddlCountry.SelectedItem.Text=="Jordan")
                {
                    Excel1.Worksheet xlSheetJordanSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetJordanSummary.Name = "Jordan Summary";
                    objStock.Location = "Jordan";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetJordanSummary, "Jordan");

                    Excel1.Worksheet xlSheet0400 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheet0400.Name = "0400";
                    objStock.Location = "0400";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0400, "0400");

                    Excel1.Worksheet xlSheet0404 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheet0404.Name = "0404";
                    objStock.Location = "0404";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0404, "0404");

                    Excel1.Worksheet xlSheet0411 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheet0411.Name = "0411";
                    objStock.Location = "0411";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0411, "0411");

                    Excel1.Worksheet xlSheet0422 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheet0422.Name = "0422";
                    objStock.Location = "0422";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0422, "0422");

                    Excel1.Worksheet xlSheet0423 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheet0423.Name = "0423";
                    objStock.Location = "0423";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0423, "0423");


                    Excel1.Worksheet xlSheet0425 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheet0425.Name = "0425";
                    objStock.Location = "0425";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0425, "0425");

                    Excel1.Worksheet xlSheet0426 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                    xlSheet0426.Name = "0426";
                    objStock.Location = "0426";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0426, "0426");


                }

                else if (ddlCountry.SelectedItem.Text == "UAE")
                {
                    Excel1.Worksheet xlSheetUaeSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetUaeSummary.Name = "UAE Summary";
                    objStock.Location = "UAE";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetUaeSummary, "UAE");


                    Excel1.Worksheet xlSheet0408 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheet0408.Name = "DC";
                    objStock.Location = "DC";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0408, "DC");
                    

                    Excel1.Worksheet xlSheet0401 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheet0401.Name = "0401";
                    objStock.Location = "0401";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0401, "0401");

                    Excel1.Worksheet xlSheet0402 = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheet0402.Name = "0402";
                    objStock.Location = "0402";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0402, "0401");

                    Excel1.Worksheet xlSheet0403 = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheet0403.Name = "0403";
                    objStock.Location = "0403";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0403, "0403");


                    Excel1.Worksheet xlSheet0405 = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheet0405.Name = "0405";
                    objStock.Location = "0405";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0405, "0405");

                    Excel1.Worksheet xlSheet0406 = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheet0406.Name = "0406";
                    objStock.Location = "0406";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0406, "0406");


                    Excel1.Worksheet xlSheet0409 = (Excel1.Worksheet)myExcelWorkbook.Sheets[8];
                    xlSheet0409.Name = "0409";
                    objStock.Location = "0409";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0409, "0409");

                    Excel1.Worksheet xlSheet0410 = (Excel1.Worksheet)myExcelWorkbook.Sheets[9];
                    xlSheet0410.Name = "0410";
                    objStock.Location = "0410";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0410, "0410");


                    Excel1.Worksheet xlSheet0414 = (Excel1.Worksheet)myExcelWorkbook.Sheets[10];
                    xlSheet0414.Name = "0414";
                    objStock.Location = "0414";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0414, "0414");

                    Excel1.Worksheet xlSheet0415 = (Excel1.Worksheet)myExcelWorkbook.Sheets[11];
                    xlSheet0415.Name = "0415";
                    objStock.Location = "0415";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0415, "0415");

                    
                    Excel1.Worksheet xlSheet0417 = (Excel1.Worksheet)myExcelWorkbook.Sheets[12];
                    xlSheet0417.Name = "0417";
                    objStock.Location = "0417";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0417, "0417");

                    Excel1.Worksheet xlSheet0418 = (Excel1.Worksheet)myExcelWorkbook.Sheets[13];
                    xlSheet0418.Name = "0418";
                    objStock.Location = "0418";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0418, "0418");

                    Excel1.Worksheet xlSheet0419 = (Excel1.Worksheet)myExcelWorkbook.Sheets[14];
                    xlSheet0419.Name = "0419";
                    objStock.Location = "0419";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0419, "0419");

                }
                else if (ddlCountry.SelectedItem.Text == "Oman")
                {
                    Excel1.Worksheet xlSheetOmanSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetOmanSummary.Name = "Oman Summary";
                    objStock.Location = "Oman";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetOmanSummary, "Oman");

                    Excel1.Worksheet xlSheet0407 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheet0407.Name = "0407";
                    objStock.Location = "0407";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0407, "0407");

                    Excel1.Worksheet xlSheet0421 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheet0421.Name = "0421";
                    objStock.Location = "0421";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0421, "0421");

                }
                else if (ddlCountry.SelectedItem.Text == "Qatar")
                {
                    Excel1.Worksheet xlSheetQatarSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetQatarSummary.Name = "Qatar Summary";
                    objStock.Location = "Qatar";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetQatarSummary, "Qatar");

                    Excel1.Worksheet xlSheet0412 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheet0412.Name = "0412";
                    objStock.Location = "0412";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0412, "0412");

                    Excel1.Worksheet xlSheetQHO = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheetQHO.Name = "QATARHO";
                    objStock.Location = "QATARHO";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetQHO, "QATARHO");

                }

                else if (ddlCountry.SelectedItem.Text == "Bahrain")
                {
                    Excel1.Worksheet xlSheetBahrainSummary = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetBahrainSummary.Name = "Bahrain Summary";
                    objStock.Location = "Bahrain";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheetBahrainSummary, "Bahrain");

                    Excel1.Worksheet xlSheet0416 = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheet0416.Name = "0416";
                    objStock.Location = "0416";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0416, "0416");

                    Excel1.Worksheet xlSheet0427 = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheet0427.Name = "0427";
                    objStock.Location = "0427";
                    dtStock = objStock.GetProfitAndLoss();
                    WriteToExcelPL(dtStock, xlSheet0427, "0427");

                }


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

            if (location == "Summary")
            {
                Rate = "QAR";
            }
            else if (location == "Jordan")
            {
                Rate = "JOR";
            }
            else if (location == "UAE")
            {
                Rate = "AED";
            }
            else if (location =="Oman")
            {
                Rate = "OMR";
            }
            else if (location == "0407" || location == "0421")
            {
                Rate = "OMR";
            }
            else if (location == "0400" || location == "0404" || location == "0411" || location == "0425" || location == "0426")
            {
                Rate = "JOR";
            }
            else if (location == "DC"||location == "0401" || location == "0402" || location == "0403" || location == "0405" || location == "0406" || location == "0409" || location == "0410" || location == "0414" || location == "0415" || location == "0417" || location == "0418" || location == "0419")
            {
                Rate = "AED";
            }
            else if (location == "0416" || location == "0427")
            {
                Rate = "BD";
            }
            else if (location == "0424")
            {
                Rate = "SAR";
            }

            // myExcelWorksheet.get_Range("A1", misValue).Formula = location;
            myExcelWorksheet.get_Range("C1", misValue).Formula = location + " - Matalan Profit And Loss Report For  " + ddlMonth.SelectedItem.Text.ToString() + " - " + ddlYear.SelectedItem.Value.ToString() + " (" + Rate + ")";
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
                else if(Value==4)
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

        
    }
}